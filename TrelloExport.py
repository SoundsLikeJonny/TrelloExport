# TrelloExport.py
#
# Created by: Jon Evans
# Created on: April 21, 2020
# Original script written in AppleScript and can be found on my github
# https://github.com/SoundsLikeJonny

"""
    This script pulls information from Trello cards using the Trello.com API, formats it into a
    spreadsheet and emails that to management, updating Trello for the next workday.

    #####   Replace **** throughout the file with your info   #####

    create_spreadsheet_row() is the function responsible for creating a list of the card info.

    create_spreadsheet_nested_list() creates a nested list, where each list of the parent list represents a row.

    create_spreadsheet() creates a workbook and worksheet using the xlsxwriter module. It takes the nested list,
    and iterates through the lists, writing each cell to the worksheet.

    create_mime_message() will create the email message, and attach the excel sheet to the email

    email_file() creates an SMTP session with the SMTP host and port, your login info, and sends the MIME message.

    update_trello_board() will tell send a PUT request to tell the Trello API to archive the COMPLETE list,
    then sends a POST request to create a new Trello list called COMPLETE yymmdd where yymmdd is the date of
    the next work day (i.e. if today is Friday, March 31st, the next work day is Monday, April 3rd. So the final
    list name is something close to 'COMPLETE 200403')
"""


import datetime
import re
import requests
import json
from pytz import timezone
import xlsxwriter
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders

# Set the global variables
API_KEY = '****'
USER_TOKEN = '****'
AUTH = f'?key={API_KEY}&token={USER_TOKEN}'
AUTH_AND = f'&key={API_KEY}&token={USER_TOKEN}'
BOARD_ID = '****'
BOARD_URL = 'https://trello.com/1/boards/'
ALL_CARD_URL = f'{BOARD_URL}{BOARD_ID}/cards{AUTH}'
MEMBERS_URL = f'{BOARD_URL}{BOARD_ID}/members{AUTH}'
LISTS_URL = f'{BOARD_URL}{BOARD_ID}/lists{AUTH}'
CARD_URL = 'https://trello.com/1/cards/'
CUSTOM_FIELDS_URL = f'{BOARD_URL}{BOARD_ID}/customFields{AUTH}'
CUSTOM_FIELD_ITEMS_URL = '/customFieldItems'
ACTIONS_URL = '/actions?filter=all'
WEEKDAY_FORMAT = '%A'
DATING_FORMAT = '%y%m%d'
TIMEZONE = 'US/Pacific'
DATE_TIME_FORMAT_UTC = '%Y-%m-%d %H:%M:%S.%f'
DATE_TIME_FORMAT_LOCAL = '%Y-%m-%d %H:%M'
WEEKDAY_CHECK = 'Friday'

SPREADSHEET_ROW_1 = ['MODIFIED DATE', 'TYPE', 'TITLE', 'STATUS', 'WORKED ON BY', 'BACKLOG DATE',
                     'APPROVED DATE', 'EST. # OF REVISIONS', 'COMPLETED DATE', 'INFO', 'NOTES', 'URL']

# JSON Key strings
JS_NAME = 'name'
JS_ID = 'id'
JS_LAST_ACTIVITY = 'dateLastActivity'
JS_DESC = 'desc'
JS_LIST = 'idList'
JS_LABELS = 'labels'
JS_URL = 'shortUrl'
JS_MEMBERS = 'idMembers'
JS_FULLNAME = 'fullName'
JS_VALUE = 'value'
JS_TEXT = 'text'
JS_ID_CUSTOM_FIELD = 'idCustomField'
JS_CREATED_CARD = 'createCard'
JS_TYPE = 'type'
JS_DATE = 'date'
JS_LIST_BEFORE = 'listBefore'
JS_LIST_AFTER = 'listAfter'
JS_ACTION_UPDATE = 'updateCard'
JS_DATA = 'data'

LIST_BACKLOG = 'BACKLOG'
LIST_APPROVED = 'APPROVED'
LIST_COMPLETE = 'COMPLETE'


def get_date() -> list:
    """
    Gets the current date and the date of the next workday.
    :return: Two datetime objects in a list.
    """
    today_weekday_str = datetime.date.today().strftime(WEEKDAY_FORMAT)
    today_date = datetime.date.today()

    if today_weekday_str == WEEKDAY_CHECK:
        next_workday_date = today_date + datetime.timedelta(days=3)
    else:
        next_workday_date = today_date + datetime.timedelta(days=1)

    return [today_date, next_workday_date]


def get_all_cards() -> list:
    """
    Calls the Trello API and gets all the active cards on the board.
    :return: A list of all the card dictionaries.
    """
    api_result = requests.get(ALL_CARD_URL).content
    cards_json = json.loads(api_result)

    return cards_json


def get_all_members() -> dict:
    """
    Gets all of the members on the Trello board.
    :return: A dictionary of member ID's and Names.
    """
    api_result = requests.get(MEMBERS_URL).content
    members_json = json.loads(api_result)

    members_dict = {}
    for member in members_json:
        members_dict[member[JS_ID]] = member[JS_FULLNAME].split()[0]

    return members_dict


def format_time_utc_to_local(time_str: str) -> datetime:
    """
    Takes the Trello version of UTC time, removes the T and Z, and formats for pacific time.
    :param time_str: The Trello UTC value.
    :return: The Trello time converted to pacific time.
    """

    regex_data = re.compile(
        r'(\d\d\d\d-\d\d-\d\d)+'  # DATE
        r'T?'
        r'(\d\d:\d\d:\d\d.\d\d\d)+'  # Time
        r'Z?',
        re.VERBOSE)

    regex_timestamp = re.findall(regex_data, time_str)[0]
    regex_timestamp = ' '.join(regex_timestamp).strip()

    datetime_obj_utc = datetime.datetime.strptime(regex_timestamp, DATE_TIME_FORMAT_UTC)
    datetime_obj_pacific = datetime.datetime.astimezone(timezone(TIMEZONE).fromutc(datetime_obj_utc))

    return datetime_obj_pacific


def get_all_trello_lists() -> dict:
    """
    Calls the Trello API and gets all the list ID's and names.
    :return: A dictionary of dictionary keys and values of all the custom fields.
    """
    api_result = requests.get(LISTS_URL).content
    t_lists_json = json.loads(api_result)

    list_dict = {}
    for t_list in t_lists_json:
        list_dict[t_list[JS_ID]] = t_list[JS_NAME]

    return list_dict


def get_custom_field_names() -> list:
    """
    Calls the Trello API and gets all the custom field ID's and names.
    :return: A list of dictionary keys and values of all the custom fields.
    """
    api_result = requests.get(CUSTOM_FIELDS_URL).content
    custom_fields_json = json.loads(api_result)

    custom_fields_list = []
    for field in custom_fields_json:
        custom_fields_list.append({JS_ID: field[JS_ID], JS_NAME: field[JS_NAME]})

    return custom_fields_list


def get_custom_field_values(card_id: str) -> dict:
    """
    Get the custom field values.
    :param card_id: The ID of the Trello card.
    :return: Dictionary of all custom field values.
    """
    api_result = requests.get(f'{CARD_URL}'
                              f'{card_id}'
                              f'{CUSTOM_FIELD_ITEMS_URL}'
                              f'{AUTH}').content

    custom_fields_list = json.loads(api_result)

    custom_fields_dict = {}
    for field in custom_fields_list:
        custom_fields_dict[field[JS_ID_CUSTOM_FIELD]] = field[JS_VALUE][JS_TEXT]

    return custom_fields_dict


def get_custom_fields(card_fields: dict, field_names: list) -> str:
    """
    Loops through each custom filed and gets the value.
    :param card_fields: All the used fields on the current card.
    :param field_names: All the possible custom fields.
    :return: A formatted string with the custom field names and values.
    """
    custom_fields = []
    for field in field_names:
        custom_fields.append(
            field[JS_NAME].upper()
            + ': '
            + (card_fields[field[JS_ID]].upper() if field[JS_ID] in card_fields else '')
            + '\n')

    return ''.join(custom_fields)


def get_action_date_time(card_action: dict) -> datetime:
    """
    Gets the date the action took place.
    :param card_action: The card action.
    :return: The date the action took place.
    """
    return format_time_utc_to_local(card_action[JS_DATE])


def get_action_type(card_action: dict) -> str:
    """
    Gets the type of action the card took. e.g. updateCard.
    :param card_action: The card action.
    :return: The action type.
    """
    return card_action[JS_TYPE]


def get_action_list_before(card_action: dict) -> str:
    """
    Gets the name of the list the card used to be in.
    :param card_action: The action of the card containing the list change info.
    :return: The name of the list.
    """
    action_list_before = ''
    if JS_LIST_BEFORE in card_action[JS_DATA]:
        action_list_before = card_action[JS_DATA][JS_LIST_BEFORE][JS_NAME]

    return action_list_before


def get_action_list_after(card_action: dict) -> str:
    """
    Gets the name of the list the card has been placed into
    :param card_action: The action of the card containing the list change info
    :return: The name of the list
    """
    action_list_after = ''
    if JS_LIST_AFTER in card_action[JS_DATA]:
        action_list_after = card_action[JS_DATA][JS_LIST_AFTER][JS_NAME]

    return action_list_after


def get_trello_card_actions_json(card_id: str) -> list:
    """
    Calls the Trello API and gets all the actions on a card.
    :param card_id: The ID of the card.
    :return: A list of dictionaries of all the actions on the card.
    """
    api_result = requests.get(f'{CARD_URL}'
                              f'{card_id}'
                              f'{ACTIONS_URL}'
                              f'{AUTH_AND}').content

    card_actions = json.loads(api_result)

    return card_actions


def create_spreadsheet_nested_list() -> list:
    """
    Creates the nested list of all the values to be used in the spreadsheet.
    :return: A nested list.
    """
    all_cards_list = get_all_cards()
    all_custom_field_names = get_custom_field_names()

    spreadsheet_row_list = [SPREADSHEET_ROW_1]

    for card in all_cards_list:
        # Get all custom field IDs and values of the card as a dict
        custom_field_values_dict = get_custom_field_values(card[JS_ID])
        # Get all card action field IDs and values as a dict
        card_actions_json = get_trello_card_actions_json(card[JS_ID])
        # Put a new list of data into the main list
        spreadsheet_row_list += [create_spreadsheet_row(card, custom_field_values_dict,
                                                        all_custom_field_names, card_actions_json)]
        print(f'Getting info for: {card[JS_NAME]}')

    return spreadsheet_row_list


def create_spreadsheet_row(card: dict, custom_field_values_dict: dict,
                           all_custom_field_names: list, card_actions: list) -> list:
    """
    Creates a list containing all the info for a row of the spreadsheet.
    :param card: Trello card JSON info.
    :param custom_field_values_dict: A dictionary of all the custom field values of the card.
    :param all_custom_field_names: A dictionary of all the custom field names of the card.
    :param card_actions: A list of all the actions on the card.
    :return: A list of all the info for the row of the spreadsheet.
    """
    row = [get_card_last_activity(card).upper(),
           get_card_label(card).upper(),
           get_card_name(card).upper(),
           get_card_current_list(card).upper(),
           get_card_members(card).upper(),
           get_backlog_start_date(card_actions).upper(),
           get_date_approved(card_actions).upper(),
           get_no_of_revisions(card_actions).upper(),
           get_date_completed(card_actions).upper(),
           get_custom_fields(custom_field_values_dict, all_custom_field_names).upper(),
           get_card_description(card).upper(),
           get_card_url(card).upper()]

    return row


def create_spreadsheet(spreadsheet_list_nested: list, today_date: str, title: str) -> str:
    """
    Creates the spreadsheet with all the Trello board information.
    :param spreadsheet_list_nested: A nested list of all the cell values for the spreadsheet.
    :param today_date: Today's date.
    :param title: The title for the spreadsheet.
    :return: The filename of the spreadsheet.
    """
    filename = f'{title}.xlsx'
    workbook = xlsxwriter.Workbook(filename)

    worksheet = workbook.add_worksheet(today_date)

    cell_format = workbook.add_format()
    cell_format.set_align('center')
    cell_format.set_align('vcenter')

    row_index = 0
    for row in spreadsheet_list_nested:

        col_index = 0
        for column in row:

            worksheet.write(row_index, col_index, column, cell_format)
            col_index += 1

        row_index += 1

    workbook.close()
    return filename


def get_backlog_start_date(actions: list) -> str:
    """
    Gets the date the card was placed in the backlog. Defaults to the creation date.
    :param actions: A list of dictionaries of all the actions on the card.
    :return: The date the card was backlogged.
    """
    # Default to the creation date
    start_date = get_card_creation_date(actions)
    for action in actions:
        # If the type of action is an updateCard type and the list it used to belong to was the backlog list
        if get_action_type(action) == JS_ACTION_UPDATE and get_action_list_before(action) == LIST_BACKLOG:
            # If the start date from the action in the last iteration is less than this iteration
            if start_date < (action_time := format_time_utc_to_local(action[JS_DATE])):
                start_date = action_time

    return start_date.strftime(DATE_TIME_FORMAT_LOCAL)


def get_card_creation_date(actions: list) -> datetime:
    """
    Gets the creation date of the card.
    :param actions: A list of dictionaries of all the actions on the card.
    :return: A formatted date of when the card was created.
    """
    return format_time_utc_to_local(actions[-1][JS_DATE])


def get_date_approved(actions: list) -> str:
    """
    Checks the card's actions to determine if the card was approved and gets the date.
    :param actions: A list of dictionaries of all the actions on the card.
    :return: The date the card was approved, or the card creation date if not applicable.
    """
    approved_date = get_card_creation_date(actions)
    for action in actions:
        # If the type of action is an updateCard type and the list it now belongs to is the approved list
        if get_action_type(action) == JS_ACTION_UPDATE and get_action_list_after(action) == LIST_APPROVED:
            # If the start date from the action in the last iteration is less than this iteration
            if approved_date < (action_time := format_time_utc_to_local(action[JS_DATE])):
                approved_date = action_time

    return approved_date.strftime(DATE_TIME_FORMAT_LOCAL)


def get_no_of_revisions(actions: list) -> str:
    """
    Gets the number of times a card moved back from one list to another.
    :param actions: A list of dictionaries of all the actions on the card.
    :return: The number of revisions a card had.
    """
    revisions = 0
    for action in actions:
        # All of the Trello lists in order from left to right
        ordered_lists = get_ordered_lists(action)

        # If the type of action is an updateCard type
        if get_action_type(action) == JS_ACTION_UPDATE:

            # The current list
            list_after_action = get_action_list_after(action)

            # The list the card used to belong to
            list_before_action = get_action_list_before(action)

            # Check if these lists still exist on the board
            if list_before_action in ordered_lists and list_after_action in ordered_lists:

                # If the card was moved back from its current list to any previous list
                if ordered_lists.index(list_after_action) < ordered_lists.index(list_before_action):
                    revisions += 1

    return str(revisions)


def get_ordered_lists(action: dict) -> list:
    """
    Gets all Trello lists in the order they're supposed to be in.
    Uses a regex expression to determine the COMPLETE list which has the date formatted as yymmdd.
    :param action: A specific action on a card. Used to find the full name of the COMPLETE list
    :return: The full list of Trello lists in order from left to right.
    """
    regex_pattern = re.compile(r'(COMPLETE \d{6})')
    regex_find = re.findall(regex_pattern, get_action_list_after(action))
    complete = regex_string if LIST_COMPLETE in (regex_string := ''.join(regex_find)) else ''

    lists_left_to_right = ['BACKLOG',
                           'IN PROGRESS',
                           'WAITING FOR APPROVAL',
                           'APPROVED',
                           'IMPLEMENTING',
                            complete]

    return lists_left_to_right


def get_date_completed(actions: list) -> str:
    """
    Determines if a card has been moved into the complete list and returns the date if true.
    :param actions: A list of dictionaries of all the actions on the card.
    :return: The date completed or empty string if not completed.
    """
    regex_pattern = re.compile(r'(COMPLETE) \d{6}')

    completed_date = ''
    for action in actions:
        regex_find = re.findall(regex_pattern, get_action_list_after(action))

        if get_action_type(action) == JS_ACTION_UPDATE and ''.join(regex_find) == LIST_COMPLETE:
            completed_date = format_time_utc_to_local(action[JS_DATE]).strftime(DATE_TIME_FORMAT_LOCAL)

    return completed_date


def get_card_last_activity(card: dict) -> str:
    """
    Gets the date of the card's last activity.
    :param card: Trello card JSON info.
    :return: The date as string.
    """
    return format_time_utc_to_local(card[JS_LAST_ACTIVITY]).strftime(DATE_TIME_FORMAT_LOCAL)


def get_card_label(card: dict) -> str:
    """
    Gets all labels on the card.
    :param card: Trello card JSON info.
    :return: All labels on the card.
    """
    labels = [label[JS_NAME] for label in card[JS_LABELS]]
    return ', '.join(labels)


def get_card_name(card: dict) -> str:
    """
    Gets the name of the Trello card.
    :param card: Trello card JSON info.
    :return: Card name.
    """
    return card[JS_NAME]


def get_card_current_list(card: dict) -> str:
    """
    Gets the current Trello list the card belongs to.
    :param card: Trello card JSON info.
    :return: The name of the list.
    """
    lists = get_all_trello_lists()
    return lists[card[JS_LIST]]


def get_card_members(card: dict) -> str:
    """
    Compares the id's of all the Trello board members on the card and gets all the members on the card.
    :param card: Trello card JSON info.
    :return: Members on the card.
    """
    members_dict = get_all_members()

    all_members = [members_dict[member] for member in card[JS_MEMBERS]]
    return ', '.join(all_members)


def get_card_url(card: dict) -> str:
    """
    Gets the short URL of a Trello card.
    :param card: Trello card JSON info.
    :return: URL of the Trello card.
    """
    return card[JS_URL]


def get_card_description(card: dict) -> str:
    """
    Gets the first 100 characters of a description on a card.
    :param card: Trello card JSON info.
    :return: 100 characters of the description.
    """
    return card[JS_DESC][:100]


def get_complete_list_id() -> str:
    """
    Searches through all Trello lists to find the one with COMPLETE and gets the ID of that list.
    :return: Trello list ID as string.
    """
    all_lists = get_all_trello_lists()
    regex_pattern = re.compile(r'(COMPLETE) \d{6}')

    for trello_list in all_lists:

        regex_find = re.findall(regex_pattern, all_lists[trello_list])

        if ''.join(regex_find) == 'COMPLETE':
            return trello_list


def sort_spreadsheet_by_date(spreadsheet: list) -> None:
    """
    Sort's the spreadsheet nested list according to the first column of each row.
    :param spreadsheet: A nested list containing all the cell values for the spreadsheet as a nested list.
    """
    spreadsheet.sort(reverse=True, key=lambda x: x[0])


def email_file(file_path: str, today_date: str) -> None:
    """
    Creates the smtp session and emails the MIME formatted message.
    :param file_path: The file path of the spreadsheet.
    :param today_date: Today's date.
    """
    smtp_session = smtplib.SMTP(host='****', port=****)
    smtp_session.starttls()
    smtp_session.login(user='****', password='****')
    
    message = create_mime_message(file_path, today_date)
    smtp_session.send_message(message)


def create_mime_message(file_path: str, today_date: str) -> MIMEMultipart:
    """
    Creates the MIME message, encoding and attaching the file to the message.
    :param file_path: The file path of the spreadsheet.
    :param today_date: Today's date as a string.
    :return: The message as a MIME object.
    """
    message = MIMEMultipart()

    message['From'] = '****'
    message['To'] = '****'
    message['Subject'] = f'{today_date} Trello Log'
    message.attach(MIMEText('AUTOMATED EMAIL\n\nToday\'s Trello Log', 'plain'))

    # Open file with read bytes
    file = open(file_path, 'rb')

    # Encode the file
    mime_base_obj = MIMEBase('application', 'octet-stream')
    mime_base_obj.set_payload(file.read())
    encoders.encode_base64(mime_base_obj)

    mime_base_obj.add_header('Content-Disposition', f'attachment; filename={file_path}')

    message.attach(mime_base_obj)

    return message


def update_trello_board(date_next_workday: str):
    """
    Archives the today's Trello Complete list and create a new one for the next work day.
    :param date_next_workday: the date of the next workday, Monday-Friday only.
    """
    complete_list_id = get_complete_list_id()

    requests.put(f'https://api.trello.com/1/lists/'
                 f'{complete_list_id}'
                 f'/closed?value=true'
                 f'{AUTH_AND}')

    requests.post(f'https://api.trello.com/1/lists?name=COMPLETE%20'
                  f'{date_next_workday}'
                  f'&idBoard='
                  f'{BOARD_ID}' 
                  f'&pos=bottom'
                  f'{AUTH_AND}')


def main() -> None:

    today_date, next_workday_date = get_date()

    spreadsheet_nested_list = create_spreadsheet_nested_list()

    sort_spreadsheet_by_date(spreadsheet_nested_list)

    spreadsheet_filepath = create_spreadsheet(spreadsheet_nested_list, today_date.strftime(DATING_FORMAT),
                                              f'{today_date.strftime(DATING_FORMAT)} Trello Log')

    email_file(spreadsheet_filepath, today_date.strftime(DATING_FORMAT))

    update_trello_board(next_workday_date.strftime(DATING_FORMAT))

    print('Email sent, Trello list updated')


if __name__ == '__main__':
    main()