# TrelloExport
This script pulls information from Trello cards using the Trello.com API, formats it into aspreadsheet and emails that to management, updating Trello for the next workday.

## Trello
In some companies, Trello is used to manage day-to-day tasks between departments. In Trello, each task is represented by a Trello 'card'. Each card is organized into lists relating to their current status. Everything is contained in a Trello board. A user of the board is called a'member', and each member can create, move, and modify cards amongst the lists. An example of lists for a board are:

- 'BACKLOG'
- 'IN PROGRESS'
- 'WAITING FOR APPROVAL'
- 'APPROVED'
- 'IMPLEMENTING'
- 'COMPLETE yymmdd'

Where COMPLETE is followed by today's date 24hr local time formatted as yymmdd.

Each card contains information about it's name, when it was created, when it was last modified, the current list the card belongs to, a history of 'actions' (which show who modified what on the card and when), members on the card, the card description and a lot more.

## Trello API
Calling the API using the requests.get() function with the argument 'https://trello.com/1/boards/{BOARD_ID id}/cards?key={APP_KEY}&token={USER_TOKEN}' will return a JSON of all the cards on the board. We can then get a python dictionary using json.loads() – and a series of utility functions that look like "def get_card_name(card: dict) -> str:" – to get all the card individual info.
    
## Notable Functions
- create_spreadsheet_row() is the function responsible for creating a list of the card info.

- create_spreadsheet_nested_list() creates a nested list, where each list of the parent list represents a row.

- create_spreadsheet() creates a workbook and worksheet using the xlsxwriter module. It takes the nested list, and iterates through the lists, writing each cell to the worksheet.

- create_mime_message() will create the email message, and attach the excel sheet to the email

- email_file() creates an SMTP session with the SMTP host and port, your login info, and sends the MIME message.

update_trello_board() will tell send a PUT request to tell the Trello API to archive the COMPLETE list,
then sends a POST request to create a new Trello list called COMPLETE yymmdd where yymmdd is the date of
the next work day (i.e. if today is Friday, March 31st, the next work day is Monday, April 3rd. So the final
list name is something close to 'COMPLETE 200403')

