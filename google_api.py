from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from google.oauth2 import service_account
from pprint import pprint
import string

class GoogleAPI:
    SERVICE_ACCOUNT_FILE = 'keys.json'      # From Google Services
    SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/spreadsheets.currentonly']

    # The ID and range of a sample spreadsheet.
    #SAMPLE_SPREADSHEET_ID = '1fMu5w9SzsBQEgX2SwYca5lzyHHPVPTowvkwLzj-2E4w'

    # Spreadsheet range <SheetName>!<cell1>:<cell2>
    def __init__(self, spreadsheet_id):
        self.spreadsheet_id = spreadsheet_id    # SpreadSheet ID
        self.credentials = service_account.Credentials.from_service_account_file(
            self.SERVICE_ACCOUNT_FILE, scopes=self.SCOPES)
        self.service = build('sheets', 'v4', credentials=self.credentials)
        self.sheet = self.service.spreadsheets()

    def read(self, range_):
        try:
            # Call the Sheets API
            result = self.sheet.values().get(spreadsheetId=self.spreadsheet_id,
                                        range=range_).execute()
            values = result.get('values', [])
            return result
            #print(result)
        except HttpError as err:
            print(err)

    def write(self, sheet_name=None, body=None):
        # body format [[*], [*]]
        self._format_sheet(sheet_name)
        if not sheet_name:
            raise Exception('sheet undefined')
        if not body:
            raise Exception('body undefined')
        print('Writing "{}"'.format(sheet_name))
        range_ = "{}!A1".format(sheet_name)
        request = self.sheet.values().update(
            spreadsheetId=self.spreadsheet_id, range=range_, valueInputOption="USER_ENTERED", body={"values": body})
        request.execute()

    def _format_sheet(self, sheet_name):
        print('Formatting "{}"'.format(sheet_name))
        try:
            self.delete_sheet(sheet_name)
        except Exception:
            pass
        try:
            self.add_sheet(sheet_name)
        except Exception:
            pass

    def _batch_update(self, request_body=None):
        # request body format {
        #   'requests': [],
        # }
        if not request_body:
            raise Exception('request_body is undefined')
        request = self.sheet.batchUpdate(spreadsheetId=self.spreadsheet_id, body=request_body)
        request.execute()

    def get_prop(self, sheet_name):
        ranges = '{}!A1'.format(sheet_name)
        include_grid_data = False
        request = self.sheet.get(spreadsheetId=self.spreadsheet_id, ranges=ranges, includeGridData=include_grid_data)
        response = request.execute()
        return response

    def add_sheet(self, sheet_name=None):
        #print('Adding "{}" Sheet'.format(sheet_name))
        add_sheet_request = {
            'requests': [
                {
                    'addSheet': {
                        'properties': {
                            'title': sheet_name
                        }
                    }
                }
            ]
        }
        self._batch_update(add_sheet_request)

    def delete_sheet(self, sheet_name=None):
        prop = self.get_prop(sheet_name)['sheets'][0]['properties']
        #print('Deleting "{}" Sheet'.format(sheet_name))
        sheet_id = prop['sheetId']
        delete_request = {
            'requests': [
                {
                    'deleteSheet': {
                        'sheetId': sheet_id
                    }
                }
            ]
        }
        self._batch_update(delete_request)

    def list_sheets(self):
        sheet_metadata = self.sheet.get(spreadsheetId=self.spreadsheet_id).execute()
        sheets = sheet_metadata.get('sheets', '')
        return sheets


def list_to_gsheet(sheet_name, spreadsheet_id=None, body=None):
    google = GoogleAPI(spreadsheet_id)
    google.write(sheet_name=sheet_name, body=body)


def color_cell(sheet_name, coordinate=None, cell=None, spreadsheet_id=None, red=1, green=1, blue=1, alpha=1):
    google = GoogleAPI(spreadsheet_id)
    prop = google.get_prop(sheet_name)['sheets'][0]['properties']
    dimension = tuple(prop['gridProperties'].values())
    sheet_id = prop['sheetId']
    if cell:
        coordinate = (string.ascii_uppercase.index(cell[0]), int(cell[1]) - 1)
    requests = [
        {
            "updateCells": {
                "rows": [
                    {
                        "values": [
                            {
                                "userEnteredFormat": {
                                    "backgroundColor": {
                                        "red": red,
                                        "green": green,
                                        "blue": blue,
                                        "alpha": alpha
                                    }
                                }
                            }
                        ]
                    }
                ],
                "fields": 'userEnteredFormat.backgroundColor',
                "range": {
                    "sheetId": sheet_id,
                    "startRowIndex": coordinate[1],
                    #"endRowIndex": dimension[0],
                    "startColumnIndex": coordinate[0],
                    #"endColumnIndex": dimension[1]
                }
            }
        }
    ]
    body = {
        'requests': requests
    }
    google._batch_update(body)

