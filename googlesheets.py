from google_api import list_to_gsheet
from google_api import GoogleAPI
import requests
import time
import random
import string


cols_ref = list(string.ascii_uppercase) + ['A' + upper for upper in string.ascii_uppercase] + ['B' + upper for upper in string.ascii_uppercase]

class GoogleSheets:
    """Utility Class for Energy Scrapers"""
    def __init__(self, spreadsheet_id):
        self.spreadsheet_id = spreadsheet_id
        self.google = GoogleAPI(self.spreadsheet_id)

    def get_tables(self, sheet_name, last_col):
        #sheet_name = 'input values'
        row = 2
        #last_col = 'C'
        fields_input = '{}!A{}:{}{}'.format(sheet_name, 1, last_col, 1)
        data_fields = self.google.read(fields_input)['values'][0]
        data_list = []
        print('Loading rows from {}'.format(sheet_name))
        while True:
            time.sleep(random.uniform(0.2, 0.25))
            sheet_input = '{}!A{}:{}{}'.format(sheet_name, row, last_col, row)
            row_data = self.google.read(sheet_input)
            row_dict = {}
            if 'values' in row_data:
                row_values = row_data.get('values')[0]
                for i, key in enumerate(data_fields):
                    row_dict[key] = row_values[i]
                data_list.append(row_dict)
            else:
                break
            #print(str(row_dict).replace('{', '').replace('}', ''))
            print('\rRow: {}'.format(row), end='')
            row += 1
        #print()
        return data_list

    def get_full_tables(self, sheet_name, last_col='Z'):
        last_row = 10000
        fields_input = '{}!A{}:{}{}'.format(sheet_name, 1, last_col, 1)
        data_fields = self.google.read(fields_input)['values'][0]
        last_col = cols_ref[len(data_fields) - 1]
        data_list = []
        row = 2
        sheet_input = '{}!A{}:{}{}'.format(sheet_name, row, last_col, last_row)
        dataset = self.google.read(sheet_input)
        try:
            for row_values in dataset['values']:
                row_dict = {}
                for i, key in enumerate(data_fields):
                    try:
                        row_dict[key] = row_values[i]
                    except:
                        row_dict[key] = ''
                data_list.append(row_dict)
            return data_list
        except:
            return []

    def save_gsheet(self, sheet_name, list_of_dict, fieldnames):
        unique_keys = []
        for lod in list_of_dict:
            for key in lod.keys():
                if key not in unique_keys:
                    unique_keys.append(key)
        for lod in list_of_dict:
            for key in unique_keys:
                if key not in lod:
                    lod[key] = ''
        list_of_list = []
        #list_of_list.append(list(list_of_dict[0].keys()))
        list_of_list = [fieldnames]
        for lod in list_of_dict:
            row = []
            for fieldname in fieldnames:
                row.append(lod[fieldname])
            list_of_list.append(row)
        list_to_gsheet(sheet_name, self.spreadsheet_id, list_of_list)
