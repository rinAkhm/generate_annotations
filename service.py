import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

import os


class SetupApp:
    def __init__(self):
        self.sheet = 'thesis'
        self.token_page = '1SEz5_QVrWDzSb6g60KlwxShdmPGHSnq3EBw5zJ8k-jQ'
        self.credentials_json = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'cred', 'token.json')

    def get_data(self):
        credentials = ServiceAccountCredentials.from_json_keyfile_name(self.credentials_json,
                                                                       ['https://www.googleapis.com/auth/spreadsheets',
                                                                        'https://www.googleapis.com/auth/drive'])

        httpAuth = credentials.authorize(httplib2.Http())
        service = apiclient.discovery.build('sheets', 'v4', http=httpAuth)
        range_name = f'{self.sheet}!A3:J100'
        sheet = service.spreadsheets().values().get(spreadsheetId=self.token_page, range=range_name).execute().get(
            'values',
            [])
        return sheet
