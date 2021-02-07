from datetime import datetime
import calendar
import locale
locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials
from docxtpl import DocxTemplate

from pprint import pprint

def generate_docx_from_schedule(context:int, num:int):
    template = DocxTemplate('template.docx')      
    for name in context:
        temp = context[name]['name_discipline']
        name_lenght = temp.find('/')
        template.render(context[num])
        template.save("docx/%s.docx" %str(context[name]['name_discipline'][:name_lenght]))


## connect API google sheets
CREDENTIALS_FILE = 'annotation2-c46cdd2a61d8.json' # file with project data
spreadsheetId  = '1SEz5_QVrWDzSb6g60KlwxShdmPGHSnq3EBw5zJ8k-jQ' #token google sheet

credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, 
                ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'])

httpAuth = credentials.authorize(httplib2.Http()) 
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

#get rows from google sheet
range_name = 'sheet!A2:Q3'
rows = service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=range_name).execute()

data_sheets = []                 
for row in rows['values']:
    data_sheets.append(row)

#future keys
keys = ['index','name_discipline','description','block',
                'course_och', 'semester_och','form_educational', 'credit_hours', 
               'academic_hours', 'countact_hours', 'credit_hours_z', 'academic_hours_z',
                'countact_hours_z']

# match two list
dict_ = {}
for i, value in enumerate(data_sheets):
    new_match_list = dict(zip(keys, value))             
    dict_.setdefault(i, new_match_list)

#start function
for i in range(0, len(dict_)):
    generate_docx_from_schedule(dict_, i)