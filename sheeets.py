import locale
locale.setlocale(locale.LC_TIME, 'ru_RU.UTF-8')
import os
import httplib2
import apiclient.discovery
from oauth2client.service_account import ServiceAccountCredentials

from docxtpl import DocxTemplate
from docx import Document
from docx.shared import Inches
from docx.shared import Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH



def generate_docx_from_schedule(context:dict):
    template = DocxTemplate('template.docx')
    try:
        os.mkdir('annotations')
    except OSError as e:
        print('Folder is found!')
    #checking if there is a file 
    for notempty in context:
        temp = context[notempty]['name_discipline']
        index = context[notempty]['index']
        name_lenght = temp.find('/')
        if os.path.exists('annotations/{}.docx'.format(index+'_'+temp[:name_lenght])):
            os.remove('annotations/{}.docx'.format(index+'_'+temp[:name_lenght]))
    #generate docx in dir annotations
    for name in context:
        temp = context[name]['name_discipline']
        index = context[name]['index']
        name_lenght = temp.find('/')
        for num in range(len(context)):
            try:
                template.render(context[num])
                template.save("annotations/%s.docx" %str(index+'_'+temp[:name_lenght]))
            except:
                print("Oops!")



## connect API google sheets
CREDENTIALS_FILE = 'annotation2-c46cdd2a61d8.json' # file with project data
spreadsheetId  = '1SEz5_QVrWDzSb6g60KlwxShdmPGHSnq3EBw5zJ8k-jQ' #token google sheet

credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, 
                ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'])

httpAuth = credentials.authorize(httplib2.Http()) 
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)

#range rows in google sheet
range_name = 'sheet!A2:AF3'

# clear cells 
data_sheets = []                 
for row in service.spreadsheets().values().get(spreadsheetId=spreadsheetId, range=range_name).execute()['values']:
    temp = []
    for cell in range(len(row)): 
        new_cell = row[cell].replace(row[cell],row[cell].strip())     
        temp.append(new_cell)
        if cell == len(row)-1:
            data_sheets.append(temp)

#future keys
keys = ['index','name_discipline','description','block',
                'course_och', 'semester_och', 'course_z', 'form_educational', 'credit_hours', 
               'academic_hours', 'countact_hours', 'credit_hours_z', 'academic_hours_z',
                'countact_hours_z', 'topic1', 'content_topic1', 'competence1', 
                'topic2', 'content_topic2', 'competence2', 'topic3', 'content_topic3', 'competence3',
                'topic4', 'content_topic4', 'competence4', 'topic5', 'content_topic5', 'competence5',
                'topic6', 'content_topic6', 'competence6']

# match two list
dict_ = {}
for i, value in enumerate(data_sheets):
    new_match_list = dict(zip(keys, value))              
    dict_.setdefault(i, new_match_list)

## add new parametors 
for index in dict_:
    if 'В.ДВ.' in dict_[index]['index']:       
        choice_type = {'choice_type':'части, формируемой участниками образовательных отношений'} 
        type_discipline = {'type_discipline':'по выбору'} 
        dict_[index].update(choice_type)
        dict_[index].update(type_discipline)
    else:
        choice_type = {'choice_type':'обязательной части'} 
        type_discipline = {'type_discipline':'обязательной'} 
        dict_[index].update(choice_type)
        dict_[index].update(type_discipline)

document = Document('template.docx')
for id in dict_:
    long_topics = int((len(dict_[id])-16)/3)
    table = document.tables[0]
    for i in range(long_topics):
        new_row = table.add_row()
        text = table.rows[i+1].cells
        row = table.rows[i+1]
        
        paragraph = row.cells[0].add_paragraph()
        # first column 
        first_column = paragraph.add_run(f'{str(i+1)}.')
        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        first_column.bold = True
        first_column.font.name = 'Calibri'
        first_column.font.size = Pt(10)      

        # second_column 
        paragraph = row.cells[1].add_paragraph()
        second_column = paragraph.add_run(dict_[id][f'topic{i+1}'])   
        second_column = 'Calibri'
        second_column = Pt(10)
        
        # third column
        paragraph = row.cells[2].add_paragraph()
        third_column = paragraph.add_run(dict_[id][f'topic{i+1}'])   
        third_column = 'Calibri'
        third_column = Pt(10)

        # fourth column
        paragraph = row.cells[3].add_paragraph()
        fourth_colmn = paragraph.add_run(dict_[id][f'topic{i+1}'])
        fourth_colmn = 'Calibri'
        fourth_colmn = Pt(10)

        
        text[1].text = dict_[id][f'topic{i+1}']
        text[2].text = dict_[id][f'content_topic{i+1}']
        text[3].text = dict_[id][f'competence{i+1}']

        

    
    document.save('demo.docx')
    



if __name__ == '_main':
    #start function
    generate_docx_from_schedule(dict_)