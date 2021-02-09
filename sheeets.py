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



def create_table(dict_:dict, name_docx):
    document = Document(f'template.docx')

    long_topics = int((len(dict_)-16)/3)
    table = document.tables[0]
    for i in range(long_topics):
        new_row = table.add_row()
            # first column 
        paragraph = table.rows[i+1].cells[0].add_paragraph()
        first_column = paragraph.add_run(f'{str(i+1)}.')
        paragraph.paragraph_format.alignment = WD_ALIGN_PARAGRAPH.CENTER
        first_column.bold = True
        first_column.font.name = 'Calibri'
        first_column.font.size = Pt(10)      

            # second_column
        cell = table.rows[i+1].cells[1]
        cell.text = f'{{{{topic{i+1}}}}}'   #dict_[id][f'topic{i+1}']
        format_cell = cell.paragraphs[0].runs[0]
        format_cell.font.name = 'Calibri'
        format_cell.font.size = Pt(10)
            
            # third column
        cell = table.rows[i+1].cells[2]
        cell.text = f'{{{{content_topic{i+1}}}}}' #dict_[id][f'{{{{content_topic{i+1}}}}}']
        format_cell = cell.paragraphs[0].runs[0]  
        format_cell.font.name = 'Calibri'
        format_cell.font.size = Pt(10)

            # fourth_column
        cell = table.rows[i+1].cells[3]
        cell.text = f'{{{{competence{i+1}}}}}' # dict_[id][f'{{{{competence{i+1}}}}}']
        format_cell = cell.paragraphs[0].runs[0]  
        format_cell.font.name = 'Calibri'
        format_cell.font.size = Pt(10)

    document.save(f'annotations/{name_docx}.docx')       

def generate_docx_from_schedule(context:dict, lenght:int, count:int):
    #generate docx in dir annotations
    try:
        os.mkdir('annotations')
        print('Folder is create!')
    except OSError as e:

    # #checking if there is a file 
    # for notempty in context:
    #     temp = context[notempty]['name_discipline']
    #     index = context[notempty]['index']
    #     name_lenght = temp.find('/')
    #     if os.path.exists('annotations/{}.docx'.format(index+'_'+temp[:name_lenght])):
    #         os.remove('annotations/{}.docx'.format(index+'_'+temp[:name_lenght]))
 
        temp = context['name_discipline']
        index = context['index']
        name_lenght = temp.find('/')
        name_docx = str(index+'_'+temp[:name_lenght])
        try:
            create_table(context, name_docx)
            template = DocxTemplate(f'annotations/{name_docx}.docx')
            template.render(context)
            template.save("annotations/%s.docx" %str(index+'_'+temp[:name_lenght]))
        except:
            print("Oops!")
        return print(f'successful create {count}/{lenght}')



## connect API google sheets
CREDENTIALS_FILE = 'annotation2-c46cdd2a61d8.json' # file with project data
spreadsheetId  = '1SEz5_QVrWDzSb6g60KlwxShdmPGHSnq3EBw5zJ8k-jQ' #token google sheet

credentials = ServiceAccountCredentials.from_json_keyfile_name(CREDENTIALS_FILE, 
                ['https://www.googleapis.com/auth/spreadsheets',
                'https://www.googleapis.com/auth/drive'])

httpAuth = credentials.authorize(httplib2.Http()) 
service = apiclient.discovery.build('sheets', 'v4', http = httpAuth)
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


if __name__=='__main__':       
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

    for i in range(0,len(dict_)):
        generate_docx_from_schedule(dict_[i], len(dict_), i+1)