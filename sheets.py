from docxtpl import DocxTemplate

import os

from service import SetupApp


def generate_docx_from_schedule(context: dict, lenght: int, count: int):
    '''
    This function changes the tag to a word from the google sheet table
    '''
    try:
        if not os.path.isdir('output_data'):
            os.mkdir('output_data')
    except FileExistsError as e:
        pass

    name = context['name_file']
    try:
        # create_table(context, name_docx)
        template = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'template', 'thesis.docx')
        save_dir = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'output_data', f'{name}.docx')
        template = DocxTemplate(template)
        template.render(context)
        template.save(save_dir)
    except:
        print("Oops!")
    return print(f'successful create {count}/{lenght}')


if __name__ == '__main__':
    cn = SetupApp()
    data = cn.get_data()

    true_list = []
    for i, line in enumerate(data):
        if not line[0] == 'FALSE':
            true_list.append(line)

    data_sheet = []
    for row in range(len(true_list)):
        temp = []
        for cell in range(len(true_list[row])):
            new_cell = true_list[row][cell].replace(true_list[row][cell], true_list[row][cell].strip())
            temp.append(new_cell)
            if cell == len(data[row]) - 1:
                data_sheet.append(temp)

    keys = ['filter', 'author', 'topic_ru', 'topic_eng', 'supervisor', 'originality', 'borrowing',
            'date_check', 'date_defense', 'name_file']

    dict_ = {}
    for i, value in enumerate(data_sheet):
        new_match_list = dict(zip(keys, value))
        dict_.setdefault(i, new_match_list)

    data.clear()
    temp.clear()
    data_sheet.clear()

    for i in range(0, len(dict_)):
        generate_docx_from_schedule(dict_[i], len(dict_), i + 1)
