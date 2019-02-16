#!python 3

import os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill


languages = {'de-de': ['german', 'deutsch'],
             'es-es': ['spanish', 'español'],
             'fr-fr': ['french', 'français'],
             'ja-jp': ['japanese', '日本人'],
             'it-it': ['italian', 'イタリアの'],
             'pt-br': ['portugese', 'portugese (brazilian)'],
             }


def get_files(file_path, type):
    if os.path.isdir(file_path):
        file_path_split = file_path.split('\\')
        cur_dir = ''
        for elem in file_path_split:
            cur_dir += elem
            if file_path_split != elem:
                cur_dir += '\\'

        file_list = os.listdir(file_path)

        lang_dir = file_path_split[-1]
        if len(lang_dir) == 4:
            lang_code = lang_dir[:2] + '-' + lang_dir[2:]
        elif len(lang_dir) == 5:
            lang_code = lang_dir[:2] + '-' + lang_dir[3:]
        else:
            lang_code = ''

        return({'files': file_list, 'cur_dir': cur_dir, 'lang_code': lang_code})

    elif os.path.isfile(file_path):
        file_path_split = file_path.split('\\')
        file_path_split = file_path_split[:-1]
        cur_dir = ''
        for elem in file_path_split:
            cur_dir += elem
            if file_path_split != elem:
                cur_dir += '\\'

        lang_dir = file_path_split[-1]
        if len(lang_dir) == 4:
            lang_code = lang_dir[:2] + '-' + lang_dir[2:]
        elif len(lang_dir) == 5:
            lang_code = lang_dir[:2] + '-' + lang_dir[3:]
        else:
            lang_code = ''

        return({'files': [file_path], 'cur_dir': cur_dir, 'lang_code': lang_code})


def get_target_lang(lang_code):
    if lang_code in languages:
        return(languages[lang_code])
    else:
        return('')


def get_excel_contents(files, target_lang):

    segments = []

    for file_elem in files:
        wb = load_workbook(file_elem)

        ws = wb.worksheets

        for sheet in ws:
            if sheet.sheet_state == 'visible':
                for row in sheet.rows:
                    for cell in row:
                        cell_value = str(cell.value).lower()
                        if cell_value.startswith('english') or cell_value.startswith('source'):
                            start_row = cell.row + 1
                            source_col = cell.column
                            break
                    else:
                        continue
                    break

                for col in sheet.iter_cols(min_row=start_row, max_row=start_row):
                    for cell in col:
                        cell_value = str(cell.value).lower()
                        if cell_value.startswith('translation') or cell_value.startswith('target'):
                            target_col = cell.column
                            break
                        elif cell_value in target_lang:
                            target_col = cell.column
                            break
                
                for num in range(start_row, sheet.max_row):
                    source_content = sheet[source_col + str(num)].value
                    target_content = sheet[target_col + str(num)].value
                    if source_content is not None and target_content is not None:
                        segments.append({'source': source_content, 'target': target_content})
    
    return(segments)
