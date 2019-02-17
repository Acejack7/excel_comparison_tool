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


def get_files(file_path):
    if os.path.isdir(file_path):
        file_path_split = file_path.split('\\')
        cur_dir = ''
        for elem in file_path_split[:-1]:
            cur_dir += elem
            if file_path_split != elem:
                cur_dir += '\\'

        file_list = os.listdir(file_path)
        full_file_list = []
        for f in file_list:
            new_f = os.path.join(file_path, f)
            full_file_list.append(new_f)

        lang_dir = file_path_split[-1]
        if len(lang_dir) == 4:
            lang_code = lang_dir[:2] + '-' + lang_dir[2:]
        elif len(lang_dir) == 5:
            lang_code = lang_dir[:2] + '-' + lang_dir[3:]
        else:
            lang_code = ''

        return({'files': full_file_list, 'cur_dir': cur_dir, 'lang_code': lang_code})

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
        if file_elem.endswith('.xlsx'):
            wb = load_workbook(file_elem)

            ws = wb.worksheets

            for sheet in ws:
                if sheet.sheet_state == 'visible':

                    source_col = ''
                    target_col = ''

                    for row in sheet.rows:
                        for cell in row:
                            cell_value = str(cell.value).lower()
                            if cell_value.startswith('english') or cell_value.startswith('source'):
                                start_row = cell.row
                                source_col = cell.column
                                break
                        if source_col != '':
                            break

                    for col in sheet.iter_cols(min_row=start_row, max_row=start_row):
                        for cell in col:
                            cell_value = str(cell.value).lower()
                            if cell_value.startswith('translation') or cell_value.startswith('target'):
                                target_col = cell.column
                                break
                            elif target_lang != '':
                                if cell_value in target_lang:
                                    target_col = cell.column
                                    break
                        if target_col != '':
                            break

                    if source_col != '' and target_col != '':
                        for num in range(start_row+1, sheet.max_row):
                            source_content = sheet[source_col + str(num)].value
                            target_content = sheet[target_col + str(num)].value
                            if source_content is not None and target_content is not None:
                                segments.append({'source': source_content, 'target': target_content})

    return(segments)


def compare_contents(translated_content, reviewed_content):
    for translation in translated_content:
        for review in reviewed_content:
            if translation['source'] == review['source']:
                translation['review'] = review['target']
    
    full_content = translated_content
    return(full_content)


def create_report_file(full_content, cur_dir, lang_code):
    wb_report = Workbook()

    ws = wb_report.active
    ws.title = 'Report'

    ws['A1'] = 'Source'
    ws['B1'] = 'Translation'
    ws['C1'] = 'Review'
    ws['D1'] = 'Changes?'

    counter = 2
    for content in full_content:
        row = str(counter)
        source = content['source']
        target = content['target']
        review = content['review']

        ws['A' + row] = source
        ws['B' + row] = target
        ws['C' + row] = review

        if target == review:
            ws['D' + row] = 'No'
        else:
            ws['D' + row] = 'Yes'

        counter += 1
    
    wb_report.save(os.path.join(cur_dir, lang_code + '_report.xlsx'))
    return('Report created succesfully.')
