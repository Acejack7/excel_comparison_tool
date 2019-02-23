#!python 3

import os
from openpyxl import load_workbook, Workbook
from openpyxl.styles import Alignment, Font, PatternFill


languages = {'de-de': ['german', 'deutsch'],
             'es-es': ['spanish', 'español'],
             'fr-ca': ['french (canada), french, français'],
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


def verify_excel(files):
    proper_excel_files = []

    for file in files:
        try:
            wb = load_workbook(file)
            wb.close()
            proper_excel_files.append(file)
        except Exception:
            print('%s is not an excel file and got removed from the list.' % file)
            continue

    return(proper_excel_files)


def get_excel_contents(files, target_lang):

    segments = []

    for file_elem in files:
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
                            segments.append({'source': source_content, 'target': target_content,
                                             'file': file_elem, 'sheet': sheet.title, 'row': str(num)})

    return(segments)


def compare_contents(translated_content, reviewed_content):
    full_content = []

    for translation in translated_content:
        for review in reviewed_content:
            if translation['source'] == review['source']:
                source_seg = translation['source']
                target_seg = translation['target']
                rev_seg = review['target']

                trans_file = os.path.split(translation['file'])[1]
                trans_sheet = translation['sheet']
                trans_row = translation['row']

                rev_file = os.path.split(review['file'])[1]
                rev_sheet = review['sheet']
                rev_row = review['row']
                full_content.append({'source': source_seg, 'target': target_seg, 'review': rev_seg,
                                     'trans_file': trans_file, 'trans_sheet': trans_sheet, 'trans_row': trans_row,
                                     'rev_file': rev_file, 'rev_sheet': rev_sheet, 'rev_row': rev_row})
                reviewed_content.remove(review)

    return(full_content)


def sort_by_changes(full_content):
    sorted_full_content = []

    for elem in full_content:
        if elem['target'] != elem['review']:
            sorted_full_content.insert(0, elem)
            elem['changed'] = True
        else:
            sorted_full_content.append(elem)
            elem['changed'] = False

    return(sorted_full_content)


def mark_changes_in_rev(full_content):
    full_content_marked = []

    for elem in full_content:
        if elem['changed'] is True:
            trans = elem['target']
            rev = elem['review']

            trans_split = trans.split(' ')
            rev_split = rev.split(' ')

            last = 'same'
            counter = 0

            review_text = []

            for trans_elem in trans_split:
                index = trans_split.index(trans_elem)

                # catch exception where len(trans_split) > len(rev_split)
                try:
                    rev_elem = rev_split[index]
                except IndexError:
                    break

                if trans_elem == rev_elem:
                    if last != 'same':
                        last = 'same'
                        counter += 1
                    if review_text == []:
                        counter = 0

                    try:
                        review_text[counter] += rev_elem + ' '
                    except IndexError:
                        review_text.append(rev_elem + ' ')

                else:
                    if last != 'diff':
                        last = 'diff'
                        counter += 1
                    if review_text == []:
                        counter = 0

                    try:
                        review_text[counter] += rev_elem + ' '
                    except IndexError:
                        review_text.append(rev_elem + ' ')

            # check if rev_split > trans_split to add the rest of the rev text to last element of review_text
            if len(rev_split) > len(trans_split):
                len_diff = len(rev_split) - len(trans_split)
                missed_elems = rev_split[-len_diff:]
                for elem in missed_elems:
                    review_text[-1] += elem + ' '

            # remove last redundant space
            review_text = review_text[-1].rstrip(' ')

            # assign the review_text list as review text
            elem['review'] = review_text

        full_content_marked.append(elem)

    return(full_content_marked)


def create_report_file(full_content, cur_dir, lang_code):
    wb_report = Workbook()

    ws = wb_report.active
    ws.title = 'Report'

    ws['A1'] = 'Source'
    ws['B1'] = 'Translation'
    ws['C1'] = 'Review'
    ws['D1'] = 'Changed?'
    ws['E1'] = 'Trans File'
    ws['F1'] = 'Trans Sheet'
    ws['G1'] = 'Trans Row'
    ws['H1'] = 'Review File'
    ws['I1'] = 'Review Sheet'
    ws['J1'] = 'Review Row'

    ws.column_dimensions['A'].width = 35
    ws.column_dimensions['B'].width = 35
    ws.column_dimensions['C'].width = 35
    ws.column_dimensions['D'].width = 10
    ws.column_dimensions['E'].width = 50
    ws.column_dimensions['F'].width = 20
    ws.column_dimensions['G'].width = 15
    ws.column_dimensions['H'].width = 50
    ws.column_dimensions['I'].width = 20
    ws.column_dimensions['J'].width = 15

    counter = 2
    for content in full_content:
        row = str(counter)
        source = content['source']
        target = content['target']
        review = content['review']

        trans_file = content['trans_file']
        trans_sheet = content['trans_sheet']
        trans_row = content['trans_row']

        review_file = content['rev_file']
        review_sheet = content['rev_sheet']
        review_row = content['rev_row']

        ws['A' + row] = source
        ws['A' + row].alignment = Alignment(wrap_text=True)
        ws['B' + row] = target
        ws['B' + row].alignment = Alignment(wrap_text=True)
        ws['C' + row] = review
        ws['C' + row].alignment = Alignment(wrap_text=True)

        if content['changed'] is False:
            ws['D' + row] = 'No'
        else:
            ws['D' + row] = 'Yes'
            ws['D' + row].fill = PatternFill(fgColor='FF0000', fill_type='solid')
            ws['C' + row].font = Font(color='FF0000')

        ws['E' + row] = trans_file
        ws['F' + row] = trans_sheet
        ws['G' + row] = trans_row

        ws['H' + row] = review_file
        ws['I' + row] = review_sheet
        ws['J' + row] = review_row

        counter += 1

    wb_report.save(os.path.join(cur_dir, lang_code + '_report.xlsx'))
    return('Report created succesfully.')
