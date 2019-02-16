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

        return(file_list, cur_dir, lang_code)

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

        return(file_path, cur_dir, lang_code)



        
