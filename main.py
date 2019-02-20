#! python3

import excel_comparison

if __name__ == '__main__':
    print('Welcome in Excel Comparison Tool.')
    translation_files = input('Please provide file path to translated file(s): ')
    review_files = input('Please probide file path to reviewed file(s): ')

    # get information about provided paths, files
    translated_files_info = excel_comparison.get_files(translation_files)
    reviewed_files_info = excel_comparison.get_files(review_files)

    # separate the information about files and target language
    lang_code = translated_files_info['lang_code']
    cur_work_dir = translated_files_info['cur_dir']
    trans_file_list = translated_files_info['files']
    review_file_list = reviewed_files_info['files']

    # verify excel files
    verified_trans_files = excel_comparison.verify_excel(trans_file_list)
    verified_review_files = excel_comparison.verify_excel(review_file_list)

    # check target language
    target_lang = excel_comparison.get_target_lang(lang_code)

    translated_content = excel_comparison.get_excel_contents(verified_trans_files, target_lang)
    reviewed_content = excel_comparison.get_excel_contents(verified_review_files, target_lang)

    full_content = excel_comparison.compare_contents(translated_content, reviewed_content)

    print(excel_comparison.create_report_file(full_content, cur_work_dir, lang_code))