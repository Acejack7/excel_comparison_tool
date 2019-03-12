#! python3

import excel_comparison

if __name__ == '__main__':
    print('Welcome in Excel Comparison Tool.')
    translation_files = input('Please provide file path to translated file(s): ')
    review_files = input('Please provide file path to reviewed file(s): ')
    user_src_col = input('Please provide source column or leave empty to let computer recognize: ')
    user_trg_col = input('Please provide translation column or leave empty to let computer recognize: ')
    user_rev_col = input('Please provide review column or leave empty to let computer recognize: ')

    # verify columns provided by user
    while excel_comparison.verify_column(user_src_col) is False:
        print('Please provide proper column letter for source content.')
        user_src_col = input('Please provide source column or leave empty to let computer recognize: ')
        if user_src_col == '':
            break

    while excel_comparison.verify_column(user_trg_col) is False:
        print('Please provide proper column letter for translated content.')
        user_trg_col = input('Please provide translation column or leave empty to let computer recognize: ')
        if user_trg_col == '':
            break

    while excel_comparison.verify_column(user_rev_col) is False:
        print('Please provide proper column letter for reviewed content.')
        user_rev_col = input('Please provide review column or leave empty to let computer recognize: ')
        if user_trg_col == '':
            break

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

    # get translated and reviewed content: segments and file, sheet names and row
    translated_content = excel_comparison.get_excel_contents(verified_trans_files, target_lang, user_src_col, user_trg_col)
    reviewed_content = excel_comparison.get_excel_contents(verified_review_files, target_lang, user_src_col, user_rev_col)

    # compare translated and reviewed contents and merge them together
    full_content = excel_comparison.compare_contents(translated_content, reviewed_content)

    # sort by changes
    full_content_sorted = excel_comparison.sort_by_changes(full_content)

    # get changes in review and mark them
    full_content_marked = excel_comparison.mark_changes_in_rev(full_content_sorted)

    # create report file
    report_file = excel_comparison.create_report_file(full_content_marked, cur_work_dir, lang_code)
    print(report_file)

    # add additional data to report file
    # print(excel_comparison.add_data_to_report(full_content_marked, report_file))
