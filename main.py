import datetime

import xlsxwriter


from data import data_lst, dates2_lst


slno = True

with xlsxwriter.Workbook('test.xlsx') as workbook:
    # worksheet = workbook.add_worksheet('Sheet1')
    worksheet = workbook.add_worksheet()

    no_of_columns_to_freeze = len(list(data_lst[0].keys())[:-1])+1 if slno else len(list(data_lst[0].keys())[:-1])
    worksheet.freeze_panes(1, no_of_columns_to_freeze)  # Freeze the first row and first 3 columns.

    # no_of_vertical_cells_to_merge: int = len(data_lst[0]["Time"][0].keys())-1    # -1 to exclude the "Date" key from "Time"
    no_of_vertical_cells_to_merge: int = 2

    print('no_of_vertical_cells_to_merge-----> ', no_of_vertical_cells_to_merge)
    
    col_headers: list = [    # if sl no. else not
        'Sl No.',
    ]
    col_headers += list(data_lst[0].keys())[:-1] + dates2_lst + [list(data_lst[0].keys())[-1]]    # Appending OD(mins.) in the end

    print('col_headers----> ', col_headers)

    header_date_format = workbook.add_format({
                                    'num_format': 'yyyy-mm-dd',
                                    'bold': True,
                                    'align': 'center',
                                    'valign': 'vcenter',
                                    'font_color': 'white',
                                    'bg_color': 'green',
                                    'text_wrap': True,
                                    'border': 1
                    })
    datetime_format = workbook.add_format({
                                'num_format': 'yyyy-mm-dd hh:mm:ss',
                                'align': 'center',
                                'valign': 'vcenter',
                                'text_wrap': True,
                                'border': 1
                    })
    text_format = workbook.add_format({
                                'align': 'center',
                                'valign': 'vcenter',
                                'text_wrap': True,
                                'border': 1
                    })
    header_format = workbook.add_format({
                                    'bold': True,
                                    'align': 'center',
                                    'valign': 'vcenter',
                                    'font_color': 'white',
                                    'bg_color': 'green',
                                    'text_wrap': True,
                                    'border': 1
                    })
    header_slno_format = workbook.add_format({
                                    'bold': True,
                                    'align': 'center',
                                    'valign': 'vcenter',
                                    'font_color': 'white',
                                    'bg_color': '#3366FF',
                                    'text_wrap': True,
                                    'border': 1
                    })

    worksheet.autofilter(0, 0, 0, len(col_headers)-1)    # 'first_row', 'first_col', 'last_row', and 'last_col' & -1counting starts from 0.

    col_header_index: dict = {}

    # Column Headers inserting
    for col_index, each_col_header in enumerate(col_headers):
        if isinstance(each_col_header, datetime.date):
            worksheet.write_datetime(0, col_index, each_col_header, header_date_format)
        else:
            worksheet.write(0, col_index, each_col_header, header_format)    # 1st row = 0
        col_header_index[each_col_header] = col_index

    row_index = 1    # From 2nd row, as 1st is header row.
    for sl_no, each_data_dict in enumerate(data_lst):
        sl_no+=1

        for each_col_header_index in col_header_index:

            if isinstance(each_col_header_index, datetime.date) or each_col_header_index == 'Time':
                per_user_vertical_data: dict = each_data_dict['Time'][each_col_header_index]
                per_user_vertical_keys: list = list(each_data_dict['Time'][each_col_header_index].keys())
                # per_user_vertical_len: int = len(list(each_data_dict['Time'][each_col_header_index].keys()))

                if isinstance(each_col_header_index, datetime.date):
                    for len_, per_user_vertical_key in enumerate(per_user_vertical_keys):
                        worksheet.write(row_index+len_, col_header_index[each_col_header_index], each_data_dict['Time'][each_col_header_index][per_user_vertical_key], datetime_format)
                        worksheet.set_column(col_header_index[each_col_header_index], col_header_index[each_col_header_index], 20)    # 20 px

                elif each_col_header_index == 'Time':
                    for len_, per_user_vertical_key in enumerate(per_user_vertical_keys):
                        worksheet.write(row_index+len_, col_header_index[each_col_header_index], per_user_vertical_key, text_format)

            elif each_col_header_index == 'Sl No.':
                worksheet.merge_range(row_index, col_header_index[each_col_header_index], row_index+no_of_vertical_cells_to_merge-1, col_header_index[each_col_header_index], sl_no, text_format)    # -1 because if 2 needs to be merged then row_index itself is consuming 1 so we need +1 so we do -1 to cancelt it out.
                # worksheet.write(row_index, col_header_index[each_col_header_index], sl_no)


            else:
                worksheet.merge_range(row_index, col_header_index[each_col_header_index], row_index+no_of_vertical_cells_to_merge-1, col_header_index[each_col_header_index], each_data_dict[each_col_header_index], text_format)    # -1 because if 2 needs to be merged then row_index itself is consuming 1 so we need +1 so we do -1 to cancelt it out.
                # worksheet.write(row_index, col_header_index[each_col_header_index], each_data_dict[each_col_header_index])

        row_index += no_of_vertical_cells_to_merge    # As merging two rows

    worksheet.autofit()    # Makes column width as the max size of the top row, Doesn't override changes done by set_columns() above.
