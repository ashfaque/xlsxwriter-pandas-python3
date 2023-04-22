import datetime


import xlsxwriter

# ! index starts from 0

# ? https://xlsxwriter.readthedocs.io/contents.html

with xlsxwriter.Workbook('hello_world.xlsx') as workbook:
    worksheet = workbook.add_worksheet()

    worksheet.write('A1', 'Hello world')

###

# https://xlsxwriter.readthedocs.io/working_with_dates_and_time.html

workbook = xlsxwriter.Workbook('Expenses01.xlsx')
worksheet = workbook.add_worksheet('Sheet1')

row, col, item = 0, 0, 'Hulululu'
worksheet.write(row, col, item)    # ? Throughout XlsxWriter, rows and columns are zero indexed. The first cell in a worksheet, A1, is (0, 0).


bold = workbook.add_format({'bold': True})

worksheet.write('A1', 'Item', bold)    # .write(row, column, token , [format])

money = workbook.add_format({'num_format': '₹#,##0'})
worksheet.write(row+2, col+2, 1000, money)


date_format = workbook.add_format({'num_format': 'mmmm d yyyy'})
date_format = workbook.add_format({'num_format': 'yyyy-mm-dd'})
worksheet.write_datetime(row, col + 1, datetime.datetime(2023,4,1,10,0,0), date_format)


# worksheet.write_string  (row, col,     item              )
# worksheet.write_datetime(row, col + 1, date, date_format )
# worksheet.write_number  (row, col + 2, cost, money_format)

# write_string()
# write_number()
# write_blank()
# write_formula()
# write_datetime()
# write_boolean()
# write_url()

workbook = xlsxwriter.Workbook(filename, {'strings_to_numbers': True})


worksheet1 = workbook.add_worksheet()           # Sheet1
worksheet2 = workbook.add_worksheet('Foglio2')  # Foglio2
worksheet3 = workbook.add_worksheet('Data')     # Data
worksheet4 = workbook.add_worksheet()           # Sheet4, The name parameter is optional. If it is not specified, or blank, the default Excel convention will be followed, i.e. Sheet1, Sheet2, etc,


workbook.set_properties({
    'title':    'This is an example spreadsheet',
    'subject':  'With document properties',
    'author':   'John McNamara',
    'manager':  'Dr. Heinz Doofenshmirtz',
    'company':  'of Wolves',
    'category': 'Example spreadsheets',
    'keywords': 'Sample, Example, Properties',
    'created':  datetime.date(2018, 1, 1),
    'comments': 'Created with Python and XlsxWriter'})

workbook.set_custom_property('Reference number', 1.2345)

workbook.read_only_recommended()    #  This means that any changes to the file can’t be saved back to the same file and must be saved to a new file

worksheet.write_rich_string('A1',
                            'This is an ', format, 'example', ' string')    # The basic rule is to break the string into fragments and put a Format object before the fragment that you want to format



# Some bold format and a default format.
bold    = workbook.add_format({'bold': True})
default = workbook.add_format()
center = workbook.add_format({'align': 'center'})

# With default formatting:
worksheet.write_rich_string('A1',
                            'Some ',
                            bold, 'bold',
                            ' text')

# Or more explicitly:
worksheet.write_rich_string('A1',
                             default, 'Some ',
                             bold,    'bold',
                             default, ' text')


link_format = workbook.add_format({'color': 'blue', 
                                   'underline': True, 
                                   'text_wrap': True})

cell_format = workbook.add_format()
cell_format.set_text_wrap()

worksheet.write(0, 0, "Some long text to wrap in a cell", cell_format)


# ? https://xlsxwriter.readthedocs.io/working_with_colors.html#colors
cell_format = workbook.add_format()

cell_format.set_pattern(1)  # This is optional when using a solid fill.
cell_format.set_bg_color('green')

worksheet.write('A1', 'Ray', cell_format)



#The autofit() method won’t override a user defined column width set with set_column() or set_column_pixels() if it is greater than the autofit value. This allows the user to set a minimum width value for a column.
worksheet.autofit()    # Auto fit column width, starting xlsxwriter 3.0.6+

worksheet1.insert_image('B10', '../images/python.png')
worksheet1.insert_image('B2', 'python.png', {'x_offset': 15, 'y_offset': 10})
worksheet.insert_image('B3', 'python.png', {'x_scale': 0.5, 'y_scale': 0.5})
worksheet.insert_image('B4', 'python.png', {'url': 'https://python.org'})

url = 'https://python.org/logo.png'
image_data = io.BytesIO(urllib2.urlopen(url).read())

worksheet.insert_image('B5', url, {'image_data': image_data})

worksheet.insert_image('B3', 'python.png',
                       {'description': 'The logo of the Python programming language.'})

worksheet.insert_image('B3', 'python.png', {'decorative': True})

worksheet.insert_image('B3', 'python.png', {'object_position': 1})


# tables, charts, button macros, data validation, add_sparkline() -> in-cell charts, comments, formulas etc. prossible
# ? https://xlsxwriter.readthedocs.io/chart.html
# ? https://xlsxwriter.readthedocs.io/chartsheet.html
# ? https://xlsxwriter.readthedocs.io/working_with_formulas.html
# ? https://xlsxwriter.readthedocs.io/working_with_tables.html
# ? https://xlsxwriter.readthedocs.io/working_with_sparklines.html
# ? https://xlsxwriter.readthedocs.io/working_with_data_validation.html

# conditional_format(first_row, first_col, last_row, last_col, options)
# https://xlsxwriter.readthedocs.io/working_with_conditional_formats.html
'''
Returns: 0: Success.Returns: -1: Row or column is out of worksheet bounds.
Returns: -2: Incorrect parameter or option.
'''
worksheet.conditional_format('B3:K12', {'type':     'cell',
                                        'criteria': '>=',
                                        'value':    50,
                                        'format':   format1})

worksheet.conditional_format(0, 0, 2, 1, {'type':     'cell',
                                          'criteria': '>=',
                                          'value':    50,
                                          'format':   format1})




worksheet3.activate()

worksheet2.hide()

worksheet19.set_first_sheet()  # First visible worksheet tab.


# worksheet.merge_range()
# merge_range(first_row, first_col, last_row, last_col, data[, cell_format])
merge_format = workbook.add_format({'align': 'center'})
merge_format = workbook.add_format({
    'bold':     True,
    'border':   6,
    'align':    'center',
    'valign':   'vcenter',
    'fg_color': '#D7E4BC',
})

worksheet.merge_range('B3:D4', 'Merged Cells', merge_format)
worksheet.merge_range(2, 1, 3, 3, 'Merged Cells', merge_format)


# worksheet.autofilter()
worksheet.autofilter(0, 0, 10, 3)
worksheet.autofilter('A1:D11')
# ? https://xlsxwriter.readthedocs.io/working_with_autofilters.html#working-with-autofilters


# default filter also possible -> worksheet.filter_column(), The filter_column method can be used to filter columns in a autofilter range based on simple conditions.
worksheet.filter_column('A', 'x > 2000')
worksheet.filter_column('B', 'x > 2000 and x < 5000')
worksheet.filter_column(2,   'x > 2000')
worksheet.filter_column('C', 'x > 2000')

# worksheet.filter_column_list(), see 
# ? https://xlsxwriter.readthedocs.io/worksheet.html#:~:text=worksheet.filter_column_list()

# worksheet.set_top_left_cell()
worksheet.set_top_left_cell(31, 26)
worksheet.set_top_left_cell('AA32')


# worksheet.freeze_panes()
# freeze_panes(row, col[, top_row, left_col])
worksheet.freeze_panes(1, 0)  # Freeze the first row.
worksheet.freeze_panes('A2')  # Same using A1 notation.
worksheet.freeze_panes(0, 1)  # Freeze the first column.
worksheet.freeze_panes('B1')  # Same using A1 notation.
worksheet.freeze_panes(1, 2)  # Freeze first row and first 2 columns.
worksheet.freeze_panes('C2')  # Same using A1 notation.

# The parameters top_row and left_col are optional. They are used to specify the top-most or left-most visible row or column in the scrolling region of the panes. For example to freeze the first row and to have the scrolling region begin at row twenty:
worksheet.freeze_panes(1, 0, 20, 0)


# worksheet.split_panes() -> https://xlsxwriter.readthedocs.io/example_panes.html#ex-panes
# https://xlsxwriter.readthedocs.io/worksheet.html#:~:text=worksheet.split_panes()


worksheet1.set_zoom(50)
worksheet2.set_zoom(75)
worksheet3.set_zoom(300)
worksheet4.set_zoom(400)

# Urdu or Arabic
worksheet.right_to_left()


worksheet.hide_zero()    # The hide_zero() method is used to hide any zero values that appear in cells:


worksheet.set_background('logo.png')

worksheet.set_header('&C&G', {'image_center': 'watermark.png'})


worksheet1.set_tab_color('red')
worksheet2.set_tab_color('#FF9900')  # Orange

worksheet.protect('PasswordHere')

worksheet.protect('abc123', {'insert_rows': True})    # protect particular thing

# Worksheet level passwords in Excel offer very weak protection. They do not encrypt your data and are very easy to deactivate. Full workbook encryption is not supported by XlsxWriter. However, it is possible to encrypt an XlsxWriter file using a third party open source tool called msoffice-crypt. This works for macOS, Linux and Windows:
# https://github.com/herumi/msoffice
# msoffice-crypt.exe -e -p password clear.xlsx encrypted.xlsx



# worksheet.unprotect_range()
worksheet.unprotect_range('A1')
worksheet.unprotect_range('C1')
worksheet.unprotect_range('E1:E3')
worksheet.unprotect_range('G1:K100')

worksheet.unprotect_range('G4:I6', 'MyRange')    # As in Excel the ranges are given sequential names like Range1 and Range2 but a user defined name can also be specified:

workbook.close()

##########
# https://xlsxwriter.readthedocs.io/working_with_data.html
#import xlsxwriter

workbook = xlsxwriter.Workbook('write_list.xlsx')
worksheet = workbook.add_worksheet()

my_list = [[1, 1, 1, 1, 1],
           [2, 2, 2, 2, 1],
           [3, 3, 3, 3, 1],
           [4, 4, 4, 4, 1],
           [5, 5, 5, 5, 1]]

for row_num, row_data in enumerate(my_list):
    for col_num, col_data in enumerate(row_data):
        worksheet.write(row_num, col_num, col_data)




############
# ? https://xlsxwriter.readthedocs.io/working_with_pandas.html
# ? https://xlsxwriter.readthedocs.io/example_pandas_multiple.html
import pandas as pd

# Create a Pandas dataframe from the data.
df = pd.DataFrame({'Data': [10, 20, 30, 20, 15, 30, 45]})

# Create a Pandas Excel writer using XlsxWriter as the engine.
writer = pd.ExcelWriter('pandas_simple.xlsx', engine='xlsxwriter')

# Convert the dataframe to an XlsxWriter Excel object.
df.to_excel(writer, sheet_name='Sheet1')

# Close the Pandas Excel writer and output the Excel file.
writer.close()

# ? XlsxPandasFormatter is a helper class that wraps the worksheet, workbook and dataframe objects written by Pandas to_excel() method using the xlsxwriter engine to allow consistent formatting of cells.

##########
















workbook.close()







# ? Print: https://xlsxwriter.readthedocs.io/page_setup.html

# ! https://xlsxwriter.readthedocs.io/working_with_memory.html

# ? Cell Notation: https://xlsxwriter.readthedocs.io/working_with_cell_notation.html

# * Exceptions: https://xlsxwriter.readthedocs.io/exceptions.html



#################################################################################

import xlsxwriter
from data import data_lst

# Create a new Excel file and add a worksheet
workbook = xlsxwriter.Workbook('attendance.xlsx')
worksheet = workbook.add_worksheet()

# Define cell formats
header_format = workbook.add_format({
    'bold': True,
    'align': 'center',
    'valign': 'vcenter',
    'font_color': 'white',
    'bg_color': 'gray'
})
merge_format = workbook.add_format({
    'align': 'center',
    'valign': 'vcenter',
    'border': 1
})
time_format = workbook.add_format({
    'num_format': 'hh:mm:ss',
    'align': 'center',
    'valign': 'vcenter',
    'border': 1
})

# Define header titles
header_titles = ['SAP ID', 'Name', 'OD (mins.)', 'Date', 'Login Time', 'Logout Time']

# Write headers to worksheet
for col, title in enumerate(header_titles):
    worksheet.write(0, col, title, header_format)

# Merge header cells
worksheet.merge_range(0, 0, 0, 2, 'Employee Information', merge_format)
worksheet.merge_range(0, 3, 0, 5, 'Attendance Information', merge_format)

# Write data to worksheet
for row, data in enumerate(data_lst, start=1):
    # Merge first 3 columns and write employee info
    worksheet.merge_range(row, 0, row, 2, f'{data["SAP ID"]}\n{data["Name"]}', merge_format)
    worksheet.write(row, 2, data['OD (mins.)'], merge_format)
    
    # Write attendance info
    for attendance in data['Time']:
        worksheet.write(row, 3, attendance['Date'].strftime('%Y-%m-%d'), merge_format)
        worksheet.write(row, 4, attendance['Login Time'], time_format)
        worksheet.write(row, 5, attendance['Logout Time'], time_format)
        row += 1
    
# Autofit columns
for col in range(len(header_titles)):
    worksheet.set_column(col, col, 20)

# Close workbook
workbook.close()
