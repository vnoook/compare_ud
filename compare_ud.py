# сравнение номеров дел
import openpyxl
import openpyxl.styles as ops

file_xls = 'compare_ud.xlsx'
start_row = 2
data_col1 = 1
data_col2 = 2
set1 = set()
set2 = set()

wb = openpyxl.load_workbook(file_xls)
ws = wb.active
max_row = ws.max_row

for row in range(start_row, max_row+1):
    cell = ws.cell(row, data_col1)
    orig_content_cell = cell.value
    if orig_content_cell:
        content_cell = str(orig_content_cell.strip()).replace('.', '')
        if content_cell.isdigit():
            set1.add(content_cell)
        else:
            cell.fill = ops.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    else:
        cell.fill = ops.PatternFill(start_color='000000', end_color='FF0000', fill_type='solid')

for row in range(start_row, max_row+1):
    cell = ws.cell(row, data_col2)
    orig_content_cell = cell.value
    if orig_content_cell:
        content_cell = str(orig_content_cell.strip()).replace('.', '')
        if content_cell.isdigit():
            set2.add(content_cell)
        else:
            cell.fill = ops.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    else:
        cell.fill = ops.PatternFill(start_color='000000', end_color='FF0000', fill_type='solid')

dict1 = enumerate(set1)
for i, data in dict1:
    print(i, data)
    ws.cell(start_row+i, data_col1+2).value = data

dict2 = enumerate(set2)
for i, data in dict1:
    print(i, data)
    ws.cell(start_row+i, data_col2+2).value = data


wb.save(file_xls)
