# сравнение номеров дел
import openpyxl
import openpyxl.styles

# переменные для работы
file_xls = 'compare_ud.xlsx'
# строка с которой начинать собирать данные
start_row = 2
# колонки в которые пользователь вставляет данные
data_col1 = 1
data_col2 = 2
# два множества для данных из колонок со вставленными данными
set1 = set()
set2 = set()

# открывается существующая книга эксель
wb = openpyxl.load_workbook(file_xls)
ws = wb.active
# переменная самой последней заполненной строки
max_row = ws.max_row


for row in range(start_row, max_row+1):
    cell = ws.cell(row, data_col1)
    orig_content_cell = cell.value
    if orig_content_cell:
        content_cell = str(orig_content_cell.strip()).replace('.', '')
        if content_cell.isdigit():
            set1.add(content_cell)
        else:
            cell.fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    else:
        cell.fill = openpyxl.styles.PatternFill(start_color='000000', end_color='FF0000', fill_type='solid')

for row in range(start_row, max_row+1):
    cell = ws.cell(row, data_col2)
    orig_content_cell = cell.value
    if orig_content_cell:
        content_cell = str(orig_content_cell.strip()).replace('.', '')
        if content_cell.isdigit():
            set2.add(content_cell)
        else:
            cell.fill = openpyxl.styles.PatternFill(start_color='FF0000', end_color='FF0000', fill_type='solid')
    else:
        cell.fill = openpyxl.styles.PatternFill(start_color='000000', end_color='FF0000', fill_type='solid')


# заполнение колонок в экселе
tuple1 = enumerate(set1)
for i1, data1 in tuple1:
    ws.cell(start_row+i1, data_col1+2).value = data1


tuple2 = enumerate(set2)
for i2, data2 in tuple2:
    ws.cell(start_row+i2, data_col2+2).value = data2

wb.save(file_xls)
