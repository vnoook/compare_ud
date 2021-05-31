# сравнение номеров дел
# ...
# INSTALL
# pip install openpyxl

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

# сбор, очистка и обработка данных из колонок с вставленными данными
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
# чистка первой колонки - без дублей и пустых ячеек
tuple1 = enumerate(set1)
for i1, data1 in tuple1:
    ws.cell(start_row+i1, 3).value = data1

# чистка второй колонки - без дублей и пустых ячеек
tuple2 = enumerate(set2)
for i2, data2 in tuple2:
    ws.cell(start_row+i2, 4).value = data2

# Объединение списков без дублей
all_cols = set1.union(set2)
tuple3 = enumerate(all_cols)
for i3, data3 in tuple3:
    ws.cell(start_row+i3, 5).value = data3

# общие значения
common_values = set1.intersection(set2)
tuple4 = enumerate(common_values)
for i4, data4 in tuple4:
    ws.cell(start_row+i4, 6).value = data4

# пересечение 1
diff_set1 = set1.difference(set2)
tuple5 = enumerate(diff_set1)
for i5, data5 in tuple5:
    ws.cell(start_row+i5, 7).value = data5

# пересечение 2
diff_set2 = set2.difference(set1)
tuple6 = enumerate(diff_set2)
for i6, data6 in tuple6:
    ws.cell(start_row+i6, 8).value = data6


wb.save(file_xls)
