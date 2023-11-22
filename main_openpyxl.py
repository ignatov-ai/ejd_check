from openpyxl import load_workbook
from openpyxl.utils import get_column_letter


import datetime

date = datetime.datetime.now()

book = load_workbook("11 Я тест.xlsx")
sheets = book.active

# удаляем столбцы с ДЗ
sheets.delete_cols(20, 23)

# получаем список объединенных диапазонов и разъединяем объединенные ячейки
merged_cells = list(map(str, sheets.merged_cells.ranges))

for item in merged_cells:
    try:
        sheets.unmerge_cells(item)
    except KeyError:
        continue

# перемещаем куски оценок в одну общую часть
for part in range(1, 6):
    m_range = 'C'+str(1+part*50)+':S'+str((part+1)*50)
    sheets.move_range(m_range, rows=-50*part, cols=17*part)

# считаем количество учеников на листе
i = 1
while sheets['A'+str(i)].value != '':
    i += 1

# удаляем пустые строки
sheets.delete_rows(i, sheets.max_row)

# зменяем ширину всех столбцов с оценками
for i in range(3, int(sheets.max_column)):
    sheets.column_dimensions[get_column_letter(i)].width = 2

book.save("primer2.xlsx")
book.close()