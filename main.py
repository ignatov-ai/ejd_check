# -*- coding: utf8 -*-

from openpyxl import Workbook
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Alignment
from openpyxl.utils import get_column_letter

from datetime import datetime
import os
import sys
import warnings

def find_data(sheet, search_value):
    found = False
    for row in sheet.iter_rows():
        for cell in row:
            if cell.value == search_value:
                coordinate = cell.coordinate
                column = ""
                column_number = 0
                row = ""
                for char in coordinate:
                    if char.isalpha():  # Проверяем, является ли символ буквой
                        column += char
                    elif char.isdigit():  # Проверяем, является ли символ цифрой
                        row += char

                for i, letter in enumerate(reversed(column)):
                    column_number += (ord(letter.upper()) - ord('A') + 1) * (26 ** i)
                return column, column_number, row

    if not found:
        return 'Значение не найдено'


warnings.simplefilter("ignore")

data_str = ['1', '2', '3', '4', '5']
marks_str = ['2', '3', '4', '5']
marks_int = [2, 3, 4, 5]

#######################################################
#################### Загрузка УП ######################
#######################################################
up_file = 'УП'
up_book = load_workbook("up\\" + up_file + '.xlsx')

uchebniy_plan = []

up_sheet = up_book.active

for row in up_sheet.iter_rows(values_only=True):
    uchebniy_plan.append(list(row))

#print(uchebniy_plan)

#######################################################
#######################################################
#######################################################

# Получаем текущую директорию
current_dir = os.getcwd()

# Создаем список для хранения имен файлов
file_list = []

# Перебираем файлы в текущей директории
for root, dirs, files in os.walk(current_dir + '\\journals'):
    for filename in files:
        if filename[-5:] == '.xlsx':
            file_list.append(filename)
# Выводим список имен файлов
# print(file_list)

# общая книга замечаний
comment_book = Workbook()
comment_book_sheet = comment_book.active
#comment_book_sheet.append(['Класс', 'Предмет', 'Учитель', 'Ученик', 'Минимум оценок', - убрал столбик Учитель
comment_book_sheet.append(['Класс', 'Ученик', 'Предмет', 'Минимум оценок',
                           'Кол-во оценок', 'Не хватает оценок'])
comment_book_sheet.column_dimensions['A'].width = 6
comment_book_sheet.column_dimensions['B'].width = 50
comment_book_sheet.column_dimensions['C'].width = 50
comment_book_sheet.column_dimensions['D'].width = 5
comment_book_sheet.column_dimensions['E'].width = 5
comment_book_sheet.column_dimensions['F'].width = 5
#comment_book_sheet.column_dimensions['G'].width = 5

first_row = comment_book_sheet["1"]
for cell in first_row:
    style = Alignment(text_rotation=90, horizontal='center', vertical='center')
    cell.alignment = style

# выбирается файл для обработки
for file in file_list:
    file = file.split('.xlsx')[0]

    class_name = file.split(";")[0]
    lesson = file.split(";")[1]
    group = file.split(";")[2]

    book = load_workbook("journals\\" + file + '.xlsx')
    sheet = book.active


    # удаляем столбцы с ДЗ
    sheet.delete_cols(20, 23)

    # получаем список объединенных диапазонов и разъединяем объединенные ячейки
    merged_cells = list(map(str, sheet.merged_cells.ranges))

    for item in merged_cells:
        try:
            sheet.unmerge_cells(item)
        except KeyError:
            continue

    # перемещаем куски оценок в одну общую часть
    for part in range(1, 6):
        m_range = 'C'+str(1+part*50)+':S'+str((part+1)*50)
        sheet.move_range(m_range, rows=-50*part, cols=17*part)

    # считаем количество учеников на листе
    students_count = 1
    while sheet['A'+str(students_count)].value != '':
        #print(sheet['A' + str(students_count)].value,students_count)
        students_count += 1
    #students_count -= 3

    # удаляем пустые строки
    sheet.delete_rows(students_count, sheet.max_row)

    # зменяем ширину всех столбцов с оценками
    for i in range(3, int(sheet.max_column)):
        sheet.column_dimensions[get_column_letter(i)].width = 2

    # меняем тип ячейки на число, если это возможно
    for row in range(students_count-3):
        for col in range(3, int(sheet.max_column)):
            ch_range = get_column_letter(col) + str(row+3)
            if sheet[ch_range].value in data_str:
                sheet[ch_range].value = int(sheet[ch_range].value)

    # удаление лишних столбцов
    t1_column, t1_column_number, t1_row = find_data(sheet, "Т1")
    sheet.delete_cols(t1_column_number)
    t2_column, t2_column_number, t2_row = find_data(sheet, "Т2")
    sheet.delete_cols(t2_column_number)
    t3_column, t3_column_number, t3_row = find_data(sheet, "Т3")
    sheet.delete_cols(t3_column_number)
    g_column, g_column_number, g_row = find_data(sheet, "Г")
    sheet.delete_cols(g_column_number)

    #подсчет проведенных уроков на данную дату и удаление не нужных столбцов
    lessons_count = 0
    month_count = 1
    old_day_value = 0
    curent_month = 9
    curent_day = 0

    now_month = datetime.now().month
    now_day = datetime.now().day
    print(now_month, now_day)

    for i in range(2, sheet.max_column+1):
        new_cell_value = sheet[str(get_column_letter(i)) + "2"].value
        print(new_cell_value)

        if new_cell_value != "" and new_cell_value != None:
            curent_day = int(new_cell_value)
            lessons_count += 1
        elif new_cell_value == "":
            break

        if curent_day < int(old_day_value):
            month_count += 1
            curent_month += 1

        print('месяц:', curent_month, 'день', curent_day)

        if curent_month >= now_month and curent_day >= now_day:
            print("дальше еще уроков не было")
            last_column_index = i + 2 + 1
            break

        old_day_value = curent_day

    print(lessons_count)
    print(month_count)

    sheet.delete_cols(last_column_index, sheet.max_column)

    # подсчет количества оценок по ученикам




    # созранение переработанной книги в data
    book.save('data\\' + file + '_test.xlsx')
    book.close()