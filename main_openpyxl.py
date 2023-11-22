import openpyxl.styles
from openpyxl import load_workbook
from openpyxl.styles import PatternFill, Font, Alignment
from openpyxl.utils import get_column_letter

import datetime
import os

import warnings
warnings.simplefilter("ignore")

marks_str = ['2', '3', '4', '5']
marks_int = [2, 3, 4, 5]

#######################################################
#################### Загрузка УП ######################
#######################################################
up_file = 'УП'
up_book = load_workbook("UP\\" + up_file + '.xlsx')

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
for root, dirs, files in os.walk(current_dir + '\data'):
    for filename in files:
        if filename[-5:] == '.xlsx':
            file_list.append(filename)
# Выводим список имен файлов
#print(file_list)

# общая книга замечаний
comment_book = openpyxl.Workbook()
comment_book_sheet = comment_book.active
comment_book_sheet.append(['Класс', 'Предмет', 'Учитель', 'Ученик', 'Минимум оценок',
                           'Кол-во оценок', 'Не хватает оценок'])
comment_book_sheet.column_dimensions['A'].width = 6
comment_book_sheet.column_dimensions['B'].width = 65
comment_book_sheet.column_dimensions['C'].width = 40
comment_book_sheet.column_dimensions['D'].width = 40
comment_book_sheet.column_dimensions['E'].width = 5
comment_book_sheet.column_dimensions['F'].width = 5
comment_book_sheet.column_dimensions['G'].width = 5

first_row = comment_book_sheet["1"]
for cell in first_row:
    style = Alignment(text_rotation=90, horizontal='center', vertical='center')
    cell.alignment = style

# выбирается файл для обработки
for file in file_list:
    file = file.split('.xlsx')[0]

    book = load_workbook("data\\" + file + '.xlsx')

    comment_class_book = openpyxl.Workbook()
    comment_class_book_sheet = comment_class_book.active
    comment_class_book_sheet.append(['Класс', 'Предмет', 'Учитель', 'Ученик', 'Минимум оценок',
                               'Кол-во оценок', 'Не хватает оценок'])
    comment_class_book_sheet.column_dimensions['A'].width = 6
    comment_class_book_sheet.column_dimensions['B'].width = 65
    comment_class_book_sheet.column_dimensions['C'].width = 40
    comment_class_book_sheet.column_dimensions['D'].width = 40
    comment_class_book_sheet.column_dimensions['E'].width = 5
    comment_class_book_sheet.column_dimensions['F'].width = 5
    comment_class_book_sheet.column_dimensions['G'].width = 5

    first_row = comment_class_book_sheet["1"]
    for cell in first_row:
        style = Alignment(text_rotation=90, horizontal='center', vertical='center')
        cell.alignment = style

    # получаем список листов
    tabs = book.sheetnames

    # выбирается лист для работы
    for sheet_name in tabs:
        date = datetime.datetime.now()

        sheets = book[sheet_name]

        # определяем предмет
        lesson = sheets['U41'].value.split(',')[-1].lstrip().rstrip()

        # определяем учителя
        teacher_split = sheets['U43'].value.split()[1:4]
        teacher = ' '.join(teacher_split).lstrip().rstrip()

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
        students_count = 1
        while sheets['A'+str(students_count)].value != '':
            #print(sheets['A' + str(students_count)].value,students_count)
            students_count += 1
        #students_count -= 3

        # удаляем пустые строки
        sheets.delete_rows(students_count, sheets.max_row)

        # зменяем ширину всех столбцов с оценками
        for i in range(3, int(sheets.max_column)):
            sheets.column_dimensions[get_column_letter(i)].width = 2

        # меняем тип ячейки на число, если это возможно
        for row in range(students_count-3):
            for col in range(3, int(sheets.max_column)):
                ch_range = get_column_letter(col) + str(row+3)
                if sheets[ch_range].value in marks_str:
                    sheets[ch_range].value = int(sheets[ch_range].value)

        # добавляем колонку КОЛИЧЕСТВО ОЦЕНОК и КОЛИЧЕСТВО ПРОПУСКОВ и подсчитываем количество оценок и пропусков по каждому учащемуся
        sheets['CJ2'].value = 'Количество оценок'
        sheets.column_dimensions['CJ'].width = 11

        sheets['CK2'].value = 'Количество пропусков'
        sheets.column_dimensions['CK'].width = 11

        i = 3
        for col in sheets.iter_rows(min_col=3, max_col=87, min_row=3, max_row=students_count-1):
            marks_count = 0
            n_count = 0
            for cell in col:
                if cell.value in marks_int:
                    marks_count += 1
                elif cell.value == 'н':
                    n_count += 1

            sheets['CJ'+str(i)].value = marks_count
            sheets['CK' + str(i)].value = n_count
            i += 1

        # Вставляем класс, учителя и предмет
        sheets.insert_rows(1, 4)
        sheets['B1'] = file
        sheets['B2'] = lesson
        sheets['B3'] = teacher

        #######################################################
        ###########  Сравнение кол-ва уроков с УП  ############
        #######################################################

        up_class_index = uchebniy_plan[0].index(file.lower())
        #print(up_class_index, file.lower())

        up_lesson_index = 0
        for i in range (len(uchebniy_plan)):
            if uchebniy_plan[i][0] == lesson.lower():
                up_lesson_index = i
        #print(up_lesson_index, lesson.lower())

        up_lesson_nagruzka = uchebniy_plan[up_class_index][up_lesson_index]

        up_marks_number = up_lesson_nagruzka * 2 + 1

        print(file.lower(), lesson.lower(), up_lesson_nagruzka, up_marks_number)

        for row in range(students_count-3):
            ch_cell = 'CJ' + str(row+7)
            fill_cell = sheets[ch_cell]
            #print(ch_cell, sheets[ch_cell].value)
            if fill_cell.value >= up_marks_number:
                fill_cell.fill = PatternFill(fill_type='solid', fgColor="85EB6A")
            else:
                fill_cell.fill = PatternFill(fill_type='solid', fgColor="FA7080")
                # добавляем замечние в файл КЛАССА
                comment_class_book_sheet.append([file, lesson, teacher, sheets['B' + str(row + 7)].value, up_marks_number,
                                                 fill_cell.value, up_marks_number - fill_cell.value])

                if up_marks_number-fill_cell.value == 1:
                    comment_class_book_cell = comment_class_book_sheet['G' + str(len(comment_class_book_sheet['A']))]
                    comment_class_book_cell.fill = PatternFill(fill_type='solid', fgColor="FFE697")
                elif up_marks_number - fill_cell.value == 2:
                    comment_class_book_cell = comment_class_book_sheet['G' + str(len(comment_class_book_sheet['A']))]
                    comment_class_book_cell.fill = PatternFill(fill_type='solid', fgColor="FFD040")
                elif up_marks_number - fill_cell.value == 3:
                    comment_class_book_cell = comment_class_book_sheet['G' + str(len(comment_class_book_sheet['A']))]
                    comment_class_book_cell.fill = PatternFill(fill_type='solid', fgColor="F9A13E")

                # добавляем замечние в СВОДНЫЙ ОТЧЕТ
                comment_book_sheet.append(
                    [file, lesson, teacher, sheets['B' + str(row + 7)].value, up_marks_number,
                     fill_cell.value, up_marks_number - fill_cell.value])

                if up_marks_number-fill_cell.value == 1:
                    comment_book_cell = comment_book_sheet['G' + str(len(comment_book_sheet['A']))]
                    comment_book_cell.fill = PatternFill(fill_type='solid', fgColor="FFE697")
                elif up_marks_number - fill_cell.value == 2:
                    comment_book_cell = comment_book_sheet['G' + str(len(comment_book_sheet['A']))]
                    comment_book_cell.fill = PatternFill(fill_type='solid', fgColor="FFD040")
                elif up_marks_number - fill_cell.value == 3:
                    comment_book_cell = comment_book_sheet['G' + str(len(comment_book_sheet['A']))]
                    comment_book_cell.fill = PatternFill(fill_type='solid', fgColor="F9A13E")

        #######################################################
        #######################################################
        #######################################################

        # ставим дату и время проверки в ячейку B2
        sheets['B4'] = date

        #print(file, sheet_name, 'Проверен', date)

    book.save('checked\\' + file + ' ПРОВЕРЕНО.xlsx')
    book.close()
    comment_class_book.save('comments\\' + file + ' ЗАМЕЧАНИЯ.xlsx')
    comment_class_book.close()
    comment_book.save('comments\ВСЕ ЗАМЕЧАНИЯ ПО ПРОВЕРКЕ НАКОПЛЯЕМОСТИ ОЦЕНОК.xlsx')
    comment_book.close()