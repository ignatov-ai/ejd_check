from openpyxl import load_workbook
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

print(uchebniy_plan)

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

# выбирается файл для обработки
for file in file_list:
    file = file.split('.xlsx')[0]
    book = load_workbook("data\\" + file + '.xlsx')

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
            students_count += 1

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
        print(up_class_index, file.lower())

        up_lesson_index = 0
        for i in range (len(uchebniy_plan)):
            if uchebniy_plan[i][0] == lesson.lower():
                up_lesson_index = i
        print(up_lesson_index, lesson.lower())

        up_lesson_nagruzka = uchebniy_plan[up_class_index][up_lesson_index]
        print(file.lower(), lesson.lower(), up_lesson_nagruzka)



        #######################################################
        #######################################################
        #######################################################

        # ставим дату и время проверки в ячейку B2
        sheets['B4'] = date

        print(sheet_name,'Проверен', date)

    book.save('checked/' + file + ' ПРОВЕРЕНО.xlsx')
    book.close()