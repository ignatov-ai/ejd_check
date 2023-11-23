import pandas as pd
        
import warnings
warnings.simplefilter("ignore")

file_list = []
import os
for root, dirs, files in os.walk("."):  
    for filename in files:
        if filename[-4:] == 'xlsx':
            file_list.append(filename)
print(file_list)        
        
out = open('spisok_uchiteley.csv','w')
out.write('Класс;Предмет;Преподаватель\n')
for file in file_list:
    tabs = pd.ExcelFile(file).sheet_names

    #print(tabs)
    
    for sheet in tabs:
        current_sheet = pd.read_excel(io=file,engine='openpyxl',sheet_name=sheet, na_filter=False)

        #сохраняем предмет и учителя
        if current_sheet.shape[0] < 50:
            predmet = current_sheet.iloc[39,20].split(', ')
            teacher = current_sheet.iloc[41,20].split(': ')
        else:
            predmet = current_sheet.iloc[89,20].split(', ')
            teacher = current_sheet.iloc[91,20].split(': ')

        current_sheet = current_sheet.drop(0:20, axis=1)

        part_1 = current_sheet.iloc[0:49,1:19]
        part_2 = current_sheet.iloc[50:99,2:19]
        part_3 = current_sheet.iloc[100:149,2:19]

        part_1 = part_1.reset_index(drop=True)
        part_2 = part_2.reset_index(drop=True)
        part_3 = part_3.reset_index(drop=True)

        current_sheet_data = pd.concat([part_1,part_2,part_3], axis = 1)
        current_sheet_data = current_sheet_data.fillna('')
        
        print(file[:-5] +' | '+ predmet[-1] +' | '+ teacher[1][:-22])
        out.write(file[:-5] +';'+ predmet[-1] +';'+ teacher[1][:-22] +'\n')
    print()
print('Список составлен')
out.close()
