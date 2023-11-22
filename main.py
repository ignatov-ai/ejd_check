import pandas as pd
        
import warnings
warnings.simplefilter("ignore")

file_list = []
import os
for root, dirs, files in os.walk("."):  
    for filename in files:
        if filename[-4:] == 'xlsx':
            file_list.append(filename)
########print(file_list)
        
#file_list = ['8a-z.xlsx','8b-z.xlsx','8c.xlsx','8ch.xlsx','8e.xlsx','8ja.xlsx','8sh.xlsx','8ndo_samoshilov.xlsx']
out = open('otchet_net_ocenok.csv','w')
out.write('Класс;Предмет;Преподаватель;Ученик\n')
for file in file_list:
    print(file, end = "")

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

        #удаляем столбцы
        current_sheet = current_sheet.drop(['Дата','Тема','Домашнее задание'], axis=1)

        part_1 = current_sheet.iloc[0:49,1:19]
        part_2 = current_sheet.iloc[50:99,2:19]
        part_3 = current_sheet.iloc[100:149,2:19]

        part_1 = part_1.reset_index(drop=True)
        part_2 = part_2.reset_index(drop=True)
        part_3 = part_3.reset_index(drop=True)

        current_sheet_data = pd.concat([part_1,part_2,part_3], axis = 1)
        current_sheet_data = current_sheet_data.fillna('')

        '''
        # create excel writer object
        writer = pd.ExcelWriter('output.xlsx')
        # write dataframe to excel
        all_data.to_excel(writer)
        # save the excel
        writer.save()

        '''
        
        ########print('\n########## '+ sheet + ' ##########')
        ########print('########## '+ predmet[1] + ' ##########')
        ########print('########## '+ teacher[1][:-22] + ' ##########')
        for idx,row in current_sheet_data[1:50].iterrows():
            row_sum = 0
            row_count_mark = 0
            row_propusk = 0
            for i in range(1,50):
                if row[i] == 'н':
                    row_propusk += 1
                elif row[i] == 'Зачёт':
                    row[i] = 5
                elif row[i] != '':
                    if row[i] == 'Зч':
                        row_count_mark += 1
                    else:
                        row[i] = int(row[i])
                        row_sum += row[i]
                        row_count_mark += 1
            '''
            # create excel writer object
            writer = pd.ExcelWriter(sheet+'.xlsx')
            # write dataframe to excel
            current_sheet_data.to_excel(writer)
            # save the excel
            writer.save()
            '''
            
            if row[0] != '':
                if row_sum != 0:
                    #print(idx, row[0], round(row_sum / row_count_mark,2), row_propusk)
                    pass
                else:
                    #print(idx, row[0], 'НЕТ ОЦЕНОК', row_propusk)
                    #out.write(file +';'+ sheet +';'+ file[:-5] +';'+ predmet[1] +';'+ teacher[1][:-22] +';'+ row[0] +'\n')
                    out.write(file[:-5] +';'+ predmet[1] +';'+ teacher[1][:-22] +';'+ row[0] +'\n')
            else:
                break
    print(" - ПРОВЕРЕН")
print('Проверка закончена')
out.close()
