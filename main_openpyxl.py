from openpyxl import Workbook, load_workbook
import datetime

date = datetime.datetime.now()

wb_in = load_workbook('11 Я.xlsx')
wb_out = Workbook()

ws_out = wb_out.active

print(wb_in.sheetnames)

i = 0

for sheet in wb_in:
    print(sheet.title)
    i += 1
    ws_out['A'+str(i)].value = sheet.title

# Save the file
wb_out.save("test_out.xlsx")