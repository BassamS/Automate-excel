from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

wb = load_workbook('Grades.xlsx')
ws = wb.active
# change the value
# ws['A2'].value = 'Sam'

# Saving
wb.save('Grades.xlsx')

# Creating a new sheet
# wb.create_sheet('new sheet')

# print(wb.sheetnames)


for row in range(1, 11):
    for col in range(1, 5):
        char = get_column_letter(col)
        # print(ws[char + str(row)].value)
        ws[char + str(row)] = char + str(row)
