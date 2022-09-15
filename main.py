from openpyxl import Workbook, load_workbook

wb = load_workbook('Grades.xlsx')
ws = wb.active
print(ws['A2'].value)
