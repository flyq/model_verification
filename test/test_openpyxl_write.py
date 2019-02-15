import openpyxl

workbook = openpyxl.Workbook()

worksheet = workbook.active

data = 100
worksheet.cell(1,2,data)

workbook.save(filename='myfile1.xlsx')