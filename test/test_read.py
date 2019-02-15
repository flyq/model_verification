import xlrd

workbook = xlrd.open_workbook("./test.xlsx")

worksheet = workbook.sheet_by_name("Sheet1")

nrows = worksheet.nrows

ncols = worksheet.ncols

for i in range(nrows):
    print(worksheet.row_values(i))

for j in range(ncols):
    print(worksheet.col_values(j))

