import xlwt

book = xlwt.Workbook(encoding="utf-8",style_compression=0)

sheet = book.add_sheet("Sheet1", cell_overwrite_ok=True)

data1 = 'beijing'
sheet.write(1,0,data1)

book.save('test\\test_write.xlsx')