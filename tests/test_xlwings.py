import xlwings as xw

wb = xw.Book('Example.xlsx')
sht1 = wb.sheets['Sheet1']
sht1 = wb.sheets['Sheet1']
sht1.range('A1').value = 15
# sht1.range('B2').value = 90