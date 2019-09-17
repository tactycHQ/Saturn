import xlwings as xw

xw.App().visible = False

wb = xw.Book('Example.xlsx')
sht1 = wb.sheets['Sheet1']
# sht1 = wb.sheets['Sheet1']
sht1.range('A1').value = 20
print(sht1.range('C1').value)
# sht1.range('B2').value = 90