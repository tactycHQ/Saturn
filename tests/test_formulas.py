import formulas


#create workbook


#open the .xls file
xlsname = "../Test Models/fund.xlsx"
excel = formulas.ExcelModel().loads(xlsname).finish()
excel.calculate()

# tok = Tokenizer(val)
# print(t for t in tok.items)
