from openpyxl import load_workbook, Workbook
from openpyxl.formula import Tokenizer


#create workbook


#open the .xls file
xlsname = "../Test Models/LP_India Model vF.xlsx"
wb = load_workbook(filename=xlsname, data_only=False)
for sheet in wb.sheetnames:
    val = sheet['C2'].value
    print(val)

# tok = Tokenizer(val)
# print(t for t in tok.items)
