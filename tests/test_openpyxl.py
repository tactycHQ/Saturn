from openpyxl import load_workbook
from openpyxl.formula import Tokenizer

#open the .xls file
xlsname="../TestModel_v1.xlsx"
wb = load_workbook(filename=xlsname, data_only=False)
sheet = wb['Sheet1']
val = sheet['C2'].value
print(val)

tok = Tokenizer(val)
print(t for t in tok.items)
