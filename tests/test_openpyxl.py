from openpyxl import load_workbook, Workbook
from openpyxl.formula import Tokenizer


#create workbook


#open the .xls file
xlsname = "C:/Users/anubhav/Desktop/Projects/Saturn/Saturn/TestModel_v2.xlsx"
wb = load_workbook(filename=xlsname, data_only=False)
for sheet in wb:
    print(sheet.formula_attributes.items())
    for address, props in sheet.formula_attributes.items():
        if props.get('t') == 'array':
            ref_addr  = props.get('ref')


# tok = Tokenizer(val)
# print(t for t in tok.items)
