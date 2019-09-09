from pycel.excelcompiler import ExcelCompiler

xlsname="../TestModel_v1.xlsx"
excel = ExcelCompiler(xlsname)
cell ='Sheet1!C2'

# print(excel._formula_cells_dict)

print(excel._evaluate_non_iterative(cell))
# print(excel._evaluate(cell))

print(excel.formula_cells())
print(excel.cell_map)
