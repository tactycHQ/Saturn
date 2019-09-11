from pycel.excelcompiler import ExcelCompiler
# import logging
# logging.basicConfig()
# logging.getLogger().setLevel(logging.DEBUG)

xlsname="../TestModel_v1.xlsx"
excel = ExcelCompiler(xlsname)
flex = 'Sheet1!C4'
cell1 ='Sheet1!C6'
cell2 ='Sheet1!C8'

# print(excel._formula_cells_dict)

# print(excel._evaluate_non_iterative(cell))
# # print(excel._evaluate(cell))
#
# f_cells = excel.formula_cells()
# print(type(f_cells[0]))

excel.evaluate(flex)
excel.evaluate(cell1)
excel.evaluate(cell2)
# excel.set_value(flex,200)
# print(excel.evaluate(set))
# excel.recalculate()
val1 = excel.evaluate(cell1)
val2 = excel.evaluate(cell2)
# excel.recalculate()
print(val1)
print(val2)
print(excel.validate_calcs())
# print(excel._formula_cells_dict)
