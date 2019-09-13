from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import AddressCell, AddressRange
import time

# import logging
# logging.basicConfig()
# logging.getLogger().setLevel(logging.DEBUG)

xlsname = "../Test Models/iflix Business Plan Financial Model June 2017.xlsx"

start = time.time()
excel = ExcelCompiler(xlsname)
end = time.time()
print(end-start)
# flex = 'Sheet1!C4'
# cell1 ='Sheet1!C6'
print("Starting")

cell2 ='SSA.Summary!I26'



# print(excel._formula_cells_dict)

# print(excel._evaluate_non_iterative(cell))
# # print(excel._evaluate(cell))
#
# f_cells = excel.formula_cells()
# print(type(f_cells[0]))


# excel.evaluate(cell1)
start = time.time()
print(excel.evaluate(cell2))
end = time.time()
print(end-start)
# print(excel.cell_map)
# # excel.set_value(flex,200)
# # print(excel.evaluate(set))
# # excel.recalculate()
# val1 = excel.evaluate(cell1)
# val2 = excel.evaluate(cell2)
# # excel.recalculate()

# print(val1)
# print(val2)

# form_cells = excel.formula_cells()
# print(len(form_cells))
# print(form_cells[0])

# c = AddressRange(cell1)
# excel.evaluate(cell1)
# excel._gen_graph(cell1)
# excel._make_cells(c)

# print(excel.cell_map)
# print(excel.dep_graph)
# print(excel.validate_calcs())
# print(excel._formula_cells_dict)
