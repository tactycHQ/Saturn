from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import AddressCell, AddressRange
import logging
logging.basicConfig()
# logging.getLogger().setLevel(logging.DEBUG)


def main():

    xlsname = "C:/Users/anubhav/Desktop/Projects/Saturn/Saturn/TestModel_v2.xlsx"

    excel = ExcelCompiler(xlsname)
    cell2 ='Sheet1!B8'
    excel.evaluate(cell2)
    print("----------")
    excel.set_value('Sheet1!B7',1)
    print("----------")
    print(excel.evaluate(cell2))

if __name__ == '__main__':
        main()


