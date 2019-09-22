from pycel.excelcompiler import ExcelCompiler
from pycel.excelutil import AddressCell, AddressRange
import logging
logging.basicConfig()
logging.getLogger().setLevel(logging.DEBUG)


def main():

    xlsname = "C:/Users/anubhav/Desktop/Projects/Saturn/Saturn/TestModel_v2.xlsx"

    excel = ExcelCompiler(xlsname)
    cell2 ='Sheet1!B11'
    excel.evaluate(cell2)
    logging.info("\n\n-----SET VALUE-----")
    excel.set_value('Sheet1!C4','5')
    logging.info("\n\n----EVALUATE------")
    ret = excel.evaluate(cell2)
    print("Final value is: {}".format(ret))

if __name__ == '__main__':
        main()


