from loader import Loader
import logging

def main():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    xlsname = "../Test Models/TestModel_v1.xlsx"
    excel = Loader(xlsname)
    excel.getCells()

    addr = 'Sheet3!B6'
    cell = excel.makeCell(addr)
    ret = excel.evaluate(cell.tree)
    print("**** Value is: {}".format(ret))


if __name__ ==  '__main__':
    main()


