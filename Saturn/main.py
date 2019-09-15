from loader import Loader
import logging

def main():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)

    xlsname = "../Test Models/TestModel_v1.xlsx"
    excel = Loader(xlsname)
    excel.getCells()

    addr = 'Sheet1!C6'
    excel.makeCell(addr)


if __name__ ==  '__main__':
    main()


