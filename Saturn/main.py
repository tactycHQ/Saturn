from loader import Loader
import logging

def main():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    # logger.setLevel(logging.ERROR)

    xlsname = "TestModel_v2.xlsx"
    # xlsname = "../Test Models/TestModel_v1.xlsx"
    # xlsname = "../Test Models/Arraytest v1.xlsx"

    excel = Loader(xlsname)
    print(excel.cells.get('Sheet1!B11'))


    set_addr = 'Sheet1!C4'
    excel.setvalue(5, set_addr)


    logging.info(">>>>>>>>>>>Starting evaluation")
    addr = 'Sheet1!B11'
    print("Final value at {} : {}".format(addr,excel.getvalue(addr)))


if __name__ ==  '__main__':
    import traceback
    import sys
    traceback.print_exc(file=sys.stdout)
    main()


