from loader import Loader
import logging

def main():
    logger = logging.getLogger()
    # logger.setLevel(logging.DEBUG)
    logger.setLevel(logging.ERROR)

    xlsname = "TestModel_v2.xlsx"
    # xlsname = "../Test Models/TestModel_v1.xlsx"
    # xlsname = "../Test Models/LP_LatAm Model vF.xlsx"

    excel = Loader(xlsname)
    set_addr = 'Sheet3!D3'
    addr = 'Sheet3!E5'
    # c = excel.getCell(addr)
    # val = excel.evaluate(c)
    # print(val)

    # addr = 'Returns!D5'
    excel.setvalue(20,set_addr)
    print(excel.getvalue(addr))


if __name__ ==  '__main__':
    main()


