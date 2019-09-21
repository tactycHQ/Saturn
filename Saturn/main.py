from loader import Loader
import logging

def main():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    # logger.setLevel(logging.ERROR)

    # xlsname = "TestModel_v2.xlsx"
    # xlsname = "../Test Models/TestModel_v1.xlsx"
    xlsname = "../Test Models/LPI_Consolidated Model vInvestor.xlsx"

    excel = Loader(xlsname)

    set_addr = 'Quarterly!E4'
    addr = 'Quarterly!Q53'
    # formula = '=SUM(E4:E6)+10'
    excel.setvalue(0.2,set_addr)
    # excel.setformula(formula,'Sheet1!E7')
    logging.info(">>>>>>>>>>>Starting evaluation")

    print("Final value at {} : {}".format(addr,excel.getvalue(addr)))


if __name__ ==  '__main__':
    import traceback
    import sys
    traceback.print_exc(file=sys.stdout)
    main()


