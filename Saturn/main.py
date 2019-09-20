from loader import Loader
import logging

def main():
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    # logger.setLevel(logging.ERROR)

    xlsname = "TestModel_v2.xlsx"
    # xlsname = "../Test Models/TestModel_v1.xlsx"
    # xlsname = "../Test Models/LP_LatAm Model vF.xlsx"

    excel = Loader(xlsname)
    # print(excel.precMap)
    # print(excel.depMap)

    set_addr = 'Sheet1!B3'
    addr = 'Sheet1!D10'
    # formula = '=SUM(E4:E6)+10'
    excel.setvalue(5,set_addr)
    # excel.setformula(formula,'Sheet1!E7')
    logging.info(">>>>>>>>>>>Starting evaluation")

    print("Final value: {}".format(excel.getvalue(addr)))


if __name__ ==  '__main__':
    main()


