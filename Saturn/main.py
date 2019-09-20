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

    set_addr = 'Sheet1!C3'
    addr = 'Sheet1!C9'
    # formula = '=SUM(E4:E6)+10'
    excel.setvalue(0.0002,set_addr)
    # excel.setformula(formula,'Sheet1!E7')
    print("Final value: {}".format(excel.getvalue(addr)))




if __name__ ==  '__main__':
    main()


