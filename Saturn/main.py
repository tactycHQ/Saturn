from loader import Loader


if __name__ ==  '__main__':
    xlsname = "../Test Models/TestModel_v1.xlsx"
    excel = Loader(xlsname)
    excel.getCells()
    excel.makeCells()
