from openpyxl import load_workbook, workbook
import pandas as pd

class ExcelCompiler:

    def __init__(self):
        return

    def load_excel(self, file, data_only=False):
        wb = load_workbook(filename=file, data_only=data_only)
        return wb

    def getSheets(self, wb):
        return wb.sheetnames

    def getCells(self, wb, wb2):
        cell_df_formula = []

        for sheet in wb.sheetnames:
            for row in wb[sheet].iter_rows():
                for cell in row:
                    c= Cell()
                    c.formula = cell.value
                    c.row = cell.row
                    c.column = cell.column
                    c.coordinate = cell.coordinate
                    cell_df_formula.append(c)

        cell_df_val = []
        for sheet in wb2.sheetnames:
            for row in wb2[sheet].iter_rows():
                for cell in row:
                    c= Cell()
                    c.value = cell.value
                    c.row = cell.row
                    c.column = cell.column
                    c.coordinate = cell.coordinate
                    cell_df_val.append(c)

        all_cells = self.addValuestoCells(cell_df_formula,cell_df_val)
        return all_cells

    def addValuestoCells(self,cell_df_formula, cell_df_val):
        for cell1, cell2 in zip(cell_df_formula, cell_df_val):
            if cell1.coordinate == cell2.coordinate:
                cell1.value = cell2.value
        return cell_df_formula

    def calculateCells(self, all_cells):
        computed_cells = []

        for cell in all_cells:
            value = cell.compute()
        computed_cells.append()

class Cell:

    def __init__(self):
        self.value=None
        self.formula=None
        self.address=None
        self.row=None
        self.columns=None
        self.calc_value=None

if __name__ ==  '__main__':

    xlsname = "../TestModel_v1.xlsx"
    xlc = ExcelCompiler()
    wb_formula = xlc.load_excel(xlsname, data_only=False)
    wb_val = xlc.load_excel(xlsname, data_only=True)
    all_cells = xlc.getCells(wb_formula, wb_val)
    computed_cells = xlc.calculateCells(all_cells)

