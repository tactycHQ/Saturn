import logging
logging.basicConfig(level=logging.ERROR)
from openpyxl import load_workbook, workbook
from cell import Cell
from tqdm import tqdm


class Loader:
    """
    Class responsible for injecting an xlsx file and converting the file into Cell objects.
    1. Injects an xlsx file
    2. Uses openpyxl to convert extract formulas and value from each cell
    3. Instantiates Cell objects for each cell extracted
    """

    def __init__(self, file):
        '''
        Initializes a loader object and sets the self.file param to injected file
        @param file: xlsx file
        '''
        self.file = file
        logging.info("Loading excel file...")

        #Load file in read-only mode for faster execution. Need to read twice to extract formulas separately (Is there a better way?)
        self.wb_data_only = load_workbook(filename=self.file, data_only=True, read_only=True)
        logging.info("Values Loaded...")

        self.wb_formulas = load_workbook(filename=self.file, data_only=False, read_only=True)
        logging.info("Formulas Loaded...")

        logging.info("Excel file loaded")

    def getCells(self):
        '''
        Extraction of values and formulas from each source cell
        @return: Generates a val_dict for values extracted and form_dict for formulas extracted
        '''
        val_dict={}
        form_dict = {}

        logging.info("Extracting values from cells")
        for sheet in self.wb_data_only.sheetnames:
            for row in self.wb_data_only[sheet].iter_rows():
                for cell in row:
                    if cell.value is not None:
                        logging.debug(cell.coordinate)
                        address = "{}!{}".format(sheet,cell.coordinate)
                        val_dict[address] = cell.value

        logging.info("Extracting formulas from cells")
        for sheet in self.wb_formulas.sheetnames:
            for row in self.wb_formulas[sheet].iter_rows():
                for cell in row:
                    if cell.value is not None:
                        logging.debug(cell.coordinate)
                        address = "{}!{}".format(sheet, cell.coordinate)
                        form_dict[address] = cell.value

        logging.info("Retrieved list of {} values".format(len(val_dict)))
        logging.info("Retrieved list of {} formulas".format(len(form_dict)))

        #Test that values extracted match formulas extracted
        assert len(val_dict) == len (form_dict),\
            "Extraction Mismatch - Formulas extracted:{}, Valued extracted:{}".format(len(form_dict),len(val_dict))

        self.val_dict = val_dict
        self.form_dict = form_dict

    def makeCells(self):
        '''
        Wrapper function that instantiates Cell objects for each extracted cell
        @return: Returns self.cells, a list of Cell objects, created
        '''
        self.cells = []
        for (k, v), (k2, f) in zip(self.val_dict.items(), self.form_dict.items()):
            logging.debug("Making {} with formula {}".format(k,f))
            cell= Cell(k)
            cell.value = self.val_dict[k]
            cell.formula = self.form_dict[k]
            self.cells.append(cell)
        logging.info("Values and formulas combined into one cell for {} cells".format(len(self.cells)))

        return self.cells



if __name__ ==  '__main__':

    # xlsname = "Book3.xlsx"
    xlsname = "../Test Models/TestModel_v1.xlsx"
    excel = Loader(xlsname)
    excel.getCells()
    excel.makeCells()
    # print(excel.cells)

