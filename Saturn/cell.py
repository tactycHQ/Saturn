import logging
logging.basicConfig(level=logging.INFO)
from formulaconverter import Formula

class Cell:
    """
    Class responsible for creating cell objects
    """

    def __init__(self ,address):
        """
        Each cell object is initalized with:
        - address, excel formula, value, rpn formula, col_field
        @param address:
        """
        self.address = address
        self.value = None
        self._formula = None
        self._rpn = None
        self.row_field = None
        self.col_field = None

    def __repr__(self):
        '''Represents a cell object by outputting the address, value and excel formula'''
        cell_str = 'A:{}, V:{}, F:{}'.format(self.address ,self.value ,self.formula)
        return cell_str

    @property
    def formula(self):
        return self._formula

    @formula.setter
    def formula(self, excel_formula):
        '''
        If excel formula is set, immediately create corresponding rpn formula
        @param excel_formula: excel formula as a string
        @return: rpn formula
        '''
        self._formula = excel_formula
        logging.debug("Processing RPN for formula {} at cell {}".format(excel_formula,self))

        try:
            if str(excel_formula).startswith(('=','+')):
                self._rpn = Formula(self.formula)
            else:
                logging.debug("Formula does not start with = or +")
                pass
        except Exception as ex:
            logging.error("Unable to parse formula {} at all".format(excel_formula))

