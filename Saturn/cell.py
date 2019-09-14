import logging
logging.basicConfig(level=logging.INFO)
from formulaconverter import Formula

class Cell:
    """
    Class responsible for creating cell objects
    """

    def __init__(self ,address):
        self.address = address
        self._formula = None
        self.value = None
        self._rpn = None
        self.field = None

    def __repr__(self):
        cell_str = 'A:{}, V:{}, F:{}'.format(self.address ,self.value ,self.formula)
        return cell_str

    @property
    def formula(self):
        return self._formula

    @formula.setter
    def formula(self, excel_formula):
        self._formula = excel_formula
        logging.debug("Processing RPN for formula {} at cell {}".format(excel_formula,self))

        try:
            if str(excel_formula).startswith('='):
                self._rpn = Formula(self.formula)
            else:
                logging.debug("Formula does not start with = or +")
                pass
        except Exception as ex:
            logging.error("Unable to parse formula {} at all".format(excel_formula))

