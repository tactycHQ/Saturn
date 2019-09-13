from tokenizer import Tokenizer, Token
from pycel.excelformula import ExcelFormula

class Cell:

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
        self._rpn = self.parseFormula(self.formula)

    def parseFormula(self, expression):
        e = ExcelFormula(self.formula)
        # print(e._parse_to_rpn(self.formula))
