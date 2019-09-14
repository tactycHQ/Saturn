from formulaconverter import Formula

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
        if str(excel_formula).startswith('='):
           self._rpn = Formula(self.formula)

