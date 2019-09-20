import logging
from fastnumbers import fast_real
from tokenizer import Token
from openpyxl.utils import range_boundaries, get_column_letter

class RPNNode:

    def __init__(self,token):
        self.token = token
        logging.info("{} created with token {} of {}, of type {}, and subtype {}".format(self.__class__.__name__,token.value, type(token.value),
                                                                                  token.type,
                                                                                  token.subtype))

    @classmethod
    def create(self, token, sheet=None):
        if token.type == Token.OPERAND:
            #First lets check if token is a range and have has no sheet, in which case add sheet name to token value
            if token.subtype == Token.RANGE and '!' not in token.value and ':' in token.value:
                token.value = '{}!{}'.format(sheet, token.value)
                return RangeNode(token)

            elif token.subtype == Token.RANGE and '!' not in token.value and ':' not in token.value:
                token.value = '{}!{}'.format(sheet, token.value)
                return CellNode(token)

            # Then lets check if token is a number, in which case update assign float or int to token value
            elif token.subtype == Token.NUMBER:
                token.value = fast_real(token.value)
                return OperandNode(token)

            elif token.subtype == Token.LOGICAL:
                token.value = token.value == 'TRUE'
                return OperandNode(token)

            # Token must be subtype Text, Logical or Error - in which case do nothing
            else:
                return OperandNode(token)

        elif token.is_funcopen:
            return FunctionNode(token)

        elif token.is_operator:
            return OperatorNode(token)

    def __repr__(self):
        if isinstance(self.token.value, str):
            return '{}<{}>'.format(type(self).__name__,
                                   str(self.token.value.strip('(')))
        else:
            return '{}<{}>'.format(type(self).__name__,
                                   self.token.value)

class OperandNode(RPNNode):
    pass


class CellNode(OperandNode):
    def __init__(self, token):
        super().__init__(token)
        self.prec_in_range = token.value
        self.rangeadds = token.value


class RangeNode(OperandNode):
    def __init__(self, token):
        super().__init__(token)
        self.prec_in_range = None
        self.rangeadds = self.rangecells()



    def rangecells(self):
        if self.token.subtype == Token.RANGE:
            split_addrs = self.token.value.split('!')
            sheet = split_addrs[0]
            addresses = split_addrs[1]
            return self.computeRangeCells(addresses,sheet)


    def computeRangeCells(self,rn, sheet):
        x = range_boundaries(rn)

        cols = list(range(min(x[0], x[2]), max(x[0], x[2])))
        cols.append(max(x[0], x[2]))
        cols = list(map(get_column_letter, cols))

        rows = list(range(min(x[1], x[3]), max(x[1], x[3])))
        rows.append(max(x[1], x[3]))

        final = []
        flat_final = []
        for row in rows:
            slice=[]
            for col in cols:
                add_str = '{}!{}{}'.format(sheet, col, row)
                slice.append(add_str)
                flat_final.append(add_str)
            final.append(slice)

        self.prec_in_range = flat_final
        return final


class FunctionNode(RPNNode):

    def __init__(self, token):
        super().__init__(token)
        self.num_args = 0

class OperatorNode(RPNNode):
    pass



