from tokenizer import Token
from openpyxl.utils import range_boundaries, get_column_letter

class RPNNode:

    def __init__(self,token):
        self.token = token

    @classmethod
    def create(self, token):
        if token.type == Token.OPERAND:
            if token.subtype == Token.RANGE:
                return RangeNode(token)
            else:
                return OperandNode(token)

        elif token.is_funcopen:
            return FunctionNode(token)

        elif token.is_operator:
            return OperatorNode(token)

    def __repr__(self):
        return '{}<{}>'.format(type(self).__name__,
                               str(self.token.value.strip('(')))

class OperandNode(RPNNode):
    pass



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



