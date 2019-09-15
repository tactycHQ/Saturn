import logging
from tokenizer import Token
logger = logging.getLogger(__name__)

class ASTNode:
    """A generic node in the AST used to compile a cell's formula"""

    def __init__(self, token, cell=None):
        self.token = token
        self.cell = cell
        self._ast = None
        self._parent = None
        self._children = None
        self._descendants = None



    @classmethod
    def create(cls, token, cell=None):
        """Simple factory function"""
        logger.debug("Processing token: {}".format(token))

        if token.type == Token.OPERAND:
            if token.subtype == Token.RANGE:
                logging.debug("Creating RangeNode for: {}".format(token))
                return RangeNode(token, cell)

            else:
                logging.debug("Creating OperandNode for: {}".format(token))
                return OperandNode(token, cell)

        elif token.is_funcopen:
            logging.debug("Creating FunctionNode for: {}".format(token))
            return FunctionNode(token, cell)

        elif token.is_operator:
            logging.debug("Creating OperatorNode for: {}".format(token))
            return OperatorNode(token, cell)

        raise FormulaParserError('Unknown token type: {}'.format(repr(token)))

    def __str__(self):
        return str(self.token.value.strip('('))

    def __repr__(self):
        return '{}<{}>'.format(type(self).__name__,
                               str(self.token.value.strip('(')))

    @property
    def ast(self):
        return self._ast

    @ast.setter
    def ast(self, value):
        self._ast = value

    @property
    def value(self):
        return self.token.value

    @property
    def type(self):
        return self.token.type

    @property
    def subtype(self):
        return self.token.subtype

    @property
    def children(self):
        if self._children is None:
            try:
                args = self.ast.predecessors(self)
            except NetworkXError:
                args = []
            self._children = sorted(args, key=lambda x: self.ast.node[x]['pos'])
        # args.reverse()
        return self._children

    @property
    def descendants(self):
        if self._descendants is None:
            self._descendants = list(
                n for n in self.ast.nodes(self) if n[0] != self)
        return self._descendants

    @property
    def parent(self):
        if self._parent is None:
            self._parent = next(self.ast.successors(self), None)
        return self._parent

    @property
    def emit(self):
        """Emit code"""
        return self.value


class OperatorNode(ASTNode):
    op_map = {
        # convert the operator to python equivalents
        "^": "**",
        "=": "==",
        "<>": "!=",
    }

    @property
    def emit(self):
        xop = self.value

        # Get the arguments
        args = self.children

        op = self.op_map.get(xop, xop)

        if self.type == Token.OP_PRE:
            return self.value + args[0].emit

        parent = self.parent
        # don't render the ^{1,2,..} part in a linest formula
        # TODO: bit of a hack
        if op == "**":
            if parent and parent.value.lower() == "linest(":
                return args[0].emit

        if op == '%':
            ss = '{} / 100'.format(args[0].emit)
        elif op == ' ':
            # range intersection
            ss = '_R_' + ('(str({} & {}))'.format(args[0].emit, args[1].emit)
                          .replace('_R_', '_REF_')
                          .replace('_C_', '_REF_')
                          )
        else:
            if op != ',':
                op = ' ' + op

            ss = '{}{} {}'.format(args[0].emit, op, args[1].emit)

        # avoid needless parentheses
        if parent and not isinstance(parent, FunctionNode):
            ss = "(" + ss + ")"

        return ss


class OperandNode(ASTNode):

    @property
    def emit(self):
        if self.subtype == self.token.LOGICAL:
            return str(self.value.lower() == "true")

        elif self.subtype in ("TEXT", "ERROR") and len(self.value) > 2:
            # if the string contains quotes, escape them
            return '"{}"'.format(self.value.replace('""', '\\"').strip('"'))

        else:
            return self.value


class RangeNode(OperandNode):
    """Represents a spreadsheet cell or range, e.g., A5 or B3:C20"""

    @property
    def emit(self):
        return self._emit()

    def _emit(self, value=None):
        # resolve the range into cells

        sheet = self.cell and self.cell.sheet or ''
        value = value is not None and value or self.value
        if '!' in value:
            sheet = ''
        try:
            addr_str = value.replace('$', '')
            address = AddressRange.create(addr_str, sheet=sheet, cell=self.cell)
        except ValueError:
            # check for table relative address
            table_name = None
            if self.cell:
                excel = self.cell.excel
                if excel and '[' in addr_str:
                    table_name = excel.table_name_containing(self.cell.address)

            if not table_name:
                logging.getLogger('pycel').warning(
                    'Table Name not found: {}'.format(addr_str))
                return '"{}"'.format(NAME_ERROR)

            addr_str = '{}{}'.format(table_name, addr_str)
            address = AddressRange.create(
                addr_str, sheet=self.cell.address.sheet, cell=self.cell)

        if isinstance(address, AddressMultiAreaRange):
            return ', '.join(self._emit(value=str(addr)) for addr in address)
        else:
            template = '_R_("{}")' if address.is_range else '_C_("{}")'
            return template.format(address)


class FunctionNode(ASTNode):
    """AST node representing a function call"""

    """
    A dictionary that maps excel function names onto python equivalents. You
    should only add an entry to this map if the python name is different from
    the excel name (which it may need to be to prevent conflicts with
    existing python functions with that name, e.g., max).

    So if excel defines a function foobar(), all you have to do is add a
    function called foobar to this module.  You only need to add it to the
    function map, if you want to use a different name in the python code.

    Note: some functions (if, pi, and, or, array, ...) are already taken
    care of in the FunctionNode code, so adding them here will have no effect.
    """

    # dict of excel equivalent functions
    func_map = {
        "abs": "x_abs",
        "and": "x_and",
        "atan2": "xatan2",
        "gammaln": "lgamma",
        "if": "x_if",
        "int": "x_int",
        "len": "x_len",
        "max": "xmax",
        "not": "x_not",
        "or": "x_or",
        "min": "xmin",
        "round": "x_round",
        "sum": "xsum",
        "xor": "x_xor",
    }

    def __init__(self, *args):
        super(FunctionNode, self).__init__(*args)
        self.num_args = 0

    def comma_join_emit(self, fmt_str=None, to_emit=None):
        if to_emit is None:
            to_emit = self.children
        if fmt_str is None:
            return ", ".join(n.emit for n in to_emit)
        else:
            return ", ".join(
                fmt_str.format(n.emit) for n in to_emit)

    @property
    def emit(self):
        func = self.value.lower().strip('(')

        if func[0] == func[-1] == '_':
            func = func.upper()
        if func.startswith('_xlfn.'):
            func = func[6:]
        func = func.replace('.', '_')

        # if a special handler is needed
        handler = getattr(self, 'func_{}'.format(func), None)
        if handler is not None:
            return handler()

        else:
            # map to the correct name
            return "{}({})".format(
                self.func_map.get(func, func), self.comma_join_emit())

    @staticmethod
    def func_pi():
        # constant, no parens
        return "pi"

    @staticmethod
    def func_true():
        # constant, no parens
        return "True"

    @staticmethod
    def func_false():
        # constant, no parens
        return "False"

    def func_array(self):
        return '({},)'.format(self.comma_join_emit('({},)'))

    def func_arrayrow(self):
        # simply create a list
        return self.comma_join_emit()

    def func_linest(self):
        func = self.value.lower().strip('(')
        code = '{}({}'.format(func, self.comma_join_emit())

        if not self.cell or not self.cell.excel:
            degree, coef = -1, -1
        else:
            # linests are often used as part of an array formula spanning
            # multiple cells, one cell for each coefficient.  We have to
            # figure out where we currently are in that range.
            degree, coef = get_linest_degree(self.cell)

        # if we are the only linest (degree is one) and linest is nested
        # return vector, else return the coef.
        if func == "linest":
            code += ", degree=%s)" % degree
        else:
            code += ")"

        if not (degree == 1 and self.parent):  # pragma: no branch
            code += "[%s]" % (coef - 1)

        return code

    func_linestmario = func_linest

    @property
    def _build_reference(self):
        if len(self.children) == 0:
            address = '_REF_("{}")'.format(self.cell.address)
        else:
            address = self.children[0].emit
            address = address.replace('_R_', '_REF_').replace('_C_', '_REF_')
            if address.startswith('_REF_(str('):
                address = address[10:-2]
        return address

    def func_row(self):
        return 'row({})'.format(self._build_reference)

    def func_column(self):
        return 'column({})'.format(self._build_reference)

    SUBTOTAL_FUNCS = {
        1: 'average',
        2: 'count',
        3: 'counta',
        4: 'xmax',
        5: 'xmin',
        6: 'product',
        7: 'stdev',
        8: 'stdevp',
        9: 'xsum',
        10: 'var',
        11: 'varp',
    }

    def func_subtotal(self):
        # Excel reference: https://support.office.com/en-us/article/
        #   SUBTOTAL-function-7B027003-F060-4ADE-9040-E478765B9939

        # Note: This does not implement skipping hidden rows.

        func_num = coerce_to_number(self.children[0].emit)
        if func_num not in self.SUBTOTAL_FUNCS:
            if func_num - 100 in self.SUBTOTAL_FUNCS:
                func_num -= 100
            else:
                raise ValueError(
                    "Unknown SUBTOTAL function number: {}".format(func_num))

        func = self.SUBTOTAL_FUNCS[func_num]

        return "{}({})".format(
            func, self.comma_join_emit(fmt_str="{}", to_emit=self.children[1:]))