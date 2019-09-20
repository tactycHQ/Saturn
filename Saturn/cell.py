import logging
import traceback
logger = logging.getLogger(__name__)
from rpnnode import RPNNode, OperatorNode, RangeNode, OperandNode, FunctionNode
from networkx.classes.digraph import DiGraph
from tokenizer import Tokenizer, Token

class Cell:
    """
    Class responsible for creating cell objects from source addresses
    """

    def __init__(self ,address):
        """
        Each cell object is initalized with:
        - address, excel formula, value, rpn formula, col_field
        @param address:
        """
        self.address = address
        self.value = None
        self.prec = []
        self._formula = None
        self.rpn = []
        self.tree = None

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
        If excel formula is set, this TRIGGERS creation of rpn formula and tree
        @param excel_formula: excel formula as a string
        @return: rpn formula
        '''
        self._formula = excel_formula
        logging.debug("Processing RPN for formula {} at cell {}".format(excel_formula,self))

        #First check if formula starts with correct operator
        if str(excel_formula).startswith(('=','+')):
           self.rpn = self.make_rpn(excel_formula)

           # creates list of precedents (who do I depend on)
           self.createPrec()

        # This means formula must be a hardcode
        else:
            logging.debug("Formula does not start with = or +. Creating a hardcode cell")
            tok = Token(self.address,Token.LITERAL,None)
            self.rpn.append(OperandNode(tok))

        logging.info("RPN is: {}".format(self.rpn))

    def createPrec(self):
        for node in self.rpn:
            if isinstance(node,RangeNode):
                self.prec.extend(node.prec_in_range)

    def make_node(self, token):
        sheet = self.address.split('!')[0]
        if token.type == Token.OPERAND and token.subtype == Token.RANGE and '!' not in token.value:
            token.value = '{}!{}'.format(sheet, token.value)
            return RPNNode.create(token)
        else:
            return RPNNode.create(token)

    def make_rpn(self, expression):
        """
        Parse an excel formula expression into reverse polish notation

        Core algorithm taken from wikipedia with varargs extensions from
        http://www.kallisti.net.nz/blog/2008/02/extension-to-the-shunting-yard-
            algorithm-to-allow-variable-numbers-of-arguments-to-functions/
        """

        """
                Parse an excel formula expression into reverse polish notation

                Core algorithm taken from wikipedia with varargs extensions from
                http://www.kallisti.net.nz/blog/2008/02/extension-to-the-shunting-yard-
                    algorithm-to-allow-variable-numbers-of-arguments-to-functions/
                """

        lexer = Tokenizer(expression)

        logging.info("Tokens created succesfully")

        # amend token stream to ease code production
        tokens = []
        for token, next_token in zip(lexer.items, lexer.items[1:] + [None]):

            if token.matches(Token.FUNC, Token.OPEN):
                tokens.append(token)
                token = Token('(', Token.PAREN, Token.OPEN)

            elif token.matches(Token.FUNC, Token.CLOSE):
                token = Token(')', Token.PAREN, Token.CLOSE)

            elif token.matches(Token.ARRAY, Token.OPEN):
                tokens.append(token)
                tokens.append(Token('(', Token.PAREN, Token.OPEN))
                tokens.append(Token('', Token.ARRAYROW, Token.OPEN))
                token = Token('(', Token.PAREN, Token.OPEN)

            elif token.matches(Token.ARRAY, Token.CLOSE):
                tokens.append(token)
                token = Token(')', Token.PAREN, Token.CLOSE)

            elif token.matches(Token.SEP, Token.ROW):
                tokens.append(Token(')', Token.PAREN, Token.CLOSE))
                tokens.append(Token(',', Token.SEP, Token.ARG))
                tokens.append(Token('', Token.ARRAYROW, Token.OPEN))
                token = Token('(', Token.PAREN, Token.OPEN)

            elif token.matches(Token.PAREN, Token.OPEN):
                token.value = '('

            elif token.matches(Token.PAREN, Token.CLOSE):
                token.value = ')'

            tokens.append(token)

        output = []
        stack = []
        were_values = []
        arg_count = []

        # shunting yard start
        for token in tokens:
            if token.type == token.OPERAND:
                output.append(self.make_node(token))
                if were_values:
                    were_values[-1] = True

            elif token.type != token.PAREN and token.subtype == token.OPEN:

                if token.type in (token.ARRAY, Token.ARRAYROW):
                    token = Token(token.type, token.type, token.subtype)

                stack.append(token)
                arg_count.append(0)
                if were_values:
                    were_values[-1] = True
                were_values.append(False)

            elif token.type == token.SEP:

                while stack and (stack[-1].subtype != token.OPEN):
                    output.append(self.make_node(stack.pop()))

                if not len(were_values):
                    raise FormulaParserError(
                        "Mismatched or misplaced parentheses")

                were_values.pop()
                arg_count[-1] += 1
                were_values.append(False)

            elif token.is_operator:

                while stack and stack[-1].is_operator:
                    if token.precedence < stack[-1].precedence:
                        output.append(self.make_node(stack.pop()))
                    else:
                        break

                stack.append(token)

            elif token.subtype == token.OPEN:
                assert token.type in (token.FUNC, token.PAREN, token.ARRAY)
                stack.append(token)

            elif token.subtype == token.CLOSE:

                while stack and stack[-1].subtype != Token.OPEN:
                    output.append(self.make_node(stack.pop()))

                if not stack:
                    raise FormulaParserError(
                        "Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].is_funcopen:
                    f = self.make_node((stack.pop()))
                    f.num_args = arg_count.pop() + int(were_values.pop())
                    output.append(f)

            else:
                assert token.type == token.WSPACE, \
                    'Unexpected token: {}'.format(token)

        while stack:
            if stack[-1].subtype in (Token.OPEN, Token.CLOSE):
                raise FormulaParserError("Mismatched or misplaced parentheses")

            output.append(self.make_node(stack.pop()))
        return output




# if __name__ == "__main__":

