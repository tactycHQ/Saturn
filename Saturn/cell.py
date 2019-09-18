import logging
import traceback
logger = logging.getLogger(__name__)
from formula import ExceltoPython
from ast_ import ASTNode, OperatorNode, RangeNode, OperandNode, FunctionNode
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
        self.dep = None
        self.hardcode = None
        self._formula = None
        self.row_field = None
        self.col_field = None
        self.rpn = None
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

        # try:
        if str(excel_formula).startswith(('=','+')):
           f = ExceltoPython(self)
           self.rpn = f.rpn_formula
           self.hardcode = False
           self.dep = f.dep
        else:
            logging.debug("Formula does not start with = or +. Creating a hardcode cell")
            self.tree = DiGraph()
            tok = Token(self.address,Token.LITERAL,None)
            self.tree.add_node(OperandNode(tok))
            self.hardcode = True
            pass
        # except Exception as ex:
        #     print(traceback.format_exc())
        #     logging.error("Unable to parse formula at {} with formula as {}".format(self.address,excel_formula))

    # def traverse(self, tree):
    #
    #     code = []
    #
    #     root = list(tree.nodes()).pop()
    #
    #     for child in list(tree.predecessors(root)):
    #         if isinstance(child,FunctionNode):
    #             self.traverse(child)
    #         elif isinstance(child,RangeNode):
    #             self.traverse(child)
    #     make_code(children)
    #     code.append(make_cod)
    #
    #     return code

    def calculate(self):
        lam = 'lambda {}:{}'.format(args, python_code)
        print(lam)


