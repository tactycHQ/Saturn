import logging

logger = logging.getLogger(__name__)
from openpyxl import load_workbook, workbook
from ast_ import ASTNode, OperatorNode, RangeNode, OperandNode, FunctionNode
from cell import Cell
from tqdm import tqdm
import networkx as nx

OP_MAP = {
    '+': '+',
    '-': '-',
    '*': '*'
}

FUNC_MAP = {
    'sum': 'xsum'
}


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

        # Load file in read-only mode for faster execution. Need to read twice to extract formulas separately (Is there a better way?)
        self.wb_data_only = load_workbook(filename=self.file, data_only=True, read_only=True)
        logging.info("Values Loaded...")

        self.wb_formulas = load_workbook(filename=self.file, data_only=False, read_only=True)
        logging.info("Formulas Loaded...")

        logging.info("Excel file loaded")

        self.cells = {}
        self.precMap = {}
        self.depMap = {}

        #Parse cells for value
        self.parseCells()

        #Make cells with just RPN for now. AST tree has not been compiled yet
        self.makeCells()

    def parseCells(self):
        '''
        Extraction of values and formulas from each source cell
        @return: Generates a val_dict for values extracted and form_dict for formulas extracted
        '''
        val_dict = {}
        form_dict = {}

        logging.info("Extracting values from cells")
        for sheet in self.wb_data_only.sheetnames:
            for row in self.wb_data_only[sheet].iter_rows():
                for cell in row:
                    if cell.value is not None:
                        address = "{}!{}".format(sheet, cell.coordinate)
                        val_dict[address] = cell.value

        logging.info("Extracting formulas from cells")
        for sheet in self.wb_formulas.sheetnames:
            for row in self.wb_formulas[sheet].iter_rows():
                for cell in row:
                    if cell.value is not None:
                        # logging.debug(cell.coordinate)
                        address = "{}!{}".format(sheet, cell.coordinate)
                        form_dict[address] = cell.value

        logging.info("Retrieved list of {} values".format(len(val_dict)))
        logging.info("Retrieved list of {} formulas".format(len(form_dict)))

        # Test that values extracted match formulas extracted
        assert len(val_dict) == len(form_dict), \
            "Extraction Mismatch - Formulas extracted:{}, Valued extracted:{}".format(len(form_dict), len(val_dict))

        self.val_dict = val_dict
        self.form_dict = form_dict

    def makeCells(self):
        '''
        Wrapper function that instantiates Cell objects for each extracted cell
        @return: Returns self.cells, a list of Cell objects, created
        '''
        for (k, v), (k2, f) in zip(self.val_dict.items(), self.form_dict.items()):
            cell = self.makeCell(k)
        logging.info("--------{} total cell objects created------------".format(len(self.cells)))

        # Now that all cells are created, go ahead and create dependency map
        self.createDepMap()

    def makeCell(self, address):
        '''
        Wrapper function that instantiates 1 Cell object for each extracted cell
        @return: Returns the cell object
        '''
        if address not in self.cells:
            cell = Cell(address)
            logging.info("Not in cellmap, so making cell {}".format(address))
            cell.value = self.val_dict[address]
            cell.formula = self.form_dict[address]
            logging.info("1 cell object created for {} with value {}".format(address, cell.value))

            #Update master cell dictionary
            self.cells[cell.address] = cell

            # Update master precedent dictionary
            self.precMap.update({address: cell.prec})
            return cell

        else:
            logging.info('-----Found {} in cellmap----'.format(address))
            return self.cells[address]

    def getCell(self, address):
        '''
        Returns a cell object at a specified address
        @param address: address
        '''
        try:
            return self.cells[address]
        except Exception as ex:
            logging.info("No cell found")

    def getvalue(self,address):
        return self.getCell(address).value

    def setvalue(self, newvalue, address):
        '''
        Sets value of a cell to a specified new value
        @param newvalue: New value to be set
        @param address: Cell to be set
        '''

        # Set value of cell object to neew value
        cell = self.getCell(address)
        cell.value = newvalue

        # Update master cell list with new cell
        self.cells[address] = cell
        logging.info("Value in cell {} set to {}".format(address, newvalue))

        # Update dependent cells
        self.updateDepCells(address)

    def createDepMap(self):
        '''
        Creates computed list of dependents from the precedent map
        @return: Updates a dictionary of precedents with format {address: [list of dependent addresses]}
        '''
        for address, precedents in self.precMap.items():
            for prec in precedents:
                self.depMap.setdefault(prec, []).append(address)

    def updateDepCells(self, address):
        '''
        Update all dependent cells for a given address
        @param address: Address of source cell that has been changed
        @return:
        '''
        dep_addrs = self.depMap[address]
        logging.info("Updating dependent cells {}".format(dep_addrs))

        for addrs in dep_addrs:
            dep = self.getCell(addrs)
            self.evaluate(dep)



    def evaluate(self, cell):
        '''
        Calculates the formula in a cell
        @param cell: Cell to be calculated
        @return: Value of calculation
        '''

        logging.info("******** Evaluating cell {} with RPN {}".format(cell.address, cell.rpn))
        logging.info("******** {}".format(cell.__dict__))

        #Build AST Tree
        # cell.build_ast()
        tree = cell.rpn

        stack = []

        for node in tree:
            if isinstance(node, OperatorNode):
                arg2 = stack.pop()
                arg1 = stack.pop()
                op = OP_MAP[node.token.value]
                eval_str = '{}{}{}'.format(arg1, op, arg2)
                result = eval(eval_str)
                stack.append(result)

            elif isinstance(node,FunctionNode):
                if node.num_args:
                    args = stack[-node.num_args:]
                    # print(stack)
                    del stack[-node.num_args:]
                    # for i, a in enumerate(args):
                    eval_str = 'lambda x,y,z:x+y+z'
                    result = eval(eval_str)(args[0],args[1],args[2])
                    stack.append(result)
                        # tree.add_node(a, pos=i)
                        # tree.add_edge(a, node)

            else:
                stack.append(self.getCell(node.token.value).value)

        result = stack.pop()





        #Set cell value to new calculated value
        cell.value = result
        return result


        # print(tree.edges)
        # allnodes = tree.nodes
        # print(allnodes)

        # Check if this node is already a hardcode
        # if tree.number_of_nodes() == 1:
        #     for root in tree:
        #         ret = self.getvalue(root.token.value)
        #         logging.info("No child found. Storing value {}".format(ret))
        # if cell.hardcode is True:
        #     logging.info("****Found a hardcode at {} with value {}".format(cell.address, cell.value))
        #     ret = cell.value
        # else:
        #
        #     args = []
        #     ops = []
        #
        #     for node in tree:
        #         logging.info("**** Processing node {} *******".format(node.token.value))
        #         if isinstance(node, OperatorNode):
        #             ops.append(OP_MAP[node.token.value])
        #
        #         if isinstance(node, FunctionNode):
        #             fn = FUNC_MAP[node.token.value]
        #
        #         if isinstance(node, OperandNode):
        #             logging.info("**** Found and evaluating child {}".format(node.token.value))
        #             child_value = self.evaluate(self.getCell(node.token.value))
        #
        #             #Keep track of argument order
        #             args.append(child_value)
        #             pos = tree.node[node]['pos']
        #
        #     eval_str = '{}{}{}'.format(args[0], ops[0], args[1])
        #     logging.info('Python code is: {}'.format(eval_str))
        #     ret = eval(eval_str)
        #
        # #Set cell value to new calculated value
        # cell.value = ret
        #
        # return ret

    # def evaluate2(self, rpn):
    #     '''
    #     Calculates the formula in a cell
    #     @param cell: Cell to be calculated
    #     @return: Value of calculation
    #     '''
    #
    #     logging.info("******** Evaluating cell {} with RPN {}".format(cell.address, cell.rpn))
    #     logging.info("******** {}".format(cell.__dict__))
    #
    #
    #     if cell.hardcode is True:
    #         logging.info("****Found a hardcode at {} with value {}".format(cell.address, cell.value))
    #         ret = cell.value
    #     else:
    #         for i, node in enumerate(rpn):
    #             if isinstance(node,OperatorNode):
    #                 op = OP_MAP[node.token.value]
    #                 arg1 = rpn[i-2]
    #                 arg2 = rpn[i-1]
    #             if isinstance(node, OperandNode):
    #                 child_value = self.evaluate2(self.getCell(node.token.value))
    #
    #     eval_str = '{}{}{}'.format(arg1, ops[0], arg2)
    #     logging.info('Python code is: {}'.format(eval_str))
    #     ret = eval(eval_str)
    #
    #     # Set cell value to new calculated value
    #     cell.value = ret
    #     return ret
    #
    #     # # Set cell value to new calculated value
    #     # cell.value = ret
    #     #
        # return ret

    # def traverseCell(self,cell):
    #     tree = cell.tree
    #
    #     for node in tree:
    #         print(node.token.type)
