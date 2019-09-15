import logging
logger = logging.getLogger(__name__)
from tokenizer import Tokenizer, Token
from ast_tree import ASTNode


"""
The make_rpn function has been taken from pycel's implementation of the _parse_to_rpn
The Tokenizer and Token is also taken from pycel's implementation - which itself is an amendment to openpyxl's Tokenizer and Token implementation 

The algorithm is as follows:
1. Tokenize an Excel Formula into list of tokens
2. Pass list of tokens to shunting yard algo to create a rpn formula
3. Create an AST comprising of [] [] [] - the AST is created by the shunting yard algorithm itself.

https://www.reddit.com/r/learnprogramming/comments/3cybca/how_do_i_go_about_building_an_ast_from_an_infix/ 
"""


class ExceltoPython():
    '''
    Class responsible for converting an excel formula to its RPN notation
    '''

    def __init__(self, xl_formula):
        self.rpn_formula = self.make_rpn(xl_formula)


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

        for token in tokens:
            if token.type == token.OPERAND:
                output.append(ASTNode.create(token))
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
                    output.append(ASTNode.create(stack.pop()))

                if not len(were_values):
                    raise FormulaParserError(
                        "Mismatched or misplaced parentheses")

                were_values.pop()
                arg_count[-1] += 1
                were_values.append(False)

            elif token.is_operator:

                while stack and stack[-1].is_operator:
                    if token.precedence < stack[-1].precedence:
                        output.append(ASTNode.create(stack.pop()))
                    else:
                        break

                stack.append(token)

            elif token.subtype == token.OPEN:
                assert token.type in (token.FUNC, token.PAREN, token.ARRAY)
                stack.append(token)

            elif token.subtype == token.CLOSE:

                while stack and stack[-1].subtype != Token.OPEN:
                    output.append(ASTNode.create(stack.pop()))

                if not stack:
                    raise FormulaParserError(
                        "Mismatched or misplaced parentheses")

                stack.pop()

                if stack and stack[-1].is_funcopen:
                    f = ASTNode.create(stack.pop())
                    f.num_args = arg_count.pop() + int(were_values.pop())
                    output.append(f)

            else:
                assert token.type == token.WSPACE, \
                    'Unexpected token: {}'.format(token)

        while stack:
            if stack[-1].subtype in (Token.OPEN, Token.CLOSE):
                raise FormulaParserError("Mismatched or misplaced parentheses")

            output.append(ASTNode.create(stack.pop()))

        return output

if __name__ == '__main__':

    # Test inputs
    inputs = [
        # # Simple test formulae
        # '=1+3+5',
        # '=3 * 4 + 5',
        # '=50',
        # '=1+1',
        # '=$A1',
        # '=$B$2',
        # '=SUM(B5:B15)',
        # '=SUM(B5:B15,D5:D15)',
        # '=SUM(B5:B15 A7:D7)',
        # '=SUM(sheet1!$A$1:$B$2)',
        # '=[data.xls]sheet1!$A$1',
        # '=SUM((A:A 1:1))',
        # '=SUM((A:A,1:1))',
        # '=SUM((A:A A1:B1))',
        # '=SUM(D9:D11,E9:E11,F9:F11)',
        # '=SUM((D9:D11,(E9:E11,F9:F11)))',
        # '=IF(P5=1.0,"NA",IF(P5=2.0,"A",IF(P5=3.0,"B",IF(P5=4.0,"C",IF(P5=5.0,"D",IF(P5=6.0,"E",IF(P5=7.0,"F",IF(P5=8.0,"G"))))))))',
        # '={SUM(B2:D2*B3:D3)}',
        # '=SUM(123 + SUM(456) + (45<6))+456+789',
        '=Sheet2!C3+1'
        # '=AVG(((((123 + 4 + AVG(A1:A2))))))',

        # E. W. Bachtal's test formulae
        # '=IF("a"={"a","b";"c",#N/A;-1,TRUE}, "yes", "no") &   "  more ""test"" text"',
        # '=+ AName- (-+-+-2^6) = {"A","B"} + @SUM(R1C1) + (@ERROR.TYPE(#VALUE!) = 2)',
        # '=IF(R13C3>DATE(2002,1,6),0,IF(ISERROR(R[41]C[2]),0,IF(R13C3>=R[41]C[2],0, IF(AND(R[23]C[11]>=55,R[24]C[11]>=20),R53C3,0))))',
        # '=IF(R[39]C[11]>65,R[25]C[42],ROUND((R[11]C[11]*IF(OR(AND(R[39]C[11]>=55, ' +
        # 'R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")),R[44]C[11],R[43]C[11]))+(R[14]C[11] ' +
        # '*IF(OR(AND(R[39]C[11]>=55,R[40]C[11]>=20),AND(R[40]C[11]>=20,R11C3="YES")), ' +
        # 'R[45]C[11],R[43]C[11])),0))',
    ]

    for i in inputs:
        print("========================================")
        print("Formula:     " + i)
        tok = ExceltoPython(i)
