from openpyxl.formula.tokenizer import Tokenizer as _Tokenizer
from openpyxl.formula.tokenizer import Token as _Token
import re

"""
This module contains a tokenizer for Excel formulae.

The tokenizer is based on the Javascript tokenizer found at
http://ewbi.blogs.com/develops/2004/12/excel_formula_p.html written by Eric
Bachtal

This tokenizer has also been amended to reflect the changes by pycel at
https://github.com/dgorissen/pycel
"""

class Tokenizer(_Tokenizer):
    """Amend openpyxl tokenizer"""

    def __init__(self, formula):
        super(Tokenizer, self).__init__(formula)
        self.items = self._items()

    def _items(self):
        """Convert to use our Token"""
        t = [None] + [Token.from_token(t) for t in self.items] + [None]

        # convert or remove unneeded whitespace
        tokens = []
        for prev_token, token, next_token in zip(t, t[1:], t[2:]):
            if token.type != Token.WSPACE or not prev_token or not next_token:
                # drop unary +
                if not token.matches(type_=Token.OP_PRE, value='+'):
                    tokens.append(token)

            elif (
                prev_token.matches(type_=Token.FUNC, subtype=Token.CLOSE) or
                prev_token.matches(type_=Token.PAREN, subtype=Token.CLOSE) or
                prev_token.type == Token.OPERAND
            ) and (
                next_token.matches(type_=Token.FUNC, subtype=Token.OPEN) or
                next_token.matches(type_=Token.PAREN, subtype=Token.OPEN) or
                next_token.type == Token.OPERAND
            ):
                # this whitespace is an intersect operator
                tokens.append(Token(token.value, Token.OP_IN, Token.INTERSECT))

        return tokens

class Token(_Token):
    """Amend openpyxl token"""

    INTERSECT = "INTERSECT"
    ARRAYROW = "ARRAYROW"

    class Precedence:
        """Small wrapper class to manage operator precedence during parsing"""

        def __init__(self, precedence, associativity):
            self.precedence = precedence
            self.associativity = associativity

        def __lt__(self, other):
            return (self.precedence < other.precedence or
                    self.associativity == "left" and
                    self.precedence == other.precedence
                    )

    precedences = {
        # http://office.microsoft.com/en-us/excel-help/
        #   calculation-operators-and-precedence-HP010078886.aspx
        ':': Precedence(8, 'left'),
        ' ': Precedence(8, 'left'),  # range intersection
        ',': Precedence(8, 'left'),
        'u': Precedence(7, 'left'),  # unary operator
        '%': Precedence(6, 'left'),
        '^': Precedence(5, 'left'),
        '*': Precedence(4, 'left'),
        '/': Precedence(4, 'left'),
        '+': Precedence(3, 'left'),
        '-': Precedence(3, 'left'),
        '&': Precedence(2, 'left'),
        '=': Precedence(1, 'left'),
        '<': Precedence(1, 'left'),
        '>': Precedence(1, 'left'),
        '<=': Precedence(1, 'left'),
        '>=': Precedence(1, 'left'),
        '<>': Precedence(1, 'left'),
    }

    @classmethod
    def from_token(cls, token, value=None, type_=None, subtype=None):
        return cls(
            token.value if value is None else value,
            token.type if type_ is None else type_,
            token.subtype if subtype is None else subtype
        )

    @property
    def is_operator(self):
        return self.type in (Token.OP_PRE, Token.OP_IN, Token.OP_POST)

    @property
    def is_funcopen(self):
        return self.subtype == Token.OPEN and self.type in (
            Token.FUNC, Token.ARRAY, Token.ARRAYROW)

    def matches(self, type_=None, subtype=None, value=None):
        return ((type_ is None or self.type == type_) and
                (subtype is None or self.subtype == subtype) and
                (value is None or self.value == value))

    @property
    def precedence(self):
        assert self.is_operator
        return self.precedences[
            'u' if self.type == Token.OP_PRE else self.value]

if __name__ == "__main__":

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
        '=SUM(123 + SUM(456) + (45<6))+456+789',
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
        tok =Tokenizer(i)
        for t in tok.items:
            print("Pretty printed:\n", t.value, t.type, t.subtype)