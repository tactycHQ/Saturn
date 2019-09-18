from tokenizer import Token

class ASTNode:

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

    # def __str__(self):
    #     return str(self.token.value.strip('('))
    def __repr__(self):
        return '{}<{}>'.format(type(self).__name__,
                               str(self.token.value.strip('(')))

class OperandNode(ASTNode):
    pass



class RangeNode(OperandNode):
    pass


class FunctionNode(ASTNode):
    pass

class OperatorNode(ASTNode):
    pass



