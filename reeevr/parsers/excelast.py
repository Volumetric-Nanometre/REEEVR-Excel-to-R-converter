from openpyxl.formula import Tokenizer

class ParseNode:

    def __init__(self,type,value):
        self.type = type
        self.value = value

class OperandNode:

    def __init__(self,type,subtype,value):
        self.type = type
        self.subtype = subtype
        self.value = value

class ExpressionNode:

    def __init__(self,type,name,params):
        self.type = type
        self.name = name
        self.params = params

class ProgramNode:

    def __init__(self, type, body):
        self.type = type
        self.body = body
class ExcelAST:
    """
    Handle the conversion of IF functions
    from excel to python.
    e.g. IF(X=Y,A,B) ->
        if X==Y:
            A
        else:
            B
    """
    def __init__(self, tokenstring):

        self.tokenstring =tokenstring.items
        self.count = 0
        self.AST = ProgramNode("Program", self._walk())

    def _walk(self):

        nodelist =[]

        while self.count < len(self.tokenstring):
            token = self.tokenstring[self.count]

            if token.type == 'NUMERIC':
                self.count += 1
                nodelist.append(ParseNode("NumberLiteral",token.value))
            elif token.type == 'WHITE-SPACE':
                self.count += 1
                pass
            elif token.type == 'OPERAND':
                self.count += 1
                nodelist.append(OperandNode("Operand",token.subtype, token.value))
            elif token.type == 'OPERATOR-INFIX':
                self.count += 1
                nodelist.append(ParseNode("InfixOperator", token.value))
            elif token.type == 'FUNC' and token.subtype == 'OPEN':
                self.count += 1
                nodelist.append(ExpressionNode("FunctionCall", token.value, self._walk()))
            elif token.type == 'FUNC' and token.subtype == 'CLOSE':
                self.count += 1
                nodelist.append(ParseNode("FunctionClose", token.value))
                return nodelist
            elif token.type == 'PAREN':
                self.count += 1
                nodelist.append(ParseNode("Paren", token.value))
            elif token.type == 'SEP':
                self.count += 1
                nodelist.append(ParseNode("Separator", token.value))
            else:
                raise ValueError(f"{token.type} not a recognised type")

        return nodelist


if __name__ == "__main__":
    import json

    tokenlist = Tokenizer('= 1 + IF(IF(sheet2!A3 = "yes",sheet4!A3 + 1,sheet!2!A1 + 3),SUM(W10:W20), 10 + sheet!B4 + (3*3)^3')

    a = ExcelAST(tokenlist)
    import jsonpickle
    serialized = jsonpickle.encode(a.AST)
    print( json.dumps(json.loads(serialized), indent=4))
