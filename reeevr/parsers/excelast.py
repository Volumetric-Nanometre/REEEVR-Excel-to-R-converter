from openpyxl.formula import Tokenizer

class ParseNode:

    def __init__(self,type,value):
        self.type = type
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
                nodelist.append(ParseNode("Operand", token.value))
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

#    def if_function_constructor(self):
#        """
#        Method to output the runcode for an if statement test in a
#        self-contained method.
#
#        :return: code string containing runcode for if statement
#        """
#
#        indent = '    '
#        code = ""
#        code += f'def {self.funcname}_IF({self.LHS},{self.RHS},{self.A},{self.B}):\n'
#        code += f'{indent}if {self.LHS} {self.comparator} {self.RHS}:\n'
#        code += f'{indent}{indent}return {self.A}\n'
#        code += f'{indent}else:\n'
#        code += f'{indent}{indent}return {self.B}\n'
#
#        return code


"""
given sheet1!A1 = 1 + IF(sheet2!A3 = "yes",sheet4!A3 + 1,sheet2!A1 + 3)


Need to return

def funcname(LHS, RHS,A,B):
    if LHS [comparator] RHS:
        return A
    else:
        return B

i.e
def inner_sheet1A1_IF(sheet2_A3, "yes",sheet4_A3 +1,sheet2_A1 +3):
    if LHS [comparator] RHS:
        return A
    else:
        return B


def outer_sheet1_A1_IF(inner_sheet1A1_IF, "",sum(W10:W20),sheet2_A1 +3):
    if LHS [comparator] RHS:
        return A
    else:
        return B
        
sheet1_A1 =outer_sheet1_A1_IF()

"""


if __name__ == "__main__":
    import json

    tokenlist = Tokenizer('= 1 + IF(IF(sheet2!A3 = "yes",sheet4!A3 + 1,sheet!2!A1 + 3),SUM(W10:W20), 10 + sheet!B4 + (3*3)^3')

    a = ExcelAST(tokenlist)
    import jsonpickle
    serialized = jsonpickle.encode(a.AST)
    print( json.dumps(json.loads(serialized), indent=4))

    #print(a.if_function_constructor("sheet2!A3","==",'"yes"',"sheet4!A3 + 1","sheet!2!A1 + 3","sheet1!A1"))