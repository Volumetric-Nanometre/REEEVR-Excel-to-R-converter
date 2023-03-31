
class TranslatorNode:

    def __init__(self,type,value):
        self.type = type
        self.value = value

class CalleeNode:
    def __init__(self, type,name):
        self.type = type
        self.name = name
class TranslatorExpressionNode:

    def __init__(self, type,name,callee, arguments):
        self.type = type
        self.name = name
        self.callee = callee
        self.arguments = arguments

class TranslatorProgramNode:

    def __init__(self, type, body):
        self.type = type
        self.body = body
class PythonTranslator:

    def __init__(self, ExcelAST):
        self.ExcelAST = ExcelAST
        self.PythonAST = TranslatorProgramNode("Program", [])
        self.position = self.PythonAST.body
        self.transformer(self.ExcelAST)

    def transformer(self,originalAST):
        def NumberLiteral(self,node):

            self.position.append(TranslatorNode('NumericLiteral',node.value))

        def FunctionCall(self,node,parent):

             expression = TranslatorExpressionNode('FunctionCall',CalleeNode('Identifier',node.name),[])

             prevposition = self.position

             self.position = expression.arguments
             if parent.type != 'FunctionCall':
                 expression = TranslatorProgramNode('ExpressionStatement',expression)

             prevposition.append(expression)

        self.traverse(originalAST,{'NumberLiteral':NumberLiteral,'FunctionCall':FunctionCall})

    def traverse(self,ast,visitors):
        def walk_node(node, parent):
            try:
                method = visitors[node.type]

                if method:
                    method(node, parent)
            except:
                pass
            if node.type == 'Program':
                walk_nodes(node.body,node)
            elif node.type == 'FunctionCall':
                walk_nodes(node.params,node)

        def walk_nodes(nodes, parent):
            for node in nodes:
                walk_node(node, parent)

        walk_node(ast,None)


if __name__ == "__main__":
    import json
    from openpyxl.formula import Tokenizer
    from excelast import ExcelAST
    tokenlist = Tokenizer('= 1 + IF(IF(sheet2!A3 = "yes",sheet4!A3 + 1,sheet!2!A1 + 3),SUM(W10:W20), 10 + sheet!B4 + (3*3)^3')

    excelAST = ExcelAST(tokenlist)
    import jsonpickle

    #print(a.if_function_constructor("sheet2!A3","==",'"yes"',"sheet4!A3 + 1","sheet!2!A1 + 3","sheet1!A1"))

    test = PythonTranslator(excelAST.AST)

    serialized = jsonpickle.encode(test.PythonAST)
    print( json.dumps(json.loads(serialized), indent=4))
