from parsers.traverse import TraverseTree


class PythonTransform(TraverseTree):
    """
    Class that deals explicitly in Excel -> Python
    transformations only.
    """


    def __init__(self,  excelast, sheet, coordinate, varconverter):

        super().__init__(excelast, sheet, coordinate, varconverter)
        self.function_transformations = {"IF(": self.IF,
                                         "SUM(": self.SUM,
                                         "AVERAGE(": self.AVERAGE,
                                         "SQRT(": self.SQRT}

    def IF(self, params):

        expressionname = f"{self.outputvarname}_if_{self.depth}_{self.count}"
        simplesyntax = self.walk(params)

        if_state = "".join(simplesyntax)
        if_state = if_state[:-1]
        if_state = if_state.split('%sep%')

        if '=' in if_state[0]:
            if ('<' or '>') in if_state[0]:
                pass
            else:
                if_state[0] = if_state[0].replace("=", "==")

        indent = '    '
        code = f"def {expressionname}():\n" \
               f"{indent}if {if_state[0]}:\n" \
               f"{indent}{indent}return {if_state[1]}\n" \
               f"{indent}else:\n" \
               f"{indent}{indent}return {if_state[2]}\n"

        self.code += code

        return f'{expressionname}()'

    def SUM(self, params):
        simplesyntax = self.walk(params)
        return f"numpy.sum({''.join(simplesyntax)}"

    def AVERAGE(self, params):
        """
        Arithmetic mean
        """
        simplesyntax = self.walk(params)
        return f"numpy.mean({''.join(simplesyntax)}"

    def SQRT(self,params):
        """
        Return sqrt of value
        """
        simplesyntax = self.walk(params)
        return f"numpy.sqrt({''.join(simplesyntax)}"

if __name__ == "__main__":

    from openpyxl.formula import Tokenizer
    from parsers.excelast import ExcelAST
    tokenlist = Tokenizer('= 1 + IF(IF(sheet10!A1 = "yes", AVERAGE(A10:A20), 23), SUM(B10:V20),50) + '
                          'IF(IF(OMG!A1 = "yes", SUM(A10:A20), 70),"shit",\'My stuff\'!A1) +20 + SUM(A10:A20)')

    excelAST = ExcelAST(tokenlist)
    import jsonpickle

    serialized = jsonpickle.encode(excelAST.AST)
    #print(json.dumps(json.loads(serialized), indent=4))

    test = PythonTransform(excelAST.AST, "Sheet10", "C10","Sheet10_C10")
    test.walk(test.tree)
    print(test.code)
