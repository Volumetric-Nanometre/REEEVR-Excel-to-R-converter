from reeevr.parsers.traverse import TraverseTree


class RTransform(TraverseTree):
    """
    Class that deals explicitly in Excel -> Python
    transformations only.
    """


    def __init__(self,  excelast, sheet, coordinate,varconverter):

        super().__init__(excelast, sheet, coordinate,varconverter)
        self.function_transformations = {"IF(": self.IF,
                                         "SUM(": self.SUM,
                                         "AVERAGE(": self.AVERAGE,
                                         "SQRT(": self.SQRT,
                                         "RAND(": self.RAND,
                                         "LN(": self.NATURALLOG,
                                         "EXP(": self.EXP,
                                         "_xlfn.BETA.INV(": self.BETAINV,
                                         "BETAINV(": self.BETAINV,
                                         "_xlfn.NORM.INV(": self.NORMINV,
                                         "_xlfn.STDEV.S(": self.STDEVSAMPLE,
                                         "_xlfn.CONCAT(": self.CONCAT,
                                         "GAMMAINV(": self.GAMMAINV,
                                         "CHOOSE(": self.CHOOSE,
                                         "COUNTA(": self.IGNORE,
                                         "COUNTIF(": self.IGNORE,
                                         "IFERROR(": self.IFERROR}

    def IF(self, params):

        expressionname = f"{self.outputvarname}_if_{self.depth}_{self.count}"
        simplesyntax = self.walk(params)

        if_state = "".join(simplesyntax)
        if_state = if_state[:-1]
        if_state = if_state.split('%sep%')

        if '=' in if_state[0]:
            if ('<' or '>' or '!') in if_state[0]:
                pass
            else:
                if_state[0] = if_state[0].replace("=", "==")

        indent = '    '
        code = f"{expressionname} <- function(){{\n" \
               f"{indent}if({if_state[0]}){{\n" \
               f"{indent}{indent}return({if_state[1]})\n" \
               f"{indent}}}\n" \
               f"{indent}else{{\n" \
               f"{indent}{indent}return({if_state[2]})\n" \
               f"{indent}}}\n" \
               f"}}\n"

        self.code += code

        return f'{expressionname}()'

    def SUM(self, params):
        simplesyntax = self.walk(params)
        return f"sum({''.join(simplesyntax)}"

    def SQRT(self,params):
        """
        Return sqrt of value
        """
        simplesyntax = self.walk(params)
        return f"sqrt({''.join(simplesyntax)}"

    def NATURALLOG(self, params):
        """
        Returns the natural log of the paramaters
        """
        simplesyntax = self.walk(params)
        return f"log({''.join(simplesyntax)}"

    def EXP(self, params):
        """
        Returns the exp() of the value
        """
        simplesyntax = self.walk(params)
        return f"exp({''.join(simplesyntax)}"

    def AVERAGE(self, params):
        """
        Arithmetic mean
        """
        simplesyntax = self.walk(params)
        return f"mean({''.join(simplesyntax)}"

    def RAND(self,params):
        """
        Generates a random number with even distribution
        of the form 0<=x<=1
        """
        simplesyntax = self.walk(params)
        return f"runif(numberOfRuns{''.join(simplesyntax)}"

    def BETAINV(self,params):
        """
        Return function to calculate the inverse of the beta cumulative probability density function
        Must be of the form:
        dinvbeta(x,α, ß, log=FALSE)
        where x is the locatiin vector
        """
        simplesyntax = self.walk(params)
        return f"qbeta({''.join(simplesyntax)}"

    def NORMINV(self,params):
        """
        Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation
        """
        simplesyntax = self.walk(params)
        return f"qnorm({''.join(simplesyntax)}"

    def STDEVSAMPLE(self,params):
        """
        Standard deviation of a sample
        """
        simplesyntax = self.walk(params)
        return f"sd({''.join(simplesyntax)}"

    def GAMMAINV(self,params):
        """
        Inverse cumulative gamma distribution
        """
        simplesyntax = self.walk(params)
        return f"qinvgamma({''.join(simplesyntax)}"

    def CONCAT(self, params):
        """
        Return the concatination of the inputs
        """

        simplesyntax = self.walk(params)
        return f"paste({''.join(simplesyntax)}"

    def CHOOSE(self,params):
        "excelChoose <- function"
        print("CHOOSE - Need to substitute in the correct code")
        simplesyntax = self.walk(params)
        return f"excel_choose({''.join(simplesyntax)}"

    def IFERROR(self,params):
        simplesyntax = self.walk(params)

        code = ''.join(simplesyntax)
        code = code.split("%sep%")[0]

        return f"{code}"

    def IGNORE(self,params):

        simplesyntax = self.walk(params)
        return f"excel_ignore({''.join(simplesyntax)}"

if __name__ == "__main__":

    from openpyxl.formula import Tokenizer
    from reeevr.parsers.excelast import ExcelAST
    tokenlist = Tokenizer('= 1 + IF(IF(sheet10!A1 = "yes", AVERAGE(A10:A20), 23), SUM(B10:V20),50) + '
                          'IF(IF(OMG!A1 = "yes", SUM(A10:A20), 70),"shit",\'My stuff\'!A1) +20 + SUM(A10:A20)')

    excelAST = ExcelAST(tokenlist)
    import jsonpickle

    serialized = jsonpickle.encode(excelAST.AST)
    #print(json.dumps(json.loads(serialized), indent=4))

    test = RTransform(excelAST.AST, "Sheet10", "C10")
    test.walk(test.tree)
    print(test.code)
