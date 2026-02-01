from traverse import TraverseTree


class RTransform(TraverseTree):
    """
    Class that deals explicitly in Excel -> R
    transformations only.
    """


    def __init__(self,  excelast, sheet, coordinate,varconverter):

        super().__init__(excelast, sheet, coordinate,varconverter)

        self.simple_transform = [ "SQRT(",
                                  "EXP(",
                                  "ABS(",
                                  "MAX(",
                                  "MIN(",
                                  "ROUND(",
                                  "LOG10("

                                ]

        self.adaptive_transform = { "RAND(": self.RAND,
                                    "LN(": self.NATURALLOG,
                                    "_xlfn.BETA.INV(": self.BETAINV,
                                    "BETAINV(": self.BETAINV,
                                    "_xlfn.NORM.INV(": self.NORMINV,
                                    "NORMINV(":self.NORMINV,
                                    "NORMSINV(": self.NORMINV,
                                    "_xlfn.NORM.S.INV(": self.NORMINV,
                                    "_xlfn.LOGNORM.INV(": self.LOGNORMINV,
                                    "LOGINV(": self.LOGNORMINV,
                                    "IF(" : self.IF,
                                    "INT(" : self.INT,
                                    "LEN(" : self.LEN,
                                    "UPPER(" : self.UPPER,
                                    "LOWER(" : self.LOWER,
                                    "PI(" : self.PI,

                                    }

        self.reeevr_transform = [ "SUM(", "_xlfn.GAMMA.INV(", "GAMMAINV(", "_xlfn.BINOM.INV(", "_xlfn.CONCAT(",
                                  "CONCATENATE(", "INDEX(", "CHOOSE(", "IFERROR(", "AND(", "_xlfn.IFS(", "OFFSET(",
                                  "POWER(", "PRODUCT(", "FLOOR(", "CEILING(", "LEFT(", "RIGHT(", "MID(",
                                  "_xlfn.NORM.DIST(", "NORMDIST(", "_xlfn.NORM.S.DIST(", "NORMSDIST(",
                                  "_xlfn.GAMMA.DIST(", "GAMMADIST(", "_xlfn.LOGNORM.DIST(", "LOGNORMDIST(",
                                  "_xlfn.BINOM.DIST(", "BINOMDIST(", "_xlfn.BETA.DIST(", "BETADIST(",
                                  "_xlfn.WEIBULL.DIST(", "WEIBULL(", "LOG(", "NEGBINOMDIST(", "_xlfn.NEGBINOM.DIST(",
                                  "CHIDIST(", "CHIINV(", "_xlfn.CHISQ.DIST(", "_xlfn.CHISQ.DIST.RT(",
                                  "_xlfn.CHISQ.INV(", "_xlfn.CHISQ.INV.RT(", "_xlfn.EXPON.DIST(", "EXPONDIST(",
                                  "_xlfn.F.DIST(", "_xlfn.F.DIST.RT(", "_xlfn.F.INV(", "_xlfn.F.INV.RT(", "POISSON(",
                                  "_xlfn.POISSON.DIST(", "_xlfn.T.DIST(", "_xlfn.T.DIST.2T(", "_xlfn.T.DIST.RT(",
                                  "_xlfn.T.INV(", "_xlfn.T.INV.2T(", "TDIST(", "MEDIAN(", "AVERAGE(", "STDEV(",
                                  "_xlfn.STDEV.S(", "STDEVP(", "_xlfn.STDEV.P(", "VAR(", "_xlfn.VAR.S(", "VARP(",
                                  "_xlfn.VAR.P(", "COUNT(", "COUNTA(", "COUNTIF("
                                  ]


    def formula_converter(self,input,params):

        if input in self.simple_transform:
            return self.simple_convert(input,params)

        elif input in self.reeevr_transform:
            return self.reeevr_convert(input,params)

        elif input in self.adaptive_transform.keys():
            return self.adaptive_transform[input](params)
        else:
            simplesyntax = self.walk(params)
            with open("missing-func.log", "a+") as f:
                f.write(f"{input} - {simplesyntax}\n")
            return f"UNKNOWN_FUNCTION_SEE_LOG_FILE({''.join(simplesyntax)}"

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
    def simple_convert(self,input,params):
        """
        Performs a simple conversion by lowering the function,
        thus returning the correct output without additional work
        """
        simplesyntax = self.walk(params)
        return f"{input.lower()}{''.join(simplesyntax)}"

    def reeevr_convert(self,input,params):
        """
        Palms off the conversion to the reeevr R package
        thus returning the correct output with external work
        """
        simplesyntax = self.walk(params)
        return f"excel_{input.lower()}{''.join(simplesyntax)}"

    def NATURALLOG(self, params):
        """
        Returns the natural log of the paramaters
        """
        simplesyntax = self.walk(params)
        return f"log({''.join(simplesyntax)}"


    def RAND(self,params):
        """
        Generates a random number with even distribution
        of the form 0<=x<=1
        """
        simplesyntax = self.walk(params)
        return f"excel_rand(numberOfRuns, converter_validate{''.join(simplesyntax)}"

    def BETAINV(self,params):
        """
        Return function to calculate the inverse of the beta cumulative probability density function
        Must be of the form:
        dinvbeta(x,α, ß, log=FALSE)
        where x is the locatiin vector
        """
        simplesyntax = self.walk(params)
        return f"qbeta({''.join(simplesyntax)}"

    def LOGNORMINV(self, params):
        """
        Returns the inverse of the lognorm cumulative distribution for the specified mean and standard deviation
        """
        simplesyntax = self.walk(params)
        return f"qlnorm({''.join(simplesyntax)}"

    def NORMINV(self,params):
        """
        Returns the inverse of the normal cumulative distribution for the specified mean and standard deviation
        """
        simplesyntax = self.walk(params)
        return f"qnorm({''.join(simplesyntax)}"


    def INT(self, params):
        """
        INT in Excel works as the floor() function in R.
        e.g.
        INT(3.1) -> 3, floor(3.1) -> 3
        INT(3.9) -> 3, floor(3.9) -> 3
        INT(-3.1) -> -4, floor(-3.1) -> -4
        INT(-3.9) -> -4, floor(-3.9) -> 4
        """
        simplesyntax = self.walk(params)
        return f"floor({''.join(simplesyntax)}"

    def LEN(self, params):
        """
        Returns the number of chars in a string
        """
        simplesyntax = self.walk(params)
        return f"nchar({''.join(simplesyntax)}"

    def UPPER(self, params):
        """
        Capitalise all letters in string
        """
        simplesyntax = self.walk(params)
        return f"toupper({''.join(simplesyntax)}"

    def LOWER(self, params):
        """
        Capitalise all letters in string
        """
        simplesyntax = self.walk(params)
        return f"tolower({''.join(simplesyntax)}"

    def PI(self, params):
        """
        Return pi
        """
        return f"pi"



if __name__ == "__main__":

    from formula import Tokenizer
    from excelast import ExcelAST
    tokenlist = Tokenizer('= 1 + IF(IF(sheet10!A1 = "yes", AVERAGE(A10:A20), 23), SUM(B10:V20),50) + '
                          'IF(IF(OMG!A1 = "yes", SUM(A10:A20), 70),"shit",\'My stuff\'!A1) +20 + SUM(A10:A20)')

    excelAST = ExcelAST(tokenlist)
    import jsonpickle

    serialized = jsonpickle.encode(excelAST.AST)
    #print(json.dumps(json.loads(serialized), indent=4))

    test = RTransform(excelAST.AST, "Sheet10", "C10")
    test.walk(test.tree)
    print(test.code)
