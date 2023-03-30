from openpyxl.formula import Tokenizer
from reeevr.converters.variable import VariableConverter

class FormulaeConverter:
    """
    Stores the formula that can be converted
    via a simple pick and place switch out.
    e.g SQRT() -> math.sqrt()

    We need to return the new code, and a list
    of any libraries that are used to support
    the new code e.g. math.sqrt() requires math
    """

    exceltopythonformulae = {')': [')', ''],
                             'SUM(': ['numpy.sum(', 'numpy'],
                             'AVERAGE(': ['numpy.mean(', 'numpy'],
                             'SQRT(': ['numpy.sqrt(', 'numpy'],
                             'IF(': ['if(', '']}

    @staticmethod
    def general_formula_convert(sheet,cell):

        varname = VariableConverter.excel_cell_to_variable(sheet,cell.coordinate)

        tokenString = Tokenizer(str(cell.value))

        codestring =""
        codevariables = []
        for token in tokenString.items:

            if token.type == 'OPERAND' and token.subtype == 'RANGE':
                operandvar = str()
                if '!' in token.value:
                    varcomponents = token.value.split("!")

                    if ':' in varcomponents[1]:
                        operandvar, variablelist = VariableConverter.excel_range_to_list(sheet,varcomponents[1])
                        codevariables += variablelist

                    else:
                        operandvar = VariableConverter.excel_cell_to_variable(varcomponents[0],varcomponents[1])
                        codevariables.append(operandvar)

                elif ':' in token.value:
                    operandvar, variablelist = VariableConverter.excel_range_to_list(sheet,token.value)
                    codevariables+=variablelist

                else:
                    operandvar = VariableConverter.excel_cell_to_variable(sheet,token.value)
                    codevariables.append(operandvar)

                codestring+=operandvar


            elif token.type == 'FUNC':
                codestring += FormulaeConverter.formulae_excel_to_python(token.value)[0]
            else:
                codestring += str(token.value)



        return {varname : [codestring,codevariables]}

    def formulae_excel_to_python(formulae):
        """
        Enter the Excel formulae string and returns
        the python formulae equivalent
        :param formulae: Excel formulae e.g SQRT
        :return: ['python formulae','library required to run'] e.g ['math.sqrt', 'math']
        """



        try:
            return FormulaeConverter.exceltopythonformulae[formulae]
        except KeyError:
            raise KeyError("Formulae not recognised in dictionary.")



if __name__ == '__main__':

    FUNC = "AVERAGE("
    print(FormulaeConverter.formulae_excel_to_python(FUNC))