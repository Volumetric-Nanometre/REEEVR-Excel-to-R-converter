
class FormulaeConverter:
    """
    Stores the formula that can be converted
    via a simple pick and place switch out.
    e.g SQRT() -> math.sqrt()

    We need to return the new code, and a list
    of any libraries that are used to support
    the new code e.g. math.sqrt() requires math
    """

    exceltopythonformulae = {'SUM': ['numpy.sum', 'numpy'],
                             'AVERAGE': ['numpy.mean', 'numpy'],
                             'SQRT': ['numpy.sqrt', 'numpy']}

    @staticmethod
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

    FUNC = "AVERAGE"
    print(FormulaeConverter.formulae_excel_to_python(FUNC))