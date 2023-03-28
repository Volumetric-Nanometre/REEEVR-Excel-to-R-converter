from openpyxl.utils.cell import cols_from_range
class VariableConverter:
    """
    Convert the Excel variables following a
    standard method for continuity
    """

    @staticmethod
    def excel_cell_to_variable(sheet, coordinate):
        """
        Takes the Excel sheet and coordinate of a cell
        and return an equivalent name.
        :param sheet: name of sheet in workbook
        :param coordinate: coordinates of cell within sheet
        :return: programatic variable name
        """
        sheetname = str(sheet).strip()
        sheetname = sheetname.replace(" ", "")
        coordinate = coordinate.replace("$", "")

        return f'{sheetname}_{coordinate}'

    @staticmethod
    def defined_name_comparison(definednames, variable):
        """
        Checks if variable is in the defined name list.
        If it is, returns the defined name instead
        of the none-defined name

        :param variable: programatic name of the variable to be compared
        :return: Either the programatic name, or the defined name.
        """

        if variable in definednames.values():

            definedname = str([key for key, value in definednames.items() if variable == value])
            return definedname

        else:
            return variable

    @staticmethod
    def excel_range_to_list(sheet, rangecoordinates):
        """
        Convert an Excel range, eg A3:E10, into a
        list containing all variables. Returns both a
        code string where the entries are variable names,
        and a list of the variables names used.
        :param sheet: sheet name from workbook
        :param rangecoordinates: Excel range coordinates, eg A3:E10
        :return: code string representing the python list of variables
        :return: list of the variable names used in the code string.
        """
        list_of_variables = []
        celllist = cols_from_range(rangecoordinates)
        for row in celllist:
            for variable in row:
                list_of_variables.append(VariableConverter.excel_cell_to_variable(sheet,variable))

        code = str(list_of_variables)
        code = code.replace("'", "")
        return code, list_of_variables


if __name__ == "__main__":
    sheet = "sheet 1"
    range = "A1"
    print(VariableConverter.excel_range_to_list(sheet,range))