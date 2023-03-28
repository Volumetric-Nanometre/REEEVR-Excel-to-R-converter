
class VariableConverter:
    """
    Convert the Excel variables following a
    standard method for continuity
    """
    @staticmethod
    def defined_name_comparison(variable):
        """
        Checks if variable is in the defined name list.
        If it is, returns the defined name instead
        of the none-defined name

        :param variable: programatic name of the variable to be compared
        :return: Either the programatic name, or the defined name.
        """
        definednames = {}

        if variable in definednames.values():

            definedname = str([key for key, value in definednames.items() if variable == value])
            return definedname

        else:
            return variable

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
    def excel_range_to_list(sheet, rangecoordinates):

        coordlist = rangecoordinates.split(":")

        columns = []

