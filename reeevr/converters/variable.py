from openpyxl.utils.cell import cols_from_range
from openpyxl.workbook.defined_name import DefinedName
class VariableConverter:
    """
    Convert the Excel variables following a
    standard method for continuity
    """

    def __init__(self,workbook):
        self.definednames = {}
        self.get_defined_names(workbook)


    def excel_cell_to_variable(self,sheet, coordinate):
        """
        Takes the Excel sheet and coordinate of a cell
        and return an equivalent name.
        :param sheet: name of sheet in workbook
        :param coordinate: coordinates of cell within sheet
        :return: programatic variable name
        """

        if coordinate in self.definednames:
            return coordinate
        else:
            sheetname = str(sheet).strip()
            sheetname = sheetname.replace(" ", "")
            coordinate = coordinate.replace("$", "")
            varname = f'{sheetname}_{coordinate}'
            varname = varname.replace("'","")

        return self.defined_name_comparison(varname)

    def get_defined_names(self,workbook):
        """
        Aquire all global defined names
        for comparison and replacement purposes
        """
        try:
            for definedName in workbook.defined_names.definedName:
                cellName = str(definedName.attr_text).split("!")

                cellLocation = self.excel_cell_to_variable(cellName[0], cellName[1])

                self.definednames[f'{definedName.name}'] = f'{cellLocation}'

        except AttributeError:
            for definedName in workbook.defined_names.items():
                cellName = str(definedName[1].attr_text).split("!")

                cellLocation = self.excel_cell_to_variable(cellName[0], cellName[1])

                self.definednames[f'{definedName[1].name}'] = f'{cellLocation}'

        except:
            raise


    def defined_name_comparison(self, variable):
        """
        Checks if variable is in the defined name list.
        If it is, returns the defined name instead
        of the none-defined name

        :param variable: programatic name of the variable to be compared
        :return: Either the programatic name, or the defined name.
        """

        if variable in self.definednames.values():

            definedname = [key for key, value in self.definednames.items() if variable == value ]
            return definedname[0]

        else:
            return variable


    def excel_range_to_list(self,sheet, rangecoordinates):
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
                list_of_variables.append(self.excel_cell_to_variable(sheet,variable))

        code = str(list_of_variables)
        code = code.replace("'", "")
        return code, list_of_variables


    def variable_string_literal(self,sheet, cell):

        varname = self.excel_cell_to_variable(sheet,cell.coordinate)

        return {varname :[f"\'{cell.value}\'",[]]}


    def variable_numeric_literal(self,sheet,cell):

        varname = self.excel_cell_to_variable(sheet, cell.coordinate)

        return {varname :[cell.value,[]]}


if __name__ == "__main__":
    import openpyxl
    sheet = "sheet 1"
    range = "A1"
    wb = openpyxl.load_workbook("C:/Users/mieha/Documents/REEVER/Test workbooks/test_workbook_3.xlsx", keep_vba=True)
    test = VariableConverter(wb)
    print(test.excel_range_to_list(sheet,range))