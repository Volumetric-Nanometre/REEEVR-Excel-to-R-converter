import openpyxl
from openpyxl.formula import Tokenizer
from reeevr.converters.variable import VariableConverter
from reeevr.parsers.excelast import ExcelAST
from reeevr.converters.pythonconverter.formula import PythonTransform
from reeevr.converters.rconverter.formula import RTransform
class ExcelReader:
    """
    Read the Excel workbook into an unordered
    dictionary.
    """

    def __init__(self,workbookpath,outputs,outputlang):
        self.supportedlanguages = {'python': PythonTransform, 'r' : RTransform}

        self.workbookpath = workbookpath
        self.workbook = openpyxl.load_workbook(self.workbookpath)
        self.unorderedcode = {}
        self.ignoredsheets = ['DSA', 'PSA']
        self.outputlang = outputlang.lower()
        self.converter = self.language_select()
        self.varconverter = VariableConverter(self.workbook,self.outputlang)
        self.outputcells = self.expand_outputs(outputs)


    def expand_outputs(self,outputs):

        expandedOutputs = []
        for output in outputs:
            sheet = output[0]
            range = output[1]


    def language_select(self):

        try:
            return self.supportedlanguages[self.outputlang]

        except KeyError:
            raise KeyError(f"Selected output language not supported: {self.outputlang}")

    def read(self):
        """
        Read in the Excel workbook cell by cell and
        add to the unordered code dictionary
        :return:
        """
        for sheet in self.workbook.sheetnames:

            if sheet in self.ignoredsheets:
                continue

            for row in self.workbook[sheet].iter_rows():
                for cell in row:
                    self.unorderedcode.update(self.cell_interpret(sheet,cell))


    def cell_interpret(self,sheet,cell):

        # {'variable' : ["codeified string", ['list','of','contained','vars'],cell.data_type]
        #unorderedcell

        if cell.data_type == "n":
            unorderedcell = self.varconverter.variable_numeric_literal(sheet,cell)

        elif cell.data_type == "s":
            unorderedcell = self.varconverter.variable_string_literal(sheet,cell)

        elif cell.data_type == "f":

            tokenizer = Tokenizer(cell.value)

            cellAST = ExcelAST(tokenizer)
            celltransform = self.converter(cellAST.AST,sheet,cell.coordinate,self.varconverter)
            celltransform.walk(celltransform.tree)

            unorderedcell = {celltransform.outputvarname :[celltransform.code,celltransform.variables,cell.data_type]}

        else:
            raise ValueError("Value type not recognised")

        return unorderedcell

if __name__ == "__main__":

    path = "C:/Users/mo14776/OneDrive - University of Bristol/Documents/Health Economics/REEVER/Examples/Tests for the Excel- R conversion/test_workbook_3.xlsx"
    a=ExcelReader(path,"Python")

    a.read()
