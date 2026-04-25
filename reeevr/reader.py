from tokenizer import Tokenizer
from excelast import ExcelAST
from formula import RTransform
import time

class ExcelReader:
    """
    Read the Excel workbook into an unordered
    dictionary.
    """

    def __init__(self,varconverter,workbook,ignoredsheets):
        self.workbook = workbook
        self.unorderedcode = varconverter.definednames
        self.ignoredsheets = ignoredsheets
        self.varconverter = varconverter
        self.formconverter = RTransform

    def read(self):
        """
        Read in the Excel workbook cell by cell and
        add to the unordered code dictionary
        :return:
        """
        mylist = []
        for sheet in self.workbook.sheetnames:
            print(f"Reading sheet: {sheet}")
            if sheet in self.ignoredsheets:
                continue
            allrows = list(self.workbook[sheet].rows)
            for index, row in enumerate(allrows):
                print(f"Reading row: {index}/{len(allrows)}")

                for indexc,cell in enumerate(row):
                    mylist.append(self.cell_interpret(sheet,cell))

        print(mylist[0])
        program_starts = time.time()
        for val in mylist:
            self.unorderedcode.update(val)
        now = time.time()
        print(f"{now - program_starts}")

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
            celltransform = self.formconverter(cellAST.AST,sheet,cell.coordinate,self.varconverter)
            celltransform.walk(celltransform.tree)

            unorderedcell = {celltransform.outputvarname :[celltransform.code,celltransform.variables,cell.data_type]}

        else:
            raise ValueError("Value type not recognised")

        return unorderedcell

if __name__ == "__main__":

    path = "C:/Users/mo14776/OneDrive - University of Bristol/Documents/Health Economics/REEVER/Examples/Tests for the Excel- R conversion/test_workbook_3.xlsx"
    a=ExcelReader(path,"r")

    a.read()
