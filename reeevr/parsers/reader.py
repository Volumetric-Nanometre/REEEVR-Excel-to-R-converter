#from openpyxl.formula import Tokenizer
from parsers.tokenizer import Tokenizer
from parsers.excelast import ExcelAST
from converters.formula import RTransform
import time

class ExcelReader:
    """
    Read the Excel workbook into an unordered
    dictionary.
    """

    def __init__(self,varconverter,workbook,outputlang,ignoredsheets):
        self.supportedlanguages = {'r' : RTransform}
        self.workbook = workbook
        self.unorderedcode = varconverter.definednames
        self.ignoredsheets = ignoredsheets #  ['DSA', 'PSA', 'PSA results', 'DSA results']
        self.outputlang = outputlang.lower()
        self.converter = self.language_select()
        self.varconverter = varconverter #VariableConverter(workbook,self.outputlang)


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
        mylist = []
        for sheet in self.workbook.sheetnames:
            print(f"Reading sheet: {sheet}")
            if sheet in self.ignoredsheets:
                continue
            allrows = list(self.workbook[sheet].rows)
            for index, row in enumerate(allrows):#enumerate(self.workbook[sheet].iter_rows()):
                print(f"Reading row: {index}/{len(allrows)}")
                #print(set(row))

                for indexc,cell in enumerate(row):
                    mylist.append(self.cell_interpret(sheet,cell))
                    #self.unorderedcode.update(self.cell_interpret(sheet,cell))
                    #print(f"Reading cell: {indexc}/{len(row)}:{index}/{len(allrows)}")
                #if(index%500):
                #    self.unorderedcode.update(mycurrentdict)
                #    mycurrentdict.clear()
            #self.unorderedcode.update(mycurrentdict)
            #mycurrentdict.clear()
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
            celltransform = self.converter(cellAST.AST,sheet,cell.coordinate,self.varconverter)
            celltransform.walk(celltransform.tree)

            unorderedcell = {celltransform.outputvarname :[celltransform.code,celltransform.variables,cell.data_type]}

        else:
            raise ValueError("Value type not recognised")

        return unorderedcell

if __name__ == "__main__":

    path = "C:/Users/mo14776/OneDrive - University of Bristol/Documents/Health Economics/REEVER/Examples/Tests for the Excel- R conversion/test_workbook_3.xlsx"
    a=ExcelReader(path,"r")

    a.read()
