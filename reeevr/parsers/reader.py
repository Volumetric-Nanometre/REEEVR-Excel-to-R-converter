import openpyxl
from openpyxl.formula import Tokenizer

class ExcelReader:
    """
    Read the Excel workbook into an unordered
    dictionary.
    """

    def __init__(self,workbookpath):
        self.workbookpath = workbookpath
        self.workbook = openpyxl.load_workbook(self.workbookpath)
        self.unorderedcode = {}
        self.ignoredsheets = ['DSA', 'PSA']

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



    def cell_interpret(self,cell):

        # {'variable' : ["codeified string", ['list','of','contained','vars']]
        unorderedcell = {}
        tokenString = Tokenizer(str(cell.value))

        for
