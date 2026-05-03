from tokenizer import Tokenizer
from excelast import ExcelAST
from formula import RTransform
import time
import logging

logger = logging.getLogger(__name__)

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
            logger.info(f"Reading sheet: {sheet}")
            if sheet in self.ignoredsheets:
                continue
            allrows = list(self.workbook[sheet].rows)
            for index, row in enumerate(allrows):
                logger.info(f"Reading row: {index}/{len(allrows)}")

                for indexc,cell in enumerate(row):
                    mylist.append(self.cell_interpret(sheet,cell))

        program_starts = time.time()
        for val in mylist:
            self.unorderedcode.update(val)
        now = time.time()
        logger.info(f"Time taken: {now - program_starts}")

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
            logger.error(f"Value type not recognised: {cell.data_type}")

            raise ValueError(f"Value type not recognised: {cell.data_type}")

        return unorderedcell

