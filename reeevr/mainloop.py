from reader import ExcelReader
from codegen import CodeGen
from outputs import ROutputs
from variable import VariableConverter
import openpyxl
import os
import io
import traceback
import logging

logger = logging.getLogger(__name__)

def print_to_string(*args, **kwargs):
    output = io.StringIO()
    print(*args, file=output, **kwargs)
    contents = output.getvalue()
    output.close()
    return contents

class MainLoop:

    def __init__(self, gui=False, progressbar = None, guitextbrowser = None):
        self.path = ""
        self.folder = ""
        self.file = ""
        self.testOutput = []
        self.ignoredsheets = []
        self.costs = []
        self.effectiveness = []
        self.treatments = []
        self.willingnesstopay = int()
        self.gui = gui
        self.completionstate = 0
        self.maxcomplete = 9
        self.progressbar = progressbar
        self.guitextbrowser=guitextbrowser
        self.text = ""
        self.BCEA = False

        logger.debug("Main loop initialised")

    def set_vals(self, path, folder, testoutput, ignoredsheets, costs, effectiveness, treatments, willingtopay, BCEA):
        self.path = path
        self.folder = folder
        print(self.folder)
        self.BCEA = BCEA

        self.file = os.path.basename(path)
        self.file = os.path.splitext(self.file)[0]
        self.file = self.file + ".R"

        self.testOutput = testoutput.replace("!", ",")
        self.testOutput = self.testOutput.split(",")

        self.testOutput = list(zip([ val.strip() for index, val in enumerate(self.testOutput) if index%2 == 0],
                                   [val.strip() for index, val in enumerate(self.testOutput) if index%2 == 1]))
        self.output(self.testOutput)

        self.ignoredsheets = ignoredsheets.split(",")
        self.ignoredsheets = [val.strip() for val in self.ignoredsheets]
        self.output(self.ignoredsheets)

        self.costs = costs.replace("!", ",")
        self.costs = self.costs.split(",")
        self.costs = list(zip([ val.strip() for index, val in enumerate(self.costs) if(index%2 == 0)],
                              [val.strip() for index, val in enumerate(self.costs) if index%2 == 1]))
        self.output(self.costs)

        self.effectiveness = effectiveness.replace("!", ",")
        self.effectiveness = self.effectiveness.split(",")
        self.effectiveness = list(zip([ val.strip() for index, val in enumerate(self.effectiveness) if index%2 == 0],
                                      [val.strip() for index, val in enumerate(self.effectiveness) if index%2 == 1]))
        self.output(self.effectiveness)

        self.treatments = treatments.replace("!", ",")
        self.treatments = self.treatments.split(",")
        self.treatments = [val.strip() for val in self.treatments]
        self.output(self.treatments)

        self.willingnesstopay = willingtopay.replace("!", ",")
        self.willingnesstopay = self.willingnesstopay.split(",")
        self.willingnesstopay = list(zip([val.strip() for index, val in enumerate(self.willingnesstopay) if index % 2 == 0],
                                         [val.strip() for index, val in enumerate(self.willingnesstopay) if index % 2 == 1]))
        self.output(self.willingnesstopay)

        logger.debug("Values initialised")


    def output(self,*args,**kwargs):
        if self.gui :
            self.text += print_to_string(*args, **kwargs)

            self.guitextbrowser.setPlainText(self.text)

    def update_progress(self):
        self.completionstate += 1
        if self.gui :
            self.progressbar.setProperty("value", self.completionstate/self.maxcomplete*100)

    def run(self):
        self.text = ""
        try:
            os.remove("missing-func.log")
        except:
            pass

        try:
            os.remove("missing-cells.txt")
        except:
            pass

        try:
            logger.info("Beginning run")
            logger.info("Open workbook ... ")

            self.output("Open workbook ... ",end="")
            workbook = openpyxl.load_workbook(self.path)
            self.output("[SUCCESS]")
            logger.info("[SUCCESS]")
            self.update_progress()

            logger.info("Initialise variable converter ... ")
            self.output("Initialise variable converter ... ",end="")
            varconverter = VariableConverter(workbook,self.ignoredsheets)
            self.output("[SUCCESS]")
            logger.info("[SUCCESS]")
            self.update_progress()

            logger.info("Initialise outputs ... ")
            self.output("Initialise outputs ... ",end="")
            outputs = ROutputs(varconverter,  self.folder, self.testOutput, self.costs, self.effectiveness, self.treatments, self.willingnesstopay, self.BCEA)
            self.output("[SUCCESS]")
            logger.info("[SUCCESS]")
            self.update_progress()

            logger.info("Create reader ... ")
            self.output("Create reader ... ",end="")
            a = ExcelReader(varconverter, workbook, self.ignoredsheets)
            self.output("[SUCCESS]")
            logger.info("[SUCCESS]")
            self.update_progress()

            logger.info("Read code ... ")
            self.output("Read code ... ",end="")
            a.read()
            self.output("[SUCCESS]")
            logger.info("[SUCCESS]")
            self.update_progress()

            logger.info("Generate unordered code ... ")
            self.output("Generate unordered code ... ",end="")
            b = CodeGen(varconverter, a.unorderedcode, outputs, self.folder, self.file)
            b.second_pass()
            self.output("[SUCCESS]")
            logger.info("[SUCCESS]")
            self.update_progress()

            logger.info("Order code ... ")
            self.output("Order Code ... ",end="")
            b.order_code_snippets()
            self.output("[SUCCESS]")
            logger.info("[SUCCESS")
            self.update_progress()

            logger.info("Prune code ... ")
            self.output("Prune Code ... ",end="")
            b.cyclic_prune()
            self.output("[SUCCESS]")
            logger.info("[SUCCESS]")
            self.update_progress()

            logger.info("Output code ... ")
            self.output("Output Code ... ",end="")
            b.generate_code()
            self.output("[SUCCESS]")
            logger.info("[SUCCESS]")
            self.update_progress()

            logger.info("Run complete")
        except:
            logger.error("[FAILED]")
            self.output("[FAILED]\n")
            self.output(traceback.print_exc())
            logger.error(f"{traceback.print_exc()}")

            self.completionstate = 0
            self.progressbar.setStyleSheet("QProgressBar"
                                             "{"
                                             "background-color : red;"
                                             "border : 1px"
                                             "}")

