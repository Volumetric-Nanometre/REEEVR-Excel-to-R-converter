from parsers.reader import ExcelReader
from parsers.codegen import CodeGen
from parsers.outputs import ROutputs
from converters.variable import VariableConverter
import openpyxl
import os

import io

def print_to_string(*args, **kwargs):
    output = io.StringIO()
    print(*args, file=output, **kwargs)
    contents = output.getvalue()
    output.close()
    return contents

class MainLoop:

    def __init__(self, gui=False, progressbar = None, guitextbrowser = None):
        self.path = ""
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
    def set_vals(self, path, testoutput, ignoredsheets, costs, effectiveness, treatments, willingtopay):
        self.path = path

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


    def output(self,*args,**kwargs):
        if self.gui :
            self.text += print_to_string(*args, **kwargs)

            self.guitextbrowser.setPlainText(self.text)
            print(*args, **kwargs)
        else:

            print(*args, **kwargs)

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
            self.output("Open workbook ... ",end="")
            workbook = openpyxl.load_workbook(self.path)
            self.output("[SUCCESS]")
            self.update_progress()

            self.output("Initialise variable converter ... ",end="")
            varconverter = VariableConverter(workbook,self.ignoredsheets, "R")
            self.output("[SUCCESS]")
            self.update_progress()

            self.output("Initialise outputs ... ",end="")
            outputs = ROutputs(varconverter, "", "", self.testOutput, self.costs, self.effectiveness, self.treatments, self.willingnesstopay)
            self.output("[SUCCESS]")
            self.update_progress()

            self.output("Create reader ... ",end="")
            a = ExcelReader(varconverter, workbook, "R", self.ignoredsheets)
            self.output("[SUCCESS]")
            self.update_progress()

            self.output("Read code ... ",end="")
            a.read()
            self.output("[SUCCESS]")
            self.update_progress()

            self.output("Generate unordered code ... ",end="")
            b = CodeGen(varconverter, a.unorderedcode, outputs, codefile="test_output.R")
            b.second_pass()
            self.output("[SUCCESS]")
            self.update_progress()

            self.output("Order Code ... ",end="")
            b.order_code_snippets()
            self.output("[SUCCESS]")
            self.update_progress()

            self.output("Prune Code ... ",end="")
            b.cyclic_prune()
            self.output("[SUCCESS]")
            self.update_progress()
            self.output("Output Code ... ",end="")
            b.generate_code()
            self.output("[SUCCESS]")
            self.update_progress()
        except:
            self.output("[FAILED]")
            self.completionstate = 0
            self.progressbar.setStyleSheet("QProgressBar"
                                             "{"
                                             "background-color : red;"
                                             "border : 1px"
                                             "}")

