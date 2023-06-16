

class CodeGen:
    """
    Takes unordered code output then:
    1. orders code according to variable usage rules
    2. checks which variables are output
    3. strips unused variables from code
    4. outputs final code file
    """

    def __init__(self, unorderedcode, outputs,codefile):
        self.unorderedcode = unorderedcode
        self.orderedcode = {}
        self.culledcode = {}
        self.unusedcode = {}
        self.outputs = outputs
        self.dependantvars = outputs.output_cells()
        self.codefile = codefile

        self.mandatoryCode = "library(reeevr)\n" \
                             "library(BCEA)\n" \
                             "numberOfRuns = 1000\n"


    def order_code_snippets(self):
        """
        Order the code snippets such that
        :return:
        """

        remainingcode = self.unorderedcode
        count = 0

        while len(remainingcode) > 0 and count < 100:
            for key in list(remainingcode.keys()):

                addtocode = True

                for var in remainingcode[key][1]:
                    if var not in self.orderedcode.keys():
                        addtocode = False
                        break
                if addtocode:
                    self.orderedcode.update({key : self.unorderedcode[key]})
                    self.dependantvars = self.dependantvars + self.unorderedcode[key][1]
                    del remainingcode[key]

            count += 1

        if count == 100:

            errorstring = ", ".join([f'{item[0]} : {item[1][1]}' for item in remainingcode.items()])

            raise KeyError(f"Code cannot be ordered as the following variables are missing their corresponding dependancies: {errorstring}")

    def cull_code_snippets(self):
        """
        Remove code snippets that are not outputs
        or used to generate the output.

        i.e prune the code tree of non-used code
        {'variable' : ["codeified string", ['list','of','contained','vars'],cell.data_type]
        """

        for key in self.orderedcode.keys():

            if key in self.dependantvars:
                self.culledcode[key] = self.orderedcode[key]
            else:
                self.unusedcode[key] = self.orderedcode[key]

    def generate_code(self):
        """
        {'variable' : ["codeified string", ['list','of','contained','vars'],cell.data_type]}
        """
        with open(self.codefile,"w") as f:
            f.write(f"{self.mandatoryCode}")

            for item in self.culledcode.items():

                if item[1][2] != "f":
                    f.write(f"{item[0]} = {item[1][0]}\n")
                else:
                    item[1][0] = item[1][0].replace("%sep%", ",")
                    f.write(f"{item[1][0]}\n")

            f.write(f"{self.outputs.add_output_code()}")

if __name__ == "__main__":

    from reeevr.parsers.reader import ExcelReader

    path = "../../tests/test workbooks/test_workbook_3.xlsx"
    a=ExcelReader(path,"R")

    a.read()
    outputs = ['Frontend_E8', 'Frontend_E9']
    b = CodeGen(a.unorderedcode, outputs, codefile="test_output.R")

    b.order_code_snippets()
    b.cull_code_snippets()
    b.generate_code()