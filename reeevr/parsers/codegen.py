

class CodeGen:
    """
    Takes unordered code output then:
    1. orders code according to variable usage rules
    2. checks which variables are output
    3. strips unused variables from code
    4. outputs final code file
    """

    def __init__(self, unorderedcode, outputs):
        self.unorderedcode = unorderedcode
        self.orderedcode = {}
        self.unusedcode = {}
        self.dependantvars = outputs


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

if __name__ == "__main__":

    from reeevr.parsers.reader import ExcelReader

    path = "C:/Users/mieha/Documents/REEVER/Test workbooks/test_workbook_3.xlsx"
    a=ExcelReader(path,"Python")

    a.read()
    b = CodeGen(a.unorderedcode, [])

    b.order_code_snippets()
    [print(item) for item in b.dependantvars]