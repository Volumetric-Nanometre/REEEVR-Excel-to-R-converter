

class ROutputs:

    def __init__(self,converter,print,file,df,costs,effs,treatments):

        self.varconverter = converter
        self.printOutputs = self.expand_outputs(print)
        self.fileOutputs = self.expand_outputs(file)
        self.dataframeOutputs = self.expand_outputs(df)
        self.costs = self.expand_outputs(costs)
        self.effectiveness = self.expand_outputs(effs)
        self.treatments = treatments

    def expand_outputs(self,outputs):

        expandedOutputs = []
        for output in outputs:
            sheet = output[0]
            range = output[1]

            if ":" in range:
                code, vars = self.varconverter.excel_range_to_r_vector(sheet,range)
                expandedOutputs += vars
            else:
                varname = self.varconverter.excel_cell_to_variable(sheet, range)
                expandedOutputs.append(varname)

        return expandedOutputs

    def output_cells(self):

        outputCells = list(set(self.dataframeOutputs + self.fileOutputs + self.dataframeOutputs))
        return outputCells
    def add_output_code(self):

        fullOutputCode = ""

        fullOutputCode += self.create_dataframe()
        fullOutputCode += self.file_output()
        fullOutputCode += self.print_output()
        fullOutputCode += self.bcea()


        return fullOutputCode
    def print_output(self):

        printOutputCode = ""

        for val in self.printOutputs:

            printOutputCode += f"print({val})\n"

        return printOutputCode

    def file_output(self):

        fileOutputCode = ""
        for val in self.fileOutputs:

            fileOutputCode += f'cat({val}, file = "{val}-output-file.txt")\n'

        return fileOutputCode

    def create_dataframe(self):
        """
        Use the R BCEA_dataframe function to generate a dataframe
        """
        dataframe_code = ",".join(self.dataframeOutputs)

        dataframe_code = f"BCEAdf = BCEA_dataframe({dataframe_code})\n"
        dataframe_code += "print(BCEAdf)\n"


        return dataframe_code

    def bcea(self):
        """
        Use the R BCEA_dataframe function to generate a dataframe
        """
        bcea_code = ""

        if self.costs and self.effectiveness and self.treatments:


            costs = ",".join(self.costs)
            effs = ",".join(self.effectiveness)
            treatments = ",".join([f'"{val}"' for val in self.treatments])
            print(treatments)
            treatments = f'c({treatments})\n'

            bcea_code = ""
            bcea_code += f"costs = BCEA_matrix({costs})\n"
            bcea_code += f"effs = BCEA_matrix({effs})\n"
            bcea_code += f"treatments = {treatments}\n"
            bcea_code += f"BCEA_all_output_loop(costs,effs,treatments,100)\n"

        return bcea_code


if __name__ == "__main__":
    printOutputs = ["test1", "test2"]
    fileOutputs = ["test3","test4"]
    dataframeOutputs = ["test5", "test6"]

    outputs = ROutputs()
    outputs.printOutputs = printOutputs
    outputs.fileOutputs = fileOutputs
    outputs.dataframeOutputs = dataframeOutputs

    print(outputs.add_output_code())