

class ROutputs:

    def __init__(self, converter, folder, PSAoutputs, costs, effs, treatments, willingnesstopay, BCEA = False):

        self.varconverter = converter
        self.outputFolder = folder
        self.PSAOutputs = self.expand_outputs(PSAoutputs)
        self.costs = self.expand_outputs(costs)
        self.effectiveness = self.expand_outputs(effs)
        self.treatments = treatments
        self.willingnesstopay = self.expand_outputs(willingnesstopay)
        self.BCEA = BCEA

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

        outputCells = list(set(self.PSAOutputs + self.costs + self.effectiveness + self.willingnesstopay))
        return outputCells
    def add_output_code(self):

        fullOutputCode = ""

        fullOutputCode += self.file_output()
        if(self.BCEA):
            fullOutputCode += self.bcea()


        return fullOutputCode

    def file_output(self):

        outputList = self.costs + self.effectiveness + self.PSAOutputs
        output = ",".join([f'{val}' for val in outputList])
        if self.outputFolder == "":
            fileOutputCode = f"psafile <- 'psa.txt'\n"
        else:
            fileOutputCode = f"psafile <- '{self.outputFolder}/psa.txt'\n"

        fileOutputCode += f"PSA_output <- PSA_output(psafile, dec = '.',{output})\n"
        return fileOutputCode

    def bcea(self):
        """
        Use the R BCEA_all_output_loop function to generate BCEA outputs
        """
        bcea_code = ""

        if self.costs and self.effectiveness and self.treatments:


            costs = ",".join(self.costs)
            effs = ",".join(self.effectiveness)
            treatments = ",".join([f'"{val}"' for val in self.treatments])
            treatments = f'c({treatments})\n'

            bcea_code = ""
            bcea_code += f"costs = BCEA_matrix({costs})\n"
            bcea_code += f"effs = BCEA_matrix({effs})\n"
            bcea_code += f"treatments = {treatments}\n"
            bcea_code += f"BCEA_all_output_loop(costs,effs,treatments,{self.willingnesstopay[0]})\n"

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