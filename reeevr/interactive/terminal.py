from reeevr.parsers.reader import ExcelReader
from reeevr.parsers.codegen import CodeGen
from reeevr.parsers.outputs import ROutputs
from reeevr.converters.variable import VariableConverter
import openpyxl

path = "../../tests/test workbooks/Two states Markov model_v0.2_03Jul2023_PSA.xlsm"
outputLang = "R"

workbook = openpyxl.load_workbook(path)


testOutput = [('Results', 'F13:H13'), ('Results', 'F12:H12')]  # [('Summary results', 'F6:G9')] #
costs = [('Results', 'F12'), ('Results', 'F13')]
effs = [('Results', 'G12'), ('Results', 'G13')]
varconverter = VariableConverter(workbook, outputLang)

outputs = ROutputs(varconverter, "", "", testOutput, costs, effs)

a = ExcelReader(varconverter, workbook, outputLang)

a.read()
b = CodeGen(a.unorderedcode, outputs, codefile="test_output.R")

b.order_code_snippets()
b.cull_code_snippets()
b.generate_code()
