from reeevr.parsers.reader import ExcelReader
from reeevr.parsers.codegen import CodeGen
from reeevr.parsers.outputs import ROutputs
from reeevr.converters.variable import VariableConverter
import openpyxl

path = "../../tests/test workbooks/test_workbook_4.xlsm"
outputLang = "R"

workbook = openpyxl.load_workbook(path)


testOutput = [('Engine', 'E5:G7')]

varconverter = VariableConverter(workbook,outputLang)

outputs = ROutputs(varconverter,"","",testOutput)

a = ExcelReader(varconverter,workbook, outputLang)

a.read()
b = CodeGen(a.unorderedcode, outputs, codefile="test_output.R")

b.order_code_snippets()
b.cull_code_snippets()
b.generate_code()