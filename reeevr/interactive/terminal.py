from reeevr.parsers.reader import ExcelReader
from reeevr.parsers.codegen import CodeGen

path = "../../tests/test workbooks/Two states Markov model_v0.1_18May2023.xlsm"

outputs = [('Results', 'G13:L13')]
a = ExcelReader(path,outputs, "R")

a.read()
b = CodeGen(a.unorderedcode, a.outputcells, codefile="test_output.R")

b.replace_averages()
b.order_code_snippets()
b.cull_code_snippets()
b.generate_code()