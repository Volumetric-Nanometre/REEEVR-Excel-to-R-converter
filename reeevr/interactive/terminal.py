from reeevr.parsers.reader import ExcelReader
from reeevr.parsers.codegen import CodeGen

path = "../../tests/test workbooks/test_workbook_4.xlsm"
a = ExcelReader(path, "R")

a.read()
outputs = ['Frontend_E9', 'Frontend_E10']
b = CodeGen(a.unorderedcode, outputs)

b.replace_averages()
b.order_code_snippets()
b.cull_code_snippets()
b.generate_code()