from reader import ExcelReader
from codegen import CodeGen
from outputs import ROutputs
from variable import VariableConverter
import openpyxl
import os

try:
    os.remove("missing-func.log")
except:
    pass

path = "../Examples/HIPS model/HIPS Markov model demo.xlsm"

print("Open workbook ... ",end="")
workbook = openpyxl.load_workbook(path)
print("[SUCCESS]")
print("Assign user defined variables ... ",end="")
testOutput = []
costs = [('Summary results', 'F6:F9')]
effs = [('Summary results', 'G6:G9')]
treatmentNames =['Cemented','Uncemented','Hybrid', 'Reverse Hybrid']
willingnessToPay =[('Setup and run','D14')]

ignoredsheets = ['DSA', 'PSA', 'PSA results', 'DSA results']
print("[SUCCESS]")
print("Initialise variable converter ... ",end="")
varconverter = VariableConverter(workbook,ignoredsheets)
print("[SUCCESS]")
print("Initialise outputs ... ",end="")
outputs = ROutputs(varconverter, "", testOutput, costs, effs,treatmentNames,willingnessToPay, BCEA= False)
print("[SUCCESS]")
print("Create reader ... ",end="")
a = ExcelReader(varconverter, workbook, ignoredsheets)
print("[SUCCESS]")
print("Read code ... ",end="")
a.read()
print("[SUCCESS]")
print("Generate unordered code ... ",end="")
b = CodeGen(varconverter, a.unorderedcode, outputs, folder="", filename="test_output.R")
print("[SUCCESS]")
print("Perform second pass ... ",end="")
b.second_pass()
print("[SUCCESS]")
print("Order Code ... ",end="")
b.order_code_snippets()
print("[SUCCESS]")
print("Prune Code ... ",end="")
b.cyclic_prune()
print("[SUCCESS]")
print("Output Code ... ",end="")
b.generate_code()
print("[SUCCESS]")
