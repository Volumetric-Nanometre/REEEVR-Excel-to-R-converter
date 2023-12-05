from parsers.reader import ExcelReader
from parsers.codegen import CodeGen
from parsers.outputs import ROutputs
from converters.variable import VariableConverter
import openpyxl
import os

try:
    os.remove("missing-func.log")
except:
    pass

#path = "../../tests/test workbooks/test_workbook_4.xlsm"
path = "C:/Users/mo14776/OneDrive - University of Bristol/Documents/Health Economics/REEVER/Examples/Hip replacement surgery Markov model V1.0 - modified.xlsm"
outputLang = "R"

workbook = openpyxl.load_workbook(path)


testOutput =[('Summary results', 'F6:G6')]  #[('Results', 'F13:H13'), ('Results', 'F12:H12')]   #
costs =[('Summary results', 'F6')]#,('Summary results', 'F7'),('Summary results', 'F8'),('Summary results', 'F9')] #[('Engine', 'E5'), ('Engine', 'F5')] #[('Results', 'F12'), ('Results', 'F13')]
effs = [('Summary results', 'G6')]#,('Summary results', 'G7'),('Summary results', 'G8'),('Summary results', 'G9')]#[('Engine', 'E6'), ('Engine', 'F6')]#[('Results', 'G12'), ('Results', 'G13')]
treatmentNames =['Cemented']#,'Uncemented','Hybrid', 'Reverse Hybrid'] #['Aspirin', 'Warfarin'] #['Drug A', 'Drug B']
willingnessToPay =[('Setup and run','D14')] #[('Model settings','N13')]

ignoredsheets = ['DSA', 'PSA', 'PSA results', 'DSA results', 'State trace - Cemented', 'State trace - Hybrid', 'State trace - Reverse hybrid']#,'State trace - Uncemented']

ignoredsheets = ['DSA', 'PSA', 'PSA results', 'DSA results']#, 'State trace - Cemented', 'State trace - Hybrid', 'State trace - Reverse hybrid','State trace - Uncemented']
print("[SUCCESS]")
print("Initialise variable converter ... ",end="")
varconverter = VariableConverter(workbook,ignoredsheets, outputLang)
print("[SUCCESS]")
print("Initialise outputs ... ",end="")
outputs = ROutputs(varconverter, "", "", testOutput, costs, effs,treatmentNames,willingnessToPay)

a = ExcelReader(varconverter, workbook, outputLang, ignoredsheets)

a.read()
b = CodeGen(a.unorderedcode, outputs, codefile="test_output.R")

b.order_code_snippets()
b.cull_code_snippets()
b.generate_code()
