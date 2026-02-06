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

#path = "../../tests/test workbooks/test_workbook_4.xlsm"
path = "C:/Users/mo14776/OneDrive - University of Bristol/Documents/Health Economics/REEVER/Examples/Hip replacement surgery Markov model V1.0 - modified.xlsm"
#path = "C:/Users/mo14776/OneDrive - University of Bristol/Documents/Health Economics/REEVER/Examples/Hip replacement surgery Markov model V1.0.xlsm"
#path = "C:/Users/mo14776/OneDrive - University of Bristol/Documents/Health Economics/REEVER/Examples/HIPS Markov model demo.xlsm"
#path = "C:/Users/mo14776/PycharmProjects/Excel-R-compiler/v1.0.0-alpha/test-funcs.xlsx"

#path = "C:\\Users\\mo14776\\Downloads\\DAP65 Cost-effectiveness model v6.0.xlsm"
#path = "C:\\Users\\mo14776\\PycharmProjects\\Excel-R-compiler\\tests\\test workbooks\\Two states Markov model_v0.2_03Jul2023_PSA.xlsm"
#path = "C:\\Users\\mo14776\\PycharmProjects\\Excel-R-compiler\\tests\\confidential workbooks\\Adolescent Obesity_02Sep2024_v1.0_Redacted.xlsm"


outputLang = "R"
print("Open workbook ... ",end="")
workbook = openpyxl.load_workbook(path)
print("[SUCCESS]")
print("Assign user defined variables ... ",end="")
testOutput = []
costs = [('Summary results', 'F6:F9')]#[('Results','E12:E14')] #[('Summary results', 'F6'),('Summary results', 'F7'),('Summary results', 'F8'),('Summary results', 'F9')]#[('Summary results', 'D6'),('Summary results', 'D7'),('Summary results', 'D8'),('Summary results', 'D9')] #[('Engine', 'E5'), ('Engine', 'F5')] #[('Results', 'F12'), ('Results', 'F13')]
effs = [('Summary results', 'G6:G9')]#[('Results','G12:G14')] #[('Summary results', 'G6'),('Summary results', 'G7'),('Summary results', 'G8'),('Summary results', 'G9')]#[('Summary results', 'E6'),('Summary results', 'E7'),('Summary results', 'E8'),('Summary results', 'E9')]#[('Engine', 'E6'), ('Engine', 'F6')]#[('Results', 'G12'), ('Results', 'G13')]
treatmentNames =['Cemented','Uncemented','Hybrid', 'Reverse Hybrid'] #['Drug A', 'Drug B','No treatment'] #['Genedrive','Genomadix cube','Laboratory genetic test', 'No test']# #['Aspirin', 'Warfarin'] #['Drug A', 'Drug B']
willingnessToPay =[('Setup and run','D14')]#[('Results','E8')] #[('Setup and run','D20')]# #[('Model settings','N13')]

ignoredsheets = ['DSA', 'PSA', 'PSA results', 'DSA results']#, 'State trace - Cemented', 'State trace - Hybrid', 'State trace - Reverse hybrid','State trace - Uncemented']
print("[SUCCESS]")
print("Initialise variable converter ... ",end="")
varconverter = VariableConverter(workbook,ignoredsheets, outputLang)
print("[SUCCESS]")
print("Initialise outputs ... ",end="")
outputs = ROutputs(varconverter, "", testOutput, costs, effs,treatmentNames,willingnessToPay, BCEA= False)
print("[SUCCESS]")
print("Create reader ... ",end="")
a = ExcelReader(varconverter, workbook, outputLang, ignoredsheets)
print("[SUCCESS]")
print("Read code ... ",end="")
a.read()
print("[SUCCESS]")
print("Generate unordered code ... ",end="")
b = CodeGen(varconverter, a.unorderedcode, outputs, codefile="test_output.R")
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
