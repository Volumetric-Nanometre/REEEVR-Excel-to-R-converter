from reader import ExcelReader
from codegen import CodeGen
from outputs import ROutputs
from variable import VariableConverter
import openpyxl
import copy
import unittest
import rpy2.robjects as robjects


class FunctionTests(unittest.TestCase):

    @classmethod
    def setUpClass(self):
        self.excelpath = "test/excelunit/simple-tests.xlsx"
        self.rpath = "test/excelunit/simple-test-output.R"
        workbook = openpyxl.load_workbook(self.excelpath)
        outputLang = 'R'
        ignoredsheets = []
        varconverter = VariableConverter(workbook, ignoredsheets, outputLang)
        outputs = ROutputs(varconverter, "", "", [], [], [], [], [])
        reader = ExcelReader(varconverter, workbook, outputLang, ignoredsheets)
        reader.read()
        codegen = CodeGen(varconverter, reader.unorderedcode,outputs , codefile=self.rpath)
        codegen.second_pass()
        codegen.order_code_snippets()
        codegen.culledcode = copy.deepcopy(codegen.orderedcode)
        codegen.none_strip()
        codegen.generate_code(writeoutputs=False)

        self.r_source = robjects.r['source']
        self.r_source(self.rpath)
        self.globalenv = robjects.globalenv

    def test_SQRT(self):
        self.assertEqual(self.globalenv['Sheet1_B3'][0], 'SQRT')
        self.assertAlmostEqual(self.globalenv['Sheet1_D3'][0], self.globalenv['Sheet1_E3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_G3'][0], self.globalenv['Sheet1_H3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J3'][0], self.globalenv['Sheet1_K3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M3'][0], self.globalenv['Sheet1_N3'][0])



    @classmethod
    def tearDownClass(self):
        print("tearDownClass")

if __name__ == '__main__':

    unittest.main()
