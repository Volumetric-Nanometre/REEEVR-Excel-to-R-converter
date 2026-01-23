from reader import ExcelReader
from codegen import CodeGen
from outputs import ROutputs
from variable import VariableConverter
import openpyxl
import copy
import unittest
import rpy2.robjects as robjects


class SimpleFunctionTests(unittest.TestCase):
    """
    Unit tests covering the functions treated as simple transforms in formula.py
    These functions should all work exactly the same in R as in Excel (including order of arguments).
    """
    @classmethod
    def setUpClass(self):
        """
        Test suite explicity loads an Excel file containing all functions and test cases.
        File is then converted using the standard conversion process (minus culling).
        File is then ran in R using rpy2, this generates an active R environment.
        We then interrogate the R environment and compare to the expected value stored in Excel.
        """
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
    #
    # Single value function tests
    #
    def test_SQRT(self):
        self.assertEqual(self.globalenv['Sheet1_B3'][0], 'SQRT')
        self.assertAlmostEqual(self.globalenv['Sheet1_D3'][0], self.globalenv['Sheet1_E3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_G3'][0], self.globalenv['Sheet1_H3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J3'][0], self.globalenv['Sheet1_K3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M3'][0], self.globalenv['Sheet1_N3'][0])

    def test_EXP(self):
        self.assertEqual(self.globalenv['Sheet1_B4'][0], 'EXP')
        self.assertAlmostEqual(self.globalenv['Sheet1_D4'][0], self.globalenv['Sheet1_E4'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_G4'][0], self.globalenv['Sheet1_H4'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J4'][0], self.globalenv['Sheet1_K4'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M4'][0], self.globalenv['Sheet1_N4'][0])

    def test_ABS(self):
        self.assertEqual(self.globalenv['Sheet1_B5'][0], 'ABS')
        self.assertAlmostEqual(self.globalenv['Sheet1_D5'][0], self.globalenv['Sheet1_E5'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_G5'][0], self.globalenv['Sheet1_H5'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J5'][0], self.globalenv['Sheet1_K5'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M5'][0], self.globalenv['Sheet1_N5'][0])

    def test_ROUND(self):
        self.assertEqual(self.globalenv['Sheet1_B6'][0], 'ROUND')
        self.assertAlmostEqual(self.globalenv['Sheet1_D6'][0], self.globalenv['Sheet1_E6'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_G6'][0], self.globalenv['Sheet1_H6'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J6'][0], self.globalenv['Sheet1_K6'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M6'][0], self.globalenv['Sheet1_N6'][0])
    #
    # Array value function tests
    #
    def test_MAX(self):
        self.assertEqual(self.globalenv['Sheet1_B6'][0], 'ROUND')
        self.assertAlmostEqual(self.globalenv['Sheet1_M9'][0], self.globalenv['Sheet1_N9'][0])

    def test_MIN(self):
        self.assertEqual(self.globalenv['Sheet1_B10'][0], 'MIN')
        self.assertAlmostEqual(self.globalenv['Sheet1_M10'][0], self.globalenv['Sheet1_N10'][0])

    @classmethod
    def tearDownClass(self):
        print("tearDownClass")

if __name__ == '__main__':

    unittest.main()
