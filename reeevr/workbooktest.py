from reader import ExcelReader
from codegen import CodeGen
from outputs import ROutputs
from variable import VariableConverter
import openpyxl
import copy
import unittest
import rpy2.robjects as robjects
import hashlib

def file_hash(path):
    hasher = hashlib.sha256()
    with open(path, 'rb') as file:
        for chunk in iter(lambda: file.read(4096), b""):
            hasher.update(chunk)
    return hasher.hexdigest()


class TestWorkbook1(unittest.TestCase):
    """

    """
    @classmethod
    def setUpClass(self):
        """
        Test suite explicity loads an Excel file containing all functions and test cases.
        File is then converted using the standard conversion process (minus culling).
        File is then ran in R using rpy2, this generates an active R environment.
        We then interrogate the R environment and compare to the expected value stored in Excel.
        """
        self.excelpath = "test/excel workbook/test_workbook_1.xlsx"
        self.rpath = "test/excel workbook/test_workbook_1_output.R"
        workbook = openpyxl.load_workbook(self.excelpath)
        outputLang = 'R'
        ignoredsheets = ["PSA"]
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

    def test_deterministic_conversion(self):
        """
        Given no changes to the Excel file, returned files should always be the same.
        We can test this using a cryptographic hash. If the contents of the files are different then the conversion
        has not generated the same file.
        """

        hashOriginal = file_hash("test/excel workbook/test_workbook_1.R")
        hashGenerated = file_hash(self.rpath)
        self.assertEqual(hashOriginal,hashGenerated, "Hash comparison failed - converter update has changed deterministic output\n")

    def test_programatic_regression(self):
        """
        While the converted file may be correct, the individual functions called may change definition when called in R
        or via incompatible changes from Excel. This test catches when these regressions may occur
        """

        self.assertAlmostEqual(self.globalenv[f'Sheet1_A1'][0], 3,
                               msg=f"Sheet1_A1 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Sheet1_A2'][0], 3,
                               msg=f"Sheet1_A2 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Sheet1_A3'][0], 5,
                               msg=f"Sheet1_A3 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Sheet1_A4'][0], 8,
                               msg=f"Sheet1_A4 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Sheet1_D2'][0], 5,
                               msg=f"Sheet1_D2 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Sheet1_D3'][0], 6,
                               msg=f"Sheet1_D3 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Sheet1_D4'][0], 7,
                               msg=f"Sheet1_D4 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Sheet1_D5'][0], 8,
                               msg=f"Sheet1_D5 ... [FAIL]\n")
        self.assertEqual(self.globalenv[f'Sheet1_C1'][0], "sdf",
                               msg=f"Sheet1_D5 ... [FAIL]\n")

    @classmethod
    def tearDownClass(self):
        robjects.globalenv.clear()


class TestWorkbook2(unittest.TestCase):
    """

    """
    @classmethod
    def setUpClass(self):
        """
        Test suite explicity loads an Excel file containing all functions and test cases.
        File is then converted using the standard conversion process (minus culling).
        File is then ran in R using rpy2, this generates an active R environment.
        We then interrogate the R environment and compare to the expected value stored in Excel.
        """
        self.excelpath = "test/excel workbook/test_workbook_2.xlsx"
        self.rpath = "test/excel workbook/test_workbook_2_output.R"
        workbook = openpyxl.load_workbook(self.excelpath)
        outputLang = 'R'
        ignoredsheets = ["PSA"]
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

    def test_deterministic_conversion(self):
        """
        Given no changes to the Excel file, returned files should always be the same.
        We can test this using a cryptographic hash. If the contents of the files are different then the conversion
        has not generated the same file.
        """

        hashOriginal = file_hash("test/excel workbook/test_workbook_2.R")
        hashGenerated = file_hash(self.rpath)
        self.assertEqual(hashOriginal,hashGenerated, "Hash comparison failed - converter update has changed deterministic output\n")

    def test_programatic_regression(self):
        """
        While the converted file may be correct, the individual functions called may change definition when called in R
        or via incompatible changes from Excel. This test catches when these regressions may occur
        """

        self.assertAlmostEqual(self.globalenv[f'willingness_to_pay'][0], 20000,
                               msg=f"willingness_to_pay (Frontend_E5) ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Frontend_E7'][0], 4600,
                               msg=f"Frontend_E5 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Frontend_E8'][0], 1540,
                               msg=f"Frontend_E8 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_E5'][0], 100,
                               msg=f"Engine_E5 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_E6'][0], 19.2,
                               msg=f"Engine_E6 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_E7'][0], 383900,
                               msg=f"Engine_E7 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_F5'][0], 560,
                               msg=f"Engine_F5 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_F6'][0], 19.3,
                               msg=f"Engine_F6 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_F7'][0], 385440,
                               msg=f"Engine_F7 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_G5'][0], 460,
                               msg=f"Engine_G5 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_G6'][0], 0.1,
                               msg=f"Engine_G6 ... [FAIL]\n")
        self.assertAlmostEqual(self.globalenv[f'Engine_G7'][0], 1540,
                               msg=f"Engine_G7 ... [FAIL]\n")

    @classmethod
    def tearDownClass(self):
        robjects.globalenv.clear()

if __name__ == '__main__':
    unittest.main()
