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
        robjects.globalenv.clear()

class AdaptiveFunctionTests(unittest.TestCase):
    """
       Unit tests covering the functions treated as simple transforms in formula.py
       These functions require an adjustment to the function name.
       e.g. LN and LOG work differently in Excel vs R.

       Test suite explicity loads an Excel file containing all functions and test cases.
       File is then converted using the standard conversion process (minus culling).
       File is then ran in R using rpy2, this generates an active R environment.
       We then interrogate the R environment and compare to the expected value stored in Excel.
       """
    @classmethod
    def setUpClass(self):
        """
        Test suite explicity loads an Excel file containing all functions and test cases.
        File is then converted using the standard conversion process (minus culling).
        File is then ran in R using rpy2, this generates an active R environment.
        We then interrogate the R environment and compare to the expected value stored in Excel.
        """
        self.excelpath = "test/excelunit/adaptive-tests.xlsx"
        self.rpath = "test/excelunit/adaptive-test-output.R"
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
    def test_LN(self):
        self.assertEqual(self.globalenv['Sheet1_B3'][0], 'LN')
        self.assertAlmostEqual(self.globalenv['Sheet1_I3'][0], self.globalenv['Sheet1_O3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J3'][0], self.globalenv['Sheet1_P3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K3'][0], self.globalenv['Sheet1_Q3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L3'][0], self.globalenv['Sheet1_R3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M3'][0], self.globalenv['Sheet1_S3'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N3'][0], self.globalenv['Sheet1_T3'][0])

    def test_BETAINV(self):
        self.assertEqual(self.globalenv['Sheet1_B4'][0], '_xlfn.BETA.INV')
        self.assertAlmostEqual(self.globalenv['Sheet1_I4'][0], self.globalenv['Sheet1_O4'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J4'][0], self.globalenv['Sheet1_P4'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K4'][0], self.globalenv['Sheet1_Q4'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L4'][0], self.globalenv['Sheet1_R4'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M4'][0], self.globalenv['Sheet1_S4'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N4'][0], self.globalenv['Sheet1_T4'][0])

        self.assertEqual(self.globalenv['Sheet1_B5'][0], 'BETAINV')
        self.assertAlmostEqual(self.globalenv['Sheet1_I5'][0], self.globalenv['Sheet1_O5'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J5'][0], self.globalenv['Sheet1_P5'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K5'][0], self.globalenv['Sheet1_Q5'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L5'][0], self.globalenv['Sheet1_R5'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M5'][0], self.globalenv['Sheet1_S5'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N5'][0], self.globalenv['Sheet1_T5'][0])

    def test_NORMINV(self):
        self.assertEqual(self.globalenv['Sheet1_B6'][0], '_xlfn.NORM.INV')
        self.assertAlmostEqual(self.globalenv['Sheet1_I6'][0], self.globalenv['Sheet1_O6'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J6'][0], self.globalenv['Sheet1_P6'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K6'][0], self.globalenv['Sheet1_Q6'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L6'][0], self.globalenv['Sheet1_R6'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M6'][0], self.globalenv['Sheet1_S6'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N6'][0], self.globalenv['Sheet1_T6'][0])

        self.assertEqual(self.globalenv['Sheet1_B7'][0], 'NORMINV')
        self.assertAlmostEqual(self.globalenv['Sheet1_I7'][0], self.globalenv['Sheet1_O7'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J7'][0], self.globalenv['Sheet1_P7'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K7'][0], self.globalenv['Sheet1_Q7'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L7'][0], self.globalenv['Sheet1_R7'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M7'][0], self.globalenv['Sheet1_S7'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N7'][0], self.globalenv['Sheet1_T7'][0])

    def test_LOGNORMINV(self):
        self.assertEqual(self.globalenv['Sheet1_B8'][0], '_xlfn.LOGNORM.INV')
        self.assertAlmostEqual(self.globalenv['Sheet1_I8'][0], self.globalenv['Sheet1_O8'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J8'][0], self.globalenv['Sheet1_P8'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K8'][0], self.globalenv['Sheet1_Q8'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L8'][0], self.globalenv['Sheet1_R8'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M8'][0], self.globalenv['Sheet1_S8'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N8'][0], self.globalenv['Sheet1_T8'][0])

        self.assertEqual(self.globalenv['Sheet1_B9'][0], 'LOGINV')
        self.assertAlmostEqual(self.globalenv['Sheet1_I9'][0], self.globalenv['Sheet1_O9'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J9'][0], self.globalenv['Sheet1_P9'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K9'][0], self.globalenv['Sheet1_Q9'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L9'][0], self.globalenv['Sheet1_R9'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M9'][0], self.globalenv['Sheet1_S9'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N9'][0], self.globalenv['Sheet1_T9'][0])

    def test_INT(self):
        self.assertEqual(self.globalenv['Sheet1_B10'][0], 'INT')
        self.assertAlmostEqual(self.globalenv['Sheet1_I10'][0], self.globalenv['Sheet1_O10'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J10'][0], self.globalenv['Sheet1_P10'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K10'][0], self.globalenv['Sheet1_Q10'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L10'][0], self.globalenv['Sheet1_R10'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M10'][0], self.globalenv['Sheet1_S10'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N10'][0], self.globalenv['Sheet1_T10'][0])


    def test_UPPER(self):
        self.assertEqual(self.globalenv['Sheet1_B11'][0], 'UPPER')
        self.assertAlmostEqual(self.globalenv['Sheet1_I11'][0], self.globalenv['Sheet1_O11'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J11'][0], self.globalenv['Sheet1_P11'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K11'][0], self.globalenv['Sheet1_Q11'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L11'][0], self.globalenv['Sheet1_R11'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M11'][0], self.globalenv['Sheet1_S11'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N11'][0], self.globalenv['Sheet1_T11'][0])

    def test_LOWER(self):
        self.assertEqual(self.globalenv['Sheet1_B12'][0], 'LOWER')
        self.assertAlmostEqual(self.globalenv['Sheet1_I12'][0], self.globalenv['Sheet1_O12'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J12'][0], self.globalenv['Sheet1_P12'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K12'][0], self.globalenv['Sheet1_Q12'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L12'][0], self.globalenv['Sheet1_R12'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M12'][0], self.globalenv['Sheet1_S12'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N12'][0], self.globalenv['Sheet1_T12'][0])

    def test_LEN(self):
        self.assertEqual(self.globalenv['Sheet1_B13'][0], 'LEN')
        self.assertAlmostEqual(self.globalenv['Sheet1_I13'][0], self.globalenv['Sheet1_O13'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J13'][0], self.globalenv['Sheet1_P13'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K13'][0], self.globalenv['Sheet1_Q13'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L13'][0], self.globalenv['Sheet1_R13'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M13'][0], self.globalenv['Sheet1_S13'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N13'][0], self.globalenv['Sheet1_T13'][0])
    def test_AVERAGE(self):
        self.assertEqual(self.globalenv['Sheet1_B16'][0], 'AVERAGE')
        self.assertAlmostEqual(self.globalenv['Sheet1_I16'][0], self.globalenv['Sheet1_O16'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J16'][0], self.globalenv['Sheet1_P16'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K16'][0], self.globalenv['Sheet1_Q16'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L16'][0], self.globalenv['Sheet1_R16'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M16'][0], self.globalenv['Sheet1_S16'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N16'][0], self.globalenv['Sheet1_T16'][0])

    def test_STDEVS(self):
        self.assertEqual(self.globalenv['Sheet1_B17'][0], '_xlfn.STDEV.S')
        self.assertAlmostEqual(self.globalenv['Sheet1_I17'][0], self.globalenv['Sheet1_O17'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J17'][0], self.globalenv['Sheet1_P17'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K17'][0], self.globalenv['Sheet1_Q17'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L17'][0], self.globalenv['Sheet1_R17'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M17'][0], self.globalenv['Sheet1_S17'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N17'][0], self.globalenv['Sheet1_T17'][0])

        self.assertEqual(self.globalenv['Sheet1_B18'][0], 'STDEV')
        self.assertAlmostEqual(self.globalenv['Sheet1_I18'][0], self.globalenv['Sheet1_O18'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_J18'][0], self.globalenv['Sheet1_P18'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_K18'][0], self.globalenv['Sheet1_Q18'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_L18'][0], self.globalenv['Sheet1_R18'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_M18'][0], self.globalenv['Sheet1_S18'][0])
        self.assertAlmostEqual(self.globalenv['Sheet1_N18'][0], self.globalenv['Sheet1_T18'][0])

    @classmethod
    def tearDownClass(self):
        robjects.globalenv.clear()

class REEEVRFunctionTests(unittest.TestCase):
    """
       Unit tests covering the functions treated as simple transforms in formula.py
       These functions require an adjustment to the function name.
       e.g. LN and LOG work differently in Excel vs R.

       Test suite explicity loads an Excel file containing all functions and test cases.
       File is then converted using the standard conversion process (minus culling).
       File is then ran in R using rpy2, this generates an active R environment.
       We then interrogate the R environment and compare to the expected value stored in Excel.
       """
    @classmethod
    def setUpClass(self):
        """
        Test suite explicity loads an Excel file containing all functions and test cases.
        File is then converted using the standard conversion process (minus culling).
        File is then ran in R using rpy2, this generates an active R environment.
        We then interrogate the R environment and compare to the expected value stored in Excel.
        """
        self.excelpath = "test/excelunit/reeevr-tests.xlsx"
        self.rpath = "test/excelunit/reeevr-test-output.R"
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
    def test_full_excel(self):

        for i in range(3,41):
            function = self.globalenv[f'Sheet1_B{i}'][0]
            self.assertAlmostEqual(self.globalenv[f'Sheet1_I{i}'][0], self.globalenv[f'Sheet1_O{i}'][0],msg=f"{function} - Sheet1_I{i}\n")
            self.assertAlmostEqual(self.globalenv[f'Sheet1_J{i}'][0], self.globalenv[f'Sheet1_P{i}'][0])
            self.assertAlmostEqual(self.globalenv[f'Sheet1_K{i}'][0], self.globalenv[f'Sheet1_Q{i}'][0])
            self.assertAlmostEqual(self.globalenv[f'Sheet1_L{i}'][0], self.globalenv[f'Sheet1_R{i}'][0])
            self.assertAlmostEqual(self.globalenv[f'Sheet1_M{i}'][0], self.globalenv[f'Sheet1_S{i}'][0])
            self.assertAlmostEqual(self.globalenv[f'Sheet1_N{i}'][0], self.globalenv[f'Sheet1_T{i}'][0])

    @classmethod
    def tearDownClass(self):
        robjects.globalenv.clear()



if __name__ == '__main__':

    unittest.main()
