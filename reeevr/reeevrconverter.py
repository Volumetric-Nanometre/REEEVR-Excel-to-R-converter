import os
import sys
from mainloop import MainLoop

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt6.QtCore import Qt
from QtDesignerGUIFile import Ui_ExcelToRConverterGUI
import traceback
import logging
import datetime

logger = logging.getLogger(__name__)


class GUI(QMainWindow):
    def __init__(self):
        super().__init__()


        # use the Ui_login_form
        self.ui = Ui_ExcelToRConverterGUI()
        self.ui.setupUi(self)

        self.setWindowTitle("REEEVR - Excel to R Converter")

        self.ui.OpenFileButton.clicked.connect(self.OpenFile)
        self.ui.OpenFolderButton.clicked.connect(self.OpenFolder)
        self.ui.PreloadFileOpen.clicked.connect(self.OpenPreloadFile)
        self.ui.ConvertButton.clicked.connect(self.Convert)
        self.ui.progressBar.setProperty("value", 0)
        self.ui.progressBar.setAlignment(Qt.AlignmentFlag["AlignHCenter"])
        self.ui.progressBar.setStyleSheet("QProgressBar"
                                          "{"
                                          "background-color : lightblue;"
                                          "border : 1px"
                                          "}")
        self.show()

        logger.debug("Initialisation complete")

    def OpenPreloadFile(self):

        logger.info("Loading preload file")

        path = QFileDialog.getOpenFileName(parent=self,caption="Select preload file",filter="(*.txt)")
        self.ui.PreloadInputsFile.setText(f"{path[0]}")

        preload = []

        logger.debug("Opening preload file")
        with open(f"{path[0]}","r") as f:
            logger.debug("File opened")
            logger.debug("Preload data:")
            for index, line in enumerate(f):
                try:
                    preload.append(line.split(": ")[1].replace("\n",""))
                except:
                    preload.append("")
                logger.debug(f"Line {index} - {preload[-1]}")

            self.ui.FileName.setText(f"{preload[0]}")
            self.ui.OutputFolderFileName.setText(f"{preload[1]}")
            self.ui.OutputcellsLedit.setText(f"{preload[2]}")
            self.ui.IgnoredSheetsLedit.setText(f"{preload[3]}")
            self.ui.CostsCellsLEdit.setText(f"{preload[4]}")
            self.ui.EffectivenesCellsLEdit.setText(f"{preload[5]}")
            self.ui.TreatmentNamesLEdit.setText(f"{preload[6]}")
            self.ui.WillingnessToPayLEdit.setText(f"{preload[7]}")


    def OpenFile(self):

        path = QFileDialog.getOpenFileName(parent=self,caption="Select excel file",filter="Workbooks (*.xlsx *xlsm)")
        self.ui.FileName.setText(f"{path[0]}")
        logger.debug(f"Path to workbook: {path[0]}")

    def OpenFolder(self):

        path = QFileDialog.getExistingDirectory(parent=self,caption="Select output folder")
        self.ui.OutputFolderFileName.setText(f"{path}")
        logger.debug(f"Path to output: {path}")

    def Convert(self):

        logger.info("Starting conversion")

        self.ui.progressBar.setStyleSheet("QProgressBar"
                                          "{"
                                          "background-color : lightblue;"
                                          "border : 1px"
                                          "}")

        path = self.ui.FileName.text()
        folder = self.ui.OutputFolderFileName.text()
        testOutput = self.ui.OutputcellsLedit.text()
        ignoredsheets = self.ui.IgnoredSheetsLedit.text()
        costs = self.ui.CostsCellsLEdit.text()
        effectiveness = self.ui.EffectivenesCellsLEdit.text()
        treatments = self.ui.TreatmentNamesLEdit.text()
        willingnesstopay = self.ui.WillingnessToPayLEdit.text()
        BCEA = self.ui.BCEACheckBox.isChecked()

        logger.info("Final input parameters:\n"
                     f"Path: {path}\n"
                     f"Folder: {folder}\n"
                     f"Additional outputs: {testOutput}\n"
                     f"Ignored sheets: {ignoredsheets}\n"
                     f"Cost cells: {costs}\n"
                     f"Effectiveness cells: {effectiveness}\n"
                     f"Treatment names: {treatments}\n"
                     f"Willingness to pay cell: {willingnesstopay}\n"
                     f"BCEA flag: {BCEA}\n")

        try:
            convert = MainLoop(gui=True, progressbar=self.ui.progressBar, guitextbrowser=self.ui.textEdit)
            convert.set_vals(path,folder, testOutput, ignoredsheets, costs, effectiveness, treatments, willingnesstopay,BCEA)
            convert.run()
        except:
            logger.critical("Converter failure traceback:\n"
                             f"{sys.exc_info()}\n"
                             f"{traceback.print_tb()}\n")
if __name__ == '__main__':

    try:
        os.mkdir("Log Files")
    except:
        pass

    logging.basicConfig(filename=f"Log Files/{datetime.datetime.now()}-reeevr-logging.log",level=logging.DEBUG)

    logger.info("Starting R-Converter")
    try:
        import pyi_splash
        app = QApplication(sys.argv)
        pyi_splash.close()
    except:
        logger.info("Splash not available")
        app = QApplication(sys.argv)
    logger.info("Loading GUI")
    GUI = GUI()
    sys.exit(app.exec())
