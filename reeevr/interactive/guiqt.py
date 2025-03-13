import sys

from mainloop import MainLoop

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt6.QtCore import Qt
from PyQt6 import QtGui
from QtDesignerGUIFile import Ui_ExcelToRConverterGUI


class GUI(QMainWindow):
    def __init__(self):
        super().__init__()


        # use the Ui_login_form
        self.ui = Ui_ExcelToRConverterGUI()
        self.ui.setupUi(self)

        #self.setWindowIcon(QtGui.QIcon('C:\\Users\mo14776\\Downloads\\piqrayLyT.png'))
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
        # show the login window
        self.show()

    def OpenPreloadFile(self):

        path = QFileDialog.getOpenFileName(parent=self,caption="Select preload file",filter="(*.txt)")
        self.ui.PreloadInputsFile.setText(f"{path[0]}")

        preload = []
        with open(f"{path[0]}","r") as f:

            for line in f:
                try:
                    preload.append(line.split(": ")[1].replace("\n",""))
                    print(preload[-1])
                except:
                    preload.append("")
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

    def OpenFolder(self):

        path = QFileDialog.getExistingDirectory(parent=self,caption="Select output folder")
        self.ui.OutputFolderFileName.setText(f"{path}")

    def Convert(self):
        self.ui.progressBar.setStyleSheet("QProgressBar"
                                          "{"
                                          "background-color : lightblue;"
                                          "border : 1px"
                                          "}")

        path = self.ui.FileName.text()
        testOutput = self.ui.OutputcellsLedit.text()
        ignoredsheets = self.ui.IgnoredSheetsLedit.text()
        costs = self.ui.CostsCellsLEdit.text()
        effectiveness = self.ui.EffectivenesCellsLEdit.text()
        treatments = self.ui.TreatmentNamesLEdit.text()
        willingnesstopay = self.ui.WillingnessToPayLEdit.text()

        convert = MainLoop(gui=True, progressbar=self.ui.progressBar, guitextbrowser=self.ui.textEdit)
        convert.set_vals(path,testOutput,ignoredsheets,costs,effectiveness,treatments,willingnesstopay)
        convert.run()

if __name__ == '__main__':
    app = QApplication(sys.argv)
    GUI = GUI()
    sys.exit(app.exec())
