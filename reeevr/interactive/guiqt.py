import sys

from mainloop import MainLoop

from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog
from PyQt6.QtCore import Qt
from QtDesignerGUIFile import Ui_ExcelToRConverterGUI


class GUI(QMainWindow):
    def __init__(self):
        super().__init__()

        # use the Ui_login_form
        self.ui = Ui_ExcelToRConverterGUI()
        self.ui.setupUi(self)

        self.ui.OpenFileButton.clicked.connect(self.OpenFile)
        self.ui.OpenFolderButton.clicked.connect(self.OpenFolder)
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
