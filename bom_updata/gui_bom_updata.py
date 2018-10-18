#!/usr/bin/env python

"""PyQt4 port of the dialogs/standarddialogs example from Qt v4.x"""

# This is only needed for Python v2 but is harmless for Python v3.
#import sip
#sip.setapi('QString', 2)

import sys
import bom_updata
from PySide import QtCore, QtGui


class Dialog(QtGui.QDialog):
    MESSAGE = "<p>Message boxes have a caption, a text, and up to three " \
            "buttons, each with standard or custom texts.</p>" \
            "<p>Click a button to close the message box. Pressing the Esc " \
            "button will activate the detected escape button (if any).</p>"

    def __init__(self, parent=None):
        super(Dialog, self).__init__(parent)

        self.openFilesPath = ''

        self.errorMessageDialog = QtGui.QErrorMessage(self)

        frameStyle = QtGui.QFrame.Sunken | QtGui.QFrame.Panel



        self.openAccessFileLable = QtGui.QLabel()
        self.openAccessFileLable.setFrameStyle(frameStyle)
        self.openAccessFileButton = QtGui.QPushButton("get access file")


        self.openExcelFileLable = QtGui.QLabel()
        self.openExcelFileLable.setFrameStyle(frameStyle)
        self.openExcelFileButton = QtGui.QPushButton("get Eacel file")

        self.logLabel = QtGui.QLabel('log:')
        self.logEdit = QtGui.QTextEdit()

        self.updata_ExcelButton = QtGui.QPushButton('updata Bom')
        self.updata_AccessButton = QtGui.QPushButton('updata DB_access')





        self.openAccessFileButton.clicked.connect(self.setOpenAccessFileName)
        self.openExcelFileButton.clicked.connect(self.setOpenExcelFileName)
        self.updata_ExcelButton.clicked.connect(self.updataBom)
        self.updata_AccessButton.clicked.connect(self.updataAccess)



        layout = QtGui.QGridLayout()
        layout.setColumnStretch(1, 1)
        layout.setColumnMinimumWidth(1, 250)
        
        layout.addWidget(self.openAccessFileButton, 0, 0)
        layout.addWidget(self.openAccessFileLable, 0, 1)
        layout.addWidget(self.openExcelFileButton,1,0)
        layout.addWidget(self.openExcelFileLable,1,1)
        layout.addWidget(self.logLabel,2,0)
        layout.addWidget(self.logEdit,2,1,5,1)
        layout.addWidget(self.updata_ExcelButton,3,0)
        layout.addWidget(self.updata_AccessButton,4,0)
    
        self.setLayout(layout)

        self.setWindowTitle("Standard Dialogs")





    def setOpenAccessFileName(self):    
        options = QtGui.QFileDialog.Options()
        # if not self.native.isChecked():
        #     options |= QtGui.QFileDialog.DontUseNativeDialog
        fileName, filtr = QtGui.QFileDialog.getOpenFileName(self,
                "get access file()",
                self.openAccessFileLable.text(),
                "All Files (*mdb)", "", options)
        if fileName:
            self.openAccessFileLable.setText(fileName)



    def setOpenExcelFileName(self):    
        options = QtGui.QFileDialog.Options()
        # if not self.native.isChecked():
        #     options |= QtGui.QFileDialog.DontUseNativeDialog
        fileName, filtr = QtGui.QFileDialog.getOpenFileName(self,
                "get excel file()",           
                self.openExcelFileLable.text(),
                "All Files (*.xls)", "", options)
        if fileName:
            self.openExcelFileLable.setText(fileName)


    def errorMessage(self):    
        self.errorMessageDialog.showMessage("This dialog shows and remembers "
                "error messages. If the checkbox is checked (as it is by "
                "default), the shown message will be shown again, but if the "
                "user unchecks the box the message will not appear again if "
                "QErrorMessage.showMessage() is called with the same message.")
        self.errorLabel.setText("If the box is unchecked, the message won't "
                "appear again.")

    def updataBom(self):
        access_file = self.openAccessFileLable.text()
        excel_file = self.openExcelFileLable.text()

        # print (self.openExcelFileLable.text())
        bom_updata.bom_updata(access_file,excel_file)
    def updataAccess(self):
        access_file = self.openAccessFileLable.text()
        excel_file = self.openExcelFileLable.text()
        bom_updata.DB_updata(access_file,excel_file)

if __name__ == '__main__':
    app = QtGui.QApplication(sys.argv)
    dialog = Dialog()
    sys.exit(dialog.exec_())
