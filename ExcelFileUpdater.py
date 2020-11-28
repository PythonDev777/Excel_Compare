# -*- coding: utf-8 -*-

# Form implementation generated from reading ui file 'ExcelFileUpdater.ui'
#
# Created by: PyQt5 UI code generator 5.9.2
#
# WARNING! All changes made in this file will be lost!

from PyQt5 import QtCore, QtGui, QtWidgets
import pandas as pd

class Ui_ExFileWindow(object):
    def __init__(self):
        self.master_file = ''
        self.secondary_file = ''
        # self.master = f"{os.path.dirname(os.path.abspath(__file__))}/FRS  MASTER FILE WITH HEARING DATES C & I COUNTY COURT (1) (1).xlsx"
        # self.format_file = f"{os.path.dirname(os.path.abspath(__file__))}/11_20_2020 BRWD FJ RAW.xlsx"
        self.master_case_numbers = []


    def setupUi(self, ExFileWindow):
        ExFileWindow.setObjectName("ExFileWindow")
        ExFileWindow.resize(750,300)
        ExFileWindow.setFixedSize(450, 225)
        self.centralwidget = QtWidgets.QWidget(ExFileWindow)
        self.centralwidget.setObjectName("centralwidget")
        self.updateButton = QtWidgets.QPushButton(self.centralwidget)
        self.updateButton.setGeometry(QtCore.QRect(320, 140, 113, 32))
        self.updateButton.setObjectName("updateButton")
        self.updateButton.clicked.connect(self.query_data_and_update_contents_excel_file)
        self.masterLabel = QtWidgets.QLabel(self.centralwidget)
        self.masterLabel.setGeometry(QtCore.QRect(29, 30, 71, 20))
        self.masterLabel.setObjectName("masterLabel")
        self.masterTextbox = QtWidgets.QLineEdit(self.centralwidget)
        self.masterTextbox.setGeometry(QtCore.QRect(110, 30, 271, 21))
        self.masterTextbox.setText("")
        self.masterTextbox.setObjectName("masterTextbox")
        self.masterFileButton = QtWidgets.QToolButton(self.centralwidget)
        self.masterFileButton.setGeometry(QtCore.QRect(390, 30, 41, 21))
        self.masterFileButton.setObjectName("masterFileButton")
        self.masterFileButton.clicked.connect(self.select_master_file_button) # run function in select master file
        self.secondaryLabel = QtWidgets.QLabel(self.centralwidget)
        self.secondaryLabel.setGeometry(QtCore.QRect(9, 90, 91, 20))
        self.secondaryLabel.setObjectName("secondaryLabel")
        self.secondaryTextBox = QtWidgets.QLineEdit(self.centralwidget)
        self.secondaryTextBox.setGeometry(QtCore.QRect(110, 90, 271, 21))
        self.secondaryTextBox.setObjectName("secondaryTextBox")
        self.secondaryFileButton = QtWidgets.QToolButton(self.centralwidget)
        self.secondaryFileButton.clicked.connect(self.select_secondary_file_button) # run function in select master file
        self.secondaryFileButton.setGeometry(QtCore.QRect(390, 90, 41, 21))
        self.secondaryFileButton.setObjectName("secondaryFileButton")
        self.exitButton = QtWidgets.QPushButton(self.centralwidget)
        self.exitButton.setGeometry(QtCore.QRect(210, 140, 113, 32))
        self.exitButton.setObjectName("exitButton")
        self.exitButton.clicked.connect(self.close_application_button)
        ExFileWindow.setCentralWidget(self.centralwidget)
        self.menubar = QtWidgets.QMenuBar(ExFileWindow)
        self.menubar.setGeometry(QtCore.QRect(0, 0, 452, 450))
        self.menubar.setObjectName("menubar")
        ExFileWindow.setMenuBar(self.menubar)
        self.statusbar = QtWidgets.QStatusBar(ExFileWindow)
        self.statusbar.setObjectName("statusbar")
        ExFileWindow.setStatusBar(self.statusbar)
        self.retranslateUi(ExFileWindow)
        QtCore.QMetaObject.connectSlotsByName(ExFileWindow)

    def retranslateUi(self, ExFileWindow):
        _translate = QtCore.QCoreApplication.translate
        ExFileWindow.setWindowTitle(_translate("ExFileWindow", "Excel File Updater"))
        self.updateButton.setText(_translate("ExFileWindow", "Update File"))
        self.masterLabel.setText(_translate("ExFileWindow", "Master File"))
        self.masterFileButton.setText(_translate("ExFileWindow", "..."))
        self.secondaryLabel.setText(_translate("ExFileWindow", "Secondary File"))
        self.secondaryFileButton.setText(_translate("ExFileWindow", "..."))
        self.exitButton.setText(_translate("ExFileWindow", "Exit"))

    def select_master_file_button(self):
        self.master_file = QtWidgets.QFileDialog.getOpenFileName()
        print(self.master_file[0])
        self.masterTextbox.setText(self.master_file[0])

    def select_secondary_file_button(self):
        self.secondary_file = QtWidgets.QFileDialog.getOpenFileName(None,"QFileDialog.getOpenFileName()")
        print(self.secondary_file[0])
        self.secondaryTextBox.setText(self.secondary_file[0])

    def close_application_button(self):
        print('closing')
        sys.exit()

    def query_data_and_update_contents_excel_file(self):
            file = pd.read_excel(self.secondary_file[0])
            master = pd.read_excel(self.master_file[0])
            file_case_num = file[['Case #']]
            for index, case_num in file_case_num.iterrows():
                case_number = case_num['Case #']
                update_cred = master.loc[master['Case Number'] == case_number]
                if update_cred.empty:
                    continue
                print('Updating Case Number ....' + case_number)
                address = update_cred['Address'].values[0]
                city = update_cred['City'].values[0]
                zip_code = update_cred['Zip Code'].values[0]
                state = update_cred['ST'].values[0]
                amount = update_cred['Code'].values[0]
                file.loc[index, 'Mailing Address'] = address
                file.loc[index, 'City'] = city
                file.loc[index, 'ST'] = state
                file.loc[index, 'Zip'] = zip_code
                file.loc[index, 'Amount $'] = amount

            with pd.ExcelWriter(self.secondary_file[0], mode='w') as writer:
                file.to_excel(writer)

            self.secondaryTextBox.setText('')
            msgbox = QtWidgets.QMessageBox()
            msgbox.setWindowTitle('Update')
            msgbox.setText("File Update Complete!")
            msgbox.exec_()




if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ExFileWindow = QtWidgets.QMainWindow()
    ui = Ui_ExFileWindow()
    ui.setupUi(ExFileWindow)
    ExFileWindow.show()
    sys.exit(app.exec_())
