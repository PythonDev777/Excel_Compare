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
        self.new_file = ''

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
        self.secondaryLabel.setGeometry(QtCore.QRect(9, 70, 91, 20))
        self.secondaryLabel.setObjectName("secondaryLabel")
        self.secondaryTextBox = QtWidgets.QLineEdit(self.centralwidget)
        self.secondaryTextBox.setGeometry(QtCore.QRect(110, 70, 271, 21))
        self.secondaryTextBox.setObjectName("secondaryTextBox")
        self.secondaryFileButton = QtWidgets.QToolButton(self.centralwidget)
        self.secondaryFileButton.clicked.connect(self.select_secondary_file_button) # run function in select master file
        self.secondaryFileButton.setGeometry(QtCore.QRect(390, 70, 41, 21))
        self.secondaryFileButton.setObjectName("secondaryFileButton")
        self.newFileLabel = QtWidgets.QLabel(self.centralwidget)
        self.newFileLabel.setGeometry(QtCore.QRect(29, 110, 71, 20))
        self.newFileLabel.setObjectName("newFileLabel")
        self.newFileTextbox = QtWidgets.QLineEdit(self.centralwidget)
        self.newFileTextbox.setGeometry(QtCore.QRect(110, 110, 271, 21))
        self.newFileTextbox.setText("")
        self.newFileTextbox.setObjectName("newFileTextbox")
        self.newFileButton = QtWidgets.QToolButton(self.centralwidget)
        self.newFileButton.setGeometry(QtCore.QRect(390, 110, 41, 21))
        self.newFileButton.setObjectName("newFileButton")
        self.newFileButton.clicked.connect(self.select_new_file_button) # run function in select master file
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
        self.newFileLabel.setText(_translate("ExFileWindow", "New File"))
        self.newFileButton.setText(_translate("ExFileWindow", "..."))
        self.exitButton.setText(_translate("ExFileWindow", "Exit"))

    def select_master_file_button(self):
        self.master_file = QtWidgets.QFileDialog.getOpenFileName()
        print(self.master_file[0])
        self.masterTextbox.setText(self.master_file[0])

    def select_secondary_file_button(self):
        self.secondary_file = QtWidgets.QFileDialog.getOpenFileName(None,"QFileDialog.getOpenFileName()")
        print(self.secondary_file[0])
        self.secondaryTextBox.setText(self.secondary_file[0])

    def select_new_file_button(self):
        self.new_file = QtWidgets.QFileDialog.getSaveFileName()
        self.newFileTextbox.setText(self.new_file[0])

    def close_application_button(self):
        print('closing')
        sys.exit()

    def query_data_and_update_contents_excel_file(self):
        file = pd.read_excel(self.secondary_file[0])
        master = pd.read_excel(self.master_file[0])
        file_case_num = file[['Case #']]
        df = ''
        for index, case_num in file_case_num.iterrows():
            case_number = case_num['Case #']
            update_cred = master.loc[master['Case Number'] == case_number]
            file_row_content = file.loc[file['Case #'] == case_number]
            date = pd.to_datetime(file_row_content["Date"]).dt.strftime("%m-%d-%Y")
            file.loc[index, 'Date'] = date.values[0]
            if update_cred.empty:
                continue
            print('Updating Case Number ....' + case_number)
            file.loc[index, 'Amount $'] = update_cred['Code'].values[0]
            file.loc[index, 'Case #'] = case_number
            file.loc[index, 'Mailing Address'] = update_cred['Address'].values[0]
            file.loc[index, 'City'] = update_cred['City'].values[0]
            file.loc[index, 'ST'] = update_cred['ST'].values[0]
            file.loc[index, 'Zip'] = update_cred['Zip Code'].values[0]
            df = file.sort_values(['Mailing Address', 'City', 'ST', 'Zip'], na_position='last')

        new_file_with_extension = self.new_file[0] + '.xlsx'
        with pd.ExcelWriter(new_file_with_extension, mode='w') as writer:
            df.to_excel(writer)

        msgbox = QtWidgets.QMessageBox()
        msgbox.setWindowTitle('Update')
        msgbox.setText("File Update Complete!")
        msgbox.exec_()
        self.secondaryTextBox.setText('')
        self.newFileTextbox.setText('')
        

if __name__ == "__main__":
    import sys
    app = QtWidgets.QApplication(sys.argv)
    ExFileWindow = QtWidgets.QMainWindow()
    ui = Ui_ExFileWindow()
    ui.setupUi(ExFileWindow)
    ExFileWindow.show()
    sys.exit(app.exec_())
