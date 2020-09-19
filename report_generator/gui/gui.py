__author__ = "Tafadzwa Brian Motsi"

from document_with_student_details.document_with_student_details import DocumentWithStudentDetails
from read_student_details.read_student_details import StudentDetails

from PyQt5.QtWidgets import QWidget, QFileDialog, QPushButton
from PyQt5 import QtCore, QtWidgets
from PyQt5.QtCore import pyqtSlot
from PyQt5.QtGui import QIcon
import os

document_object = DocumentWithStudentDetails()
student_details_object = StudentDetails()


class App(QWidget):

    def __init__(self):
        super().__init__()
        self.title = 'REPORT GENERATOR'
        self.setWindowIcon(QIcon(r'../icons/iconfinder_logo_brand_brands_logos_total_commander_3215607.png'))
        self.left = 400
        self.top = 400
        self.width = 400
        self.height = 80
        self.init_ui()
        self.setFixedSize(self.size())

    def init_ui(self):
        self.setWindowTitle(self.title)
        self.setGeometry(self.left, self.top, self.width, self.height)
        self.setStyleSheet(
            'background-color: rgb(255, 255, 255);'
        )
        self.open_file_button()
        self.show()

    @pyqtSlot()
    def on_click_1(self):
        self.open_file_name_dialog()

    def open_button(self, label, tool_tip_text, move_x, move_y, resize_x, resize_y):
        button = QPushButton(label, self)
        button.setToolTip(tool_tip_text)
        button.move(move_x, move_y)
        button.clicked.connect(self.on_click_1)
        button.resize(resize_x, resize_y)
        button.setStyleSheet(
            'background-color: hsl(0, 100%, 5%);'
            'border-style: outset;'
            'border-width: 4px;'
            'border-radius: 200px;'
            'border-color: beige;'
            'font: bold 14px;'
            'min-width: 10em;'
            'padding: 10px;'
            'color: white;'
        )

    def open_file_button(self):
        self.open_button(
            'OPEN THE INPUT FILE TO GENERATE REPORTS',
            'Navigate your system to find the input file\nThe input file MUST be a .xlsx file!',
            0, 0,
            400, 50,
        )

    def open_file_name_dialog(self):
        options = QFileDialog.Options()
        options |= QFileDialog.DontUseNativeDialog
        file_name, _ = QFileDialog.getOpenFileName(self,
                                                   "Open File",
                                                   "",
                                                   "Spread Sheet Files (*.xlsx)",
                                                   options=options)

        _dialog = QtWidgets.QInputDialog(self)
        _dialog.resize(QtCore.QSize(400, 300))
        _dialog.setWindowTitle("Class Grade Input Dialog")
        _dialog.setLabelText("Enter Class Grade")
        _dialog.setTextValue("")
        _dialog.setTextEchoMode(QtWidgets.QLineEdit.Normal)

        dialog = QtWidgets.QInputDialog(self)
        dialog.resize(QtCore.QSize(400, 200))
        dialog.setWindowTitle("Date Input Dialog")
        dialog.setLabelText("Enter Next Term's Open Date")
        dialog.setTextValue("")
        dialog.setTextEchoMode(QtWidgets.QLineEdit.Normal)
        if file_name and dialog.exec_() == QtWidgets.QDialog.Accepted and _dialog.exec_() == QtWidgets.QDialog.Accepted:
            file_path = os.path.abspath(file_name)
            document_object.generate_documents(student_details_object.student_details(file_name),
                                               "Times New Roman", _dialog.textValue(),
                                               r'\\'.join(file_path.split('\\')[:-1]),
                                               str(dialog.textValue()))
