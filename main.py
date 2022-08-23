import sys
from tkinter import W
import traceback
import os
from unittest import result
import utils
from PyQt6 import QtGui
from PyQt6.QtWidgets import (
    QApplication,
    QMainWindow,
    QWidget,
    QPushButton, 
    QFileDialog, 
    QVBoxLayout,
    QHBoxLayout,
    QLineEdit,
    QLabel,
    QMessageBox,
    QTextEdit
    )

class MyApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle('Mail Shoot')


        self.spreadsheet_label = QLabel(self)
        self.spreadsheet_label.setText('Data file:')
        self.spreadsheet_filepath_input = QLineEdit(self)
        self.spreadsheet_filepath_input.textChanged.connect(self.set_spreadsheet)
        self.select_spreadsheet_button = QPushButton('...')
        self.select_spreadsheet_button.clicked.connect(self.get_spreadsheet_filepath)


        self.cc_label = QLabel(self)
        self.cc_label.setText('Cc:')
        self.cc_input = QLineEdit(self)
        self.cc_input.textChanged.connect(self.set_cc)


        self.subject_label = QLabel(self)
        self.subject_label.setText('Subject:')
        self.subject_input = QLineEdit(self)
        self.subject_input.textChanged.connect(self.set_subject)


        self.body_label = QLabel(self)
        self.body_label.setText('Body:')
        self.body_input = QTextEdit(self)
        self.body_input.textChanged.connect(self.set_body)


        self.attachments_label = QLabel(self)
        self.attachments_label.setText('Attachments:')
        self.attachments_filepath_input = QLineEdit(self)
        self.attachments_filepath_input.textChanged.connect(self.set_attachments)
        self.select_attachments_button = QPushButton('...')
        self.select_attachments_button.clicked.connect(self.get_attachments_filepath)


        self.run_button = QPushButton('Run')
        self.run_button.clicked.connect(self.run)


        self.spreadsheet = ''
        self.attachments = []
        self.receivers = []
        self.cc = ''
        self.subject = ''
        self.body = ''

        spreadsheet_hbox = QHBoxLayout()
        spreadsheet_hbox.addWidget(self.spreadsheet_label)
        spreadsheet_hbox.addWidget(self.spreadsheet_filepath_input)
        spreadsheet_hbox.addWidget(self.select_spreadsheet_button)

        cc_hbox = QHBoxLayout()
        cc_hbox.addWidget(self.cc_label)
        cc_hbox.addWidget(self.cc_input)

        subject_hbox = QHBoxLayout()
        subject_hbox.addWidget(self.subject_label)
        subject_hbox.addWidget(self.subject_input)

        attachments_hbox = QHBoxLayout()
        attachments_hbox.addWidget(self.attachments_label)
        attachments_hbox.addWidget(self.attachments_filepath_input)
        attachments_hbox.addWidget(self.select_attachments_button)

        body_vbox = QVBoxLayout()
        body_vbox.addWidget(self.body_label)
        body_vbox.addWidget(self.body_input)

        vbox_main = QVBoxLayout()
        vbox_main.addLayout(spreadsheet_hbox)
        vbox_main.addLayout(cc_hbox)
        vbox_main.addLayout(subject_hbox)
        vbox_main.addLayout(body_vbox)
        vbox_main.addLayout(attachments_hbox)
        vbox_main.addWidget(self.run_button)

        container = QWidget()
        container.setLayout(vbox_main)


        self.setCentralWidget(container)

    def get_spreadsheet_filepath(self):
        response = QFileDialog.getOpenFileName(
            parent=self,
            caption='Select a data file',
            directory=os.getcwd(),
            filter='Data File (*.xlsx *.csv *.dat);; Excel File (*.xlsx *.xls)',
            initialFilter='Excel File (*.xlsx *.xls)'
        )
        self.spreadsheet_filepath_input.setText(response[0])
        return response[0]

    def get_attachments_filepath(self):
        response = QFileDialog.getOpenFileNames(
            parent=self,
            caption='Select a file',
            directory=os.getcwd(),
            filter='All (*.*)',
            initialFilter='All (*.*)'
        )
        self.attachments_filepath_input.setText(','.join(response[0]))
        return response[0]

    def set_spreadsheet(self):
        self.spreadsheet = self.spreadsheet_filepath_input.text()

    def set_receivers(self):
        self.receivers = ([receiver.strip() for receiver 
                            in self.receivers_input.text().split(',')]
            if self.receivers_input.text() and self.receivers_input.text() != ''
            else [])

    def set_cc(self):
        self.cc = self.cc_input.text()

    def set_subject(self):
        self.subject = self.subject_input.text()

    def set_body(self):
        self.body = self.body_input.toPlainText()

    def set_attachments(self):
        self.attachments = self.attachments_filepath_input.text().split(',')

    def run(self):
        res = utils.send_mails(
            xl_path=self.spreadsheet, 
            cc=self.cc,
            subject=self.subject,
            body=self.body,
            attachments=self.attachments
            )
        print(f"Result: {res}")
        if res:
            QMessageBox.information(self, 'Result', 'Done!')
        else:
            QMessageBox.warning(self, 'Result', 'Error! Please see the log...')
    
def error_handler(etype, value, tb):
    error_msg = ''.join(traceback.format_exception(etype, value, tb))
    print(error_msg)

if __name__ == '__main__':
    
    app = QApplication(sys.argv)
    
    myApp = MyApp()
    myApp.show()

    try:
        sys.excepthook = error_handler
        sys.exit(app.exec())
    except Exception:
        print('Closing Window...')
