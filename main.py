import sys
from PyQt5.QtWidgets import QApplication, QPushButton, QLabel, QDialog, QSpinBox, QComboBox, QCheckBox,\
    QMainWindow, QAction, QTableWidget, QFileDialog, QMenuBar, QTableWidgetItem, QWidgetAction
from PyQt5 import uic
import openpyxl
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
import pandas as pd


class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        uic.loadUi('ui.ui', self)

        # Widgets

        # File menu bar
        self.new_file_bar = self.findChild(QAction, 'actionNew_File')
        self.open_file_bar = self.findChild(QAction, 'actionOpen_File')

        # LineEdit
        self.name_ln = self.findChild(QLabel, 'name_edit')

        # SpinBox
        self.age_box = self.findChild(QSpinBox, "spinBox")

        # Table Widget
        self.table_widget = self.findChild(QTableWidget, 'tableWidget')
        # ComboBox
        self.sus_box = self.findChild(QComboBox, "comboBox")
        self.sus_box.addItem('Subscribed')
        self.sus_box.addItem('Unsubscribed')
        self.sus_box.addItem('Other')

        # CheckBox
        self.employ_box = self.findChild(QCheckBox, "checkBox")

        # pushButton
        self.insert_btn = self.findChild(QPushButton, "pushButton")

        # Actions menu bar
        self.new_file_bar.triggered.connect(self.new_file)
        self.open_file_bar.triggered.connect(self.open_file)

    def new_file(self):
        wb = Workbook()
        new_file_name = QFileDialog.getSaveFileName(
            parent=self,
            caption=' Save file location and name',
            directory='',
            filter='Data File (*.xlsx *.csv*.) ;; Excel File (*.xlsx *.xls)',
            initialFilter='Excel File (*.xlsx *.xls)'
        )

    def open_file(self):
        opened_file_name = QFileDialog.getOpenFileName(
            parent = self,
            caption = 'Open file',
            directory = '',
            filter = 'Data File (*.xlsx *.csv*.) ;; Excel File (*.xlsx *.xls)',
            initialFilter = 'Excel File (*.xlsx *.xls)'
        )
        wb_ = openpyxl.load_workbook(opened_file_name[0])
        ws_ = wb_.active

        self.load_data(opened_file_name[0])

    def load_data(self, path_file):
        path = path_file
        df = pd.read_excel(path)

        cols = list(df.columns)
        rows = df.to_numpy().tolist()
        x = len(cols)
        y = len(rows)
        #print(cols)
        #print(y)

        self.tableWidget.setRowCount(y)
        self.tableWidget.setColumnCount(x)

        for j in range(x):
            print(cols[j])
            header = QTableWidgetItem(cols[j])
            self.tableWidget.setHorizontalHeaderItem(j, header)

            for i in range(y):
                data = str(rows[i][j])
                if data == 'nan':
                    data = ''
                self.tableWidget.setItem(i,j, QTableWidgetItem(data))

if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
