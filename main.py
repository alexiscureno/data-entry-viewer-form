import sys
from PyQt5.QtWidgets import QApplication, QPushButton, QLabel, QDialog, QSpinBox, QComboBox, QCheckBox
from PyQt5 import uic


class MainWindow(QDialog):
    def __init__(self):
        super().__init__()
        uic.loadUi('ui.ui', self)

        # Widgets
        # LineEdit
        self.name_ln = self.findChild(QLabel, 'name_edit')
        # SpinBox
        self.age_box = self.findChild(QSpinBox, "spinBox")
        # ComboBox
        self.sus_box = self.findChild(QComboBox, "comboBox")
        self.sus_box.addItem('Subscribed')
        self.sus_box.addItem('Unsubscribed')
        self.sus_box.addItem('Other')
        # CheckBox
        self.employ_box = self.findChild(QCheckBox, "checkBox")
        # pushButton
        self.insert_btn = self.findChild(QPushButton, "pushButton")


if __name__ == '__main__':
    app = QApplication(sys.argv)
    window = MainWindow()
    window.show()
    sys.exit(app.exec_())
