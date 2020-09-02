# This Python file uses the following encoding: utf-8
import sys
import os
import xls_processor, utils

from PySide2.QtWidgets import QApplication, QWidget, QHBoxLayout
from PySide2.QtCore import QFile
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import Slot



class MainWidget():
    def __init__(self):
        #QHBoxLayout * layout = new   QHBoxLayout;
        #self.setLayout(QHBoxLayout())
        self.load_ui()

    def load_ui(self):
        loader = QUiLoader()
        path = os.path.join(os.path.dirname(__file__), "MainUI/form.ui")
        ui_file = QFile(path)
        ui_file.open(QFile.ReadOnly)
        self.ui = loader.load(ui_file)
        ui_file.close()

        self.ui.pushButton_refresh_exceptions.clicked.connect(self.refresh_exceptions)

        return self.ui

    def show(self):
        self.ui.show()

    def refresh_exceptions(self):
        try:
            file = self.ui.textEdit_file.toPlainText()
            file = file.lstrip("file:///")
            to = xls_processor.XlsProcessor(file)
            r = to.calc_order_exceptions()
            t = utils.gen_exception_summary(r)
            self.ui.plainTextEdit_exceptions.setPlainText(t)
        except:
            self.ui.plainTextEdit_exceptions.setPlainText(str(sys.exc_info()[0]))

if __name__ == "__main__":
    app = QApplication([])
    mw = MainWidget()
    mw.show()

    sys.exit(app.exec_())