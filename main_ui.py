# This Python file uses the following encoding: utf-8
import sys
import os
import xls_processor, utils

from PySide2.QtWidgets import QApplication, QWidget, QHBoxLayout
from PySide2.QtCore import QFile
from PySide2.QtUiTools import QUiLoader
from PySide2.QtCore import Slot



class MainWidget(QWidget):
    def __init__(self):
        super(MainWidget, self).__init__()
        #QHBoxLayout * layout = new   QHBoxLayout;
        #self.setLayout(QHBoxLayout())
        self.load_ui()

    def load_ui(self):
        loader = QUiLoader()
        path = os.path.join(os.path.dirname(__file__), "MainUI/form.ui")
        ui_file = QFile(path)
        ui_file.open(QFile.ReadOnly)
        self.ui = loader.load(ui_file, self)
        ui_file.close()

        self.ui.pushButton_refresh_exceptions.clicked.connect(self.refresh_exceptions)
        self.setAcceptDrops(True)

    def refresh_exceptions(self):

        file = self.ui.textEdit_file.toPlainText()
        print(file)
        file = file.lstrip("file:///")
        to = xls_processor.XlsProcessor(file)
        r = to.calc_order_exceptions()
        t = utils.gen_exception_summary(r)
        self.ui.plainTextEdit_exceptions.setPlainText(t)


if __name__ == "__main__":
    app = QApplication([])
    widget = MainWidget()
    widget.show()
    sys.exit(app.exec_())