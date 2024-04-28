import sys
from PyQt5.QtWidgets import QApplication
from excelCombiner import ExcelCombiner

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = ExcelCombiner()
    ex.show()
    sys.exit(app.exec_())
    