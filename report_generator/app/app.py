
__author__ = "Tafadzwa Brian Motsi"

from gui.gui import App
from PyQt5.QtWidgets import QApplication
import sys

if __name__ == '__main__':
    app = QApplication(sys.argv)
    ex = App()
    sys.exit(app.exec_())
