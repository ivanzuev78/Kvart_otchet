import sys

from Otchet_windows import *

if __name__ == "__main__":

    app = QtWidgets.QApplication(sys.argv)

    w = MainWindow()

    sys.exit(app.exec_())
