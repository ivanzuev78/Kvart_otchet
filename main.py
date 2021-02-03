from Otchet_windows import *
import sys

if __name__ == '__main__':

    app = QtWidgets.QApplication(sys.argv)

    w = MainWindow()

    sys.exit(app.exec_())