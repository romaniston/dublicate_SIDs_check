import sys
from PyQt5.QtWidgets import QApplication
from modules import gui


app = QApplication(sys.argv)
window = gui.MainWindow()
window.show()
sys.exit(app.exec_())
