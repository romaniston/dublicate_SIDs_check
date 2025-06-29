import sys

from PyQt5.QtWidgets import QApplication

from modules import gui


app = QApplication(sys.argv)
window = gui.MainWindow()
window.show()
sys.exit(app.exec_())


# TODO 1. "creating" table func (FIX BUGS)
# TODO 2. exceptions and routers when table is not exist and offer to create it
# TODO 3. realize "get left sids"
# TODO 4. test
# TODO 5. refactoring
# TODO 6. final test