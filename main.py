import sys
from PyQt5.QtWidgets import QApplication
from database import init_db
from login import LoginWindow

if __name__ == "__main__":
    init_db()
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    login = LoginWindow()
    login.show()
    sys.exit(app.exec_())
