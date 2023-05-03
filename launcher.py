import sys
from PyQt5.QtWidgets import QMainWindow, QApplication, QPushButton, QLabel, QLineEdit, QWidget
from PyQt5.QtGui import QFont
import src.combinador


class Launcher(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(" ")
        self.setFixedSize(175, 200)

        font = QFont("Arial", 18)

        self.login_label = QLabel("Login", self)
        self.login_label.setGeometry(50, 25, 75, 25)
        self.login_label.setFont(font)

        self.login_field = QLineEdit(self)
        self.login_field.setGeometry(50, 75, 75, 25)
        self.login_field.setPlaceholderText("Login")

        self.pass_field = QLineEdit(self)
        self.pass_field.setGeometry(50, 105, 75, 25)
        self.pass_field.setPlaceholderText("Senha")
        self.pass_field.setEchoMode(QLineEdit.Password)

        self.login_btn = QPushButton("Login", self)
        self.login_btn.setGeometry(50, 150, 75, 25)
        self.login_btn.clicked.connect(self.func_login)

    def func_login(self):
        self.close()
        src.combinador.AppDemo().show()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    launcher = Launcher()
    launcher.show()
    sys.exit(app.exec())
