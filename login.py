from PyQt5 import QtWidgets
from PyQt5.QtWidgets import QWidget, QLabel, QLineEdit, QPushButton, QVBoxLayout, QMessageBox
from PyQt5.QtCore import Qt
from magasin import MagasinApp  # Import de la fenêtre principale

class LoginWindow(QWidget):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Gestion Magasin - DRB Alger")
        self.setFixedSize(420, 560)
        self.setWindowFlags(Qt.FramelessWindowHint)
        self.setAttribute(Qt.WA_TranslucentBackground)

        # Conteneur principal
        container = QtWidgets.QWidget(self)
        container.setStyleSheet("""
            background: #FFFFFF; border-radius: 20px; border: 1px solid #E0E0E0;
        """)
        layout = QVBoxLayout(container)
        layout.setContentsMargins(40,50,40,50)
        layout.setSpacing(20)

        # TITRE
        title = QLabel("DRB Alger")
        title.setAlignment(Qt.AlignCenter)
        title.setStyleSheet("font-size:28px; font-weight:bold; color:#1976D2;")
        subtitle = QLabel("Gestion De Magasin")
        subtitle.setAlignment(Qt.AlignCenter)
        subtitle.setStyleSheet("font-size:18px; color:#2E3B55;")
        desc = QLabel("Gestion Magasin")
        desc.setAlignment(Qt.AlignCenter)
        desc.setStyleSheet("font-size:14px; color:#555555;")

        # MOT DE PASSE
        self.password_input = QLineEdit()
        self.password_input.setPlaceholderText("Mot de passe...")
        self.password_input.setEchoMode(QLineEdit.Password)
        self.password_input.setFixedHeight(50)
        self.password_input.setStyleSheet("""
            QLineEdit { border:2px solid #C5CAE9; border-radius:12px; padding:0 16px; background:#FFF; }
            QLineEdit:focus { border-color:#1976D2; }
        """)

        # BOUTON ENTRER
        self.btn_enter = QPushButton("Entrer / دخول")
        self.btn_enter.setFixedHeight(50)
        self.btn_enter.setCursor(Qt.PointingHandCursor)
        self.btn_enter.setStyleSheet("""
            QPushButton {
                background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #1976D2, stop:1 #1E88E5);
                color:white; font-weight:bold; font-size:16px; border-radius:12px; border:none;
            }
            QPushButton:hover { background: qlineargradient(x1:0,y1:0,x2:1,y2:0, stop:0 #1E88E5, stop:1 #2196F3); }
            QPushButton:pressed { background: #1565C0; }
        """)
        self.btn_enter.clicked.connect(self.check_login)

        # COPYRIGHT
        copyright = QLabel("© 2025 DRBA - Tous droits réservés")
        copyright.setAlignment(Qt.AlignCenter)
        copyright.setStyleSheet("font-size:11px; color:#999999;")

        # AJOUT AU LAYOUT
        layout.addWidget(title)
        layout.addWidget(subtitle)
        layout.addWidget(desc)
        layout.addWidget(self.password_input)
        layout.addWidget(self.btn_enter)
        layout.addStretch()
        layout.addWidget(copyright)

        self.center()

    def center(self):
        qr = self.frameGeometry()
        cp = QtWidgets.QDesktopWidget().availableGeometry().center()
        qr.moveCenter(cp)
        self.move(qr.topLeft())

    def check_login(self):
        CORRECT_PASSWORD = "drb2025"

        
        if self.password_input.text() == CORRECT_PASSWORD:
            self.hide()
            self.main_app = MagasinApp()
            self.main_app.show()
        else:
            QMessageBox.critical(self, "Accès refusé", "Mot de passe incorrect.")
            self.password_input.clear()
            self.password_input.setFocus()
