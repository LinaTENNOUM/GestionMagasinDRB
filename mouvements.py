from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from datetime import datetime
from database import get_conn

class MouvementWindow(QWidget):
    def __init__(self, produit_id, nom_produit, stock_actuel):
        super().__init__()
        self.produit_id = produit_id
        self.stock_actuel = stock_actuel

        self.setWindowTitle(f"Mouvements – {nom_produit}")
        self.resize(420, 380)

        self.type = QComboBox()
        self.type.addItems(["ENTREE", "SORTIE"])

        self.qte = QSpinBox()
        self.qte.setRange(1, 1_000_000)

        self.service = QLineEdit()
        self.observation = QLineEdit()
        self.date = QLineEdit(datetime.now().strftime("%Y-%m-%d"))

        btn = QPushButton("Valider mouvement")
        btn.clicked.connect(self.save)

        form = QFormLayout()
        form.addRow("Type", self.type)
        form.addRow("Quantité", self.qte)
        form.addRow("Service", self.service)
        form.addRow("Date", self.date)
        form.addRow("Observation", self.observation)
        form.addRow(btn)

        self.setLayout(form)

    def save(self):
        qte = self.qte.value()
        mvt_type = self.type.currentText()

        if mvt_type == "SORTIE" and qte > self.stock_actuel:
            QMessageBox.warning(self, "Stock insuffisant", "Quantité supérieure au stock.")
            return

        conn = get_conn()
        c = conn.cursor()

        c.execute("""
            INSERT INTO mouvements (produit_id, type, quantite, date_mvt, service, observation)
            VALUES (?, ?, ?, ?, ?, ?)
        """, (self.produit_id, mvt_type, qte, self.date.text(),
              self.service.text(), self.observation.text()))

        if mvt_type == "ENTREE":
            c.execute("UPDATE produits SET quantite = quantite + ? WHERE id=?", (qte, self.produit_id))
        else:
            c.execute("UPDATE produits SET quantite = quantite - ? WHERE id=?", (qte, self.produit_id))

        conn.commit()
        conn.close()
        QMessageBox.information(self, "Succès", "Mouvement enregistré.")
        self.close()
