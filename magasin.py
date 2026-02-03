from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from datetime import datetime

from database import get_conn
from widgets import ModernComboBox, StyledItemDelegate
from export_utils import export_excel, export_pdf
from mouvements import MouvementWindow

DESTINATAIRES = [
    "CBW Alger", "CBW Boumerdes", "CBW Laghouat", "CBW Bouira",
    "CBW Blida", "CBW Djelfa", "CBW Medea", "CBW Tizi ouzou",
    "Bureau Informatique", "Bureau Suivi", "Bureau Personnel",
    "Bureau Comptabilité", "Bureau Moyen", "secretariat", 
    "Bureau Prevision", "Bureau Reglementation", "Bureau Formation",
    "Bureau Inspection", 
    "Autres"
]

class MagasinApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.selected_id = None
        self.setWindowTitle("Gestion Magasin - DRB Alger")
        self.resize(1200, 700)
        self.setMinimumSize(1000, 600)
        self.setStyleSheet("""
            QWidget { font-family: Segoe UI, Arial; font-size: 12px; }
            QLineEdit, QSpinBox, QDoubleSpinBox {
                padding: 6px; border: 1px solid #C5CAE9; border-radius: 8px; background: #FFFFFF;
            }
            QPushButton {
                background-color: #1976D2; color: white; padding: 8px 14px; border: none; border-radius: 10px;
            }
            QPushButton:hover { background-color: #1E88E5; }
            QPushButton:disabled { background-color: #90CAF9; }
            QTableWidget {
                background: #FFFFFF; gridline-color: #E3F2FD; alternate-background-color: #F5F9FF;
            }
            QHeaderView::section {
                background-color: #1976D2; color: white; padding: 6px; border: none;
            }
            QLabel.header { font-size: 18px; font-weight: 700; color: #1565C0; }
            QLabel.badge { color: #D32F2F; font-weight: 600; }
        """)

        # Widget central pour le contenu principal
        central = QWidget()
        self.setCentralWidget(central)
        main = QVBoxLayout(central)

        # === BARRE D'OUTILS EN HAUT (TOOLBAR) ===
        toolbar = QToolBar("Outils principaux")
        toolbar.setMovable(False)  # Fixe, non déplaçable
        self.addToolBar(Qt.TopToolBarArea, toolbar)

        # Bouton Historique Article
        btn_hist_article = QPushButton("Hist. Article")
        btn_hist_article.clicked.connect(self.ouvrir_historique_article)
        toolbar.addWidget(btn_hist_article)

        # Bouton Historique par Bureau/CB
        btn_hist_dest = QPushButton("Hist. par Bureau/CB")
        btn_hist_dest.clicked.connect(self.ouvrir_historique_par_destinataire)
        toolbar.addWidget(btn_hist_dest)

        # Bouton Affectation / Sortie
        self.btn_affecter = QPushButton("Affecter / Sortie")
        self.btn_affecter.setStyleSheet("""
            QPushButton {
                background-color: #D81B60; 
                color: white; 
                padding: 8px 16px; 
                border-radius: 10px;
            }
            QPushButton:hover { background-color: #E91E63; }
        """)
        self.btn_affecter.clicked.connect(self.open_affectation)
        self.btn_affecter.setEnabled(False)  # Désactivé par défaut
        toolbar.addWidget(self.btn_affecter)

        # === RECHERCHE + FILTRE NATURE ===
        self.search = QLineEdit()
        self.search.setPlaceholderText("Rechercher par désignation...")
        self.search.textChanged.connect(self.load_table)

        self.filter_nature = ModernComboBox()
        self.filter_nature.setMaximumWidth(300)
        self.filter_nature.addItem("Toutes les catégories", "")
        categories = [
            "MATERIELS INFORMATIQUES", "FOURNITURES DE BUREAUX", "PRODUITS D'ENTRETIEN MENNAGER",
            "HABILLEMENTS", "MOBILIER DE BUREAU", "PARC AUTO", "CONFECTION DES FOURNITURS IMPRIMEES",
            "CONSOMMABLE INFORMATIQUE", "PRODUITS PHARMACEUTIQUES", "EAUX"
        ]
        for cat in categories:
            self.filter_nature.addItem(cat, cat)
        self.filter_nature.currentIndexChanged.connect(self.load_table)
        self.filter_nature.setItemDelegate(StyledItemDelegate(self.filter_nature))

        # === TABLEAU ===
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Désignation", "Nature", "Quantité", "Prix", "Seuil mini", "Date ajout", "Observation"
        ])
        self.table.setAlternatingRowColors(True)
        self.table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        self.table.setSelectionBehavior(QTableWidget.SelectRows)
        self.table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.table.clicked.connect(self.on_row_click)

        # === FORMULAIRE ===
        self.nom = QLineEdit()
        self.nom.setMaximumWidth(280)
        self.nature = ModernComboBox()
        self.nature.setMaximumWidth(280)
        for cat in categories:
            self.nature.addItem(cat, cat)
        self.nature.setItemDelegate(StyledItemDelegate(self.nature))

        self.quantite = QSpinBox(); self.quantite.setRange(0, 10_000_000)
        self.prix = QDoubleSpinBox(); self.prix.setRange(0, 10_000_000); self.prix.setDecimals(2); self.prix.setSingleStep(1.0)
        self.seuil = QSpinBox(); self.seuil.setRange(0, 10_000_000)
        self.date = QLineEdit(datetime.now().strftime("%Y-%m-%d"))
        self.date.setPlaceholderText("YYYY-MM-DD")
        self.date.setMaximumWidth(150)
        self.observation = QLineEdit()
        self.observation.setPlaceholderText("Ex: À commander, en panne, etc.")
        self.observation.setMaximumWidth(280)

        # === BOUTONS (CRUD en bas, sans les boutons déplacés en haut) ===
        self.btn_add = QPushButton("Ajouter"); self.btn_add.clicked.connect(self.add_product)
        self.btn_update = QPushButton("Modifier"); self.btn_update.clicked.connect(self.update_product); self.btn_update.setEnabled(False)
        self.btn_delete = QPushButton("Supprimer"); self.btn_delete.clicked.connect(self.delete_product); self.btn_delete.setEnabled(False)
        self.btn_clear = QPushButton("Nouveau / Vider"); self.btn_clear.clicked.connect(self.clear_form)
        self.btn_export_xlsx = QPushButton("Exporter Excel (.xlsx)"); self.btn_export_xlsx.clicked.connect(self.on_export_excel)
        self.btn_export_pdf = QPushButton("Exporter PDF"); self.btn_export_pdf.clicked.connect(self.on_export_pdf)
        self.badge_low = QLabel("")

        # === LAYOUTS ===
        top = QHBoxLayout()
        lbl = QLabel("Inventaire")
        lbl.setProperty("class", "header")
        top.addWidget(lbl)
        top.addStretch()
        top.addWidget(self.badge_low)

        search_bar = QHBoxLayout()
        search_bar.addWidget(QLabel("Recherche:"))
        search_bar.addWidget(self.search)
        search_bar.addWidget(QLabel("Filtre par Nature:"))
        search_bar.addWidget(self.filter_nature)
        search_bar.addStretch()

        form = QFormLayout()
        form.addRow("Désignation *", self.nom)
        form.addRow("Nature", self.nature)
        form.addRow("Observation", self.observation)

        hnums = QHBoxLayout()
        hnums.addWidget(self._labeled("Quantité", self.quantite))
        hnums.addWidget(self._labeled("Prix", self.prix))
        hnums.addWidget(self._labeled("Seuil mini", self.seuil))
        hnums.addWidget(self._labeled("Date ajout", self.date))

        crud = QHBoxLayout()
        crud.addWidget(self.btn_add)
        crud.addWidget(self.btn_update)
        crud.addWidget(self.btn_delete)
        crud.addWidget(self.btn_clear)
        crud.addStretch()
        crud.addWidget(self.btn_export_xlsx)
        crud.addWidget(self.btn_export_pdf)

        main.addLayout(top)
        main.addLayout(search_bar)
        main.addWidget(self.table)
        main.addLayout(form)
        main.addLayout(hnums)
        main.addLayout(crud)

        self.load_table()
        self.update_badge()

    def _labeled(self, text, widget):
        container = QWidget()
        vbox = QVBoxLayout(container)
        vbox.setContentsMargins(0, 0, 0, 0)
        vbox.setSpacing(4)
        
        lbl = QLabel(text)
        lbl.setStyleSheet("font-weight: bold; color: #2E3B55;")
        
        vbox.addWidget(lbl)
        vbox.addWidget(widget)
        
        return container

    def load_table(self):
        search_text = self.search.text().strip().lower()
        nature_filter = self.filter_nature.currentData()

        conn = get_conn()
        c = conn.cursor()
        query = "SELECT nom, nature, quantite, prix, seuil_min, date_ajout, observation, id FROM produits WHERE 1=1"
        params = []

        if search_text:
            query += " AND LOWER(nom) LIKE ?"
            params.append(f"%{search_text}%")

        if nature_filter:
            query += " AND nature = ?"
            params.append(nature_filter)

        query += " ORDER BY date_ajout DESC"

        c.execute(query, params)
        rows = c.fetchall()
        conn.close()

        self.table.setRowCount(len(rows))

        for i, row in enumerate(rows):
            for j in range(7):
                item = QTableWidgetItem(str(row[j]) if row[j] is not None else "")
                if j in (2, 4):  # Quantité, Seuil
                    item.setTextAlignment(Qt.AlignCenter)
                if j == 3:  # Prix
                    item.setTextAlignment(Qt.AlignRight)
                self.table.setItem(i, j, item)
            # ID caché dans userData
            self.table.item(i, 0).setData(Qt.UserRole, row[7])

            # Colorer en rouge si quantite < seuil_min
            qte = int(row[2] or 0)
            seuil = int(row[4] or 0)
            if qte < seuil:
                for j in range(7):
                    self.table.item(i, j).setForeground(QColor("#D32F2F"))

        self.table.resizeRowsToContents()
        self.update_badge()

    def on_row_click(self, index):
        row = index.row()
        self.nom.setText(self.table.item(row, 0).text())
        nature_text = self.table.item(row, 1).text()
        idx = self.nature.findText(nature_text)
        if idx >= 0:
            self.nature.setCurrentIndex(idx)
        self.quantite.setValue(int(self.table.item(row, 2).text() or 0))
        self.prix.setValue(float(self.table.item(row, 3).text() or 0))
        self.seuil.setValue(int(self.table.item(row, 4).text() or 0))
        self.date.setText(self.table.item(row, 5).text())
        self.observation.setText(self.table.item(row, 6).text())
        self.selected_id = self.table.item(row, 0).data(Qt.UserRole)
        self.btn_update.setEnabled(True)
        self.btn_delete.setEnabled(True)
        self.btn_affecter.setEnabled(True)

    def clear_form(self):
        self.nom.clear()
        self.nature.setCurrentIndex(0)
        self.quantite.setValue(0)
        self.prix.setValue(0)
        self.seuil.setValue(0)
        self.date.setText(datetime.now().strftime("%Y-%m-%d"))
        self.observation.clear()
        self.selected_id = None
        self.btn_update.setEnabled(False)
        self.btn_delete.setEnabled(False)
        self.btn_affecter.setEnabled(False)

    def add_product(self):
        if not self.nom.text().strip():
            QMessageBox.warning(self, "Champ requis", "La désignation est obligatoire.")
            return

        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            INSERT INTO produits (nom, nature, quantite, prix, seuil_min, date_ajout, observation)
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            self.nom.text(),
            self.nature.currentData(),
            self.quantite.value(),
            self.prix.value(),
            self.seuil.value(),
            self.date.text(),
            self.observation.text()
        ))
        conn.commit()
        conn.close()
        QMessageBox.information(self, "Succès", "Produit ajouté.")
        self.load_table()
        self.clear_form()

    def update_product(self):
        if not self.selected_id or not self.nom.text().strip():
            return

        conn = get_conn()
        c = conn.cursor()
        c.execute("""
            UPDATE produits SET nom=?, nature=?, quantite=?, prix=?, seuil_min=?, date_ajout=?, observation=?
            WHERE id=?
        """, (
            self.nom.text(),
            self.nature.currentData(),
            self.quantite.value(),
            self.prix.value(),
            self.seuil.value(),
            self.date.text(),
            self.observation.text(),
            self.selected_id
        ))
        conn.commit()
        conn.close()
        QMessageBox.information(self, "Succès", "Produit modifié.")
        self.load_table()
        self.clear_form()

    def delete_product(self):
        if not self.selected_id:
            return

        if QMessageBox.question(self, "Confirmer", "Supprimer ce produit ?") != QMessageBox.Yes:
            return

        conn = get_conn()
        c = conn.cursor()
        c.execute("DELETE FROM produits WHERE id=?", (self.selected_id,))
        conn.commit()
        conn.close()
        QMessageBox.information(self, "Succès", "Produit supprimé.")
        self.load_table()
        self.clear_form()

    def on_export_excel(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exporter Excel", "", "Excel (*.xlsx)")
        if path:
            conn = get_conn()
            c = conn.cursor()
            c.execute("SELECT nom, nature, quantite, prix, seuil_min, date_ajout, observation FROM produits")
            rows = c.fetchall()
            conn.close()
            export_excel(rows, path)
            QMessageBox.information(self, "Exporté", "Fichier Excel créé.")

    def on_export_pdf(self):
        path, _ = QFileDialog.getSaveFileName(self, "Exporter PDF", "", "PDF (*.pdf)")
        if path:
            conn = get_conn()
            c = conn.cursor()
            c.execute("SELECT nom, nature, quantite, prix, seuil_min, date_ajout, observation FROM produits")
            rows = c.fetchall()
            conn.close()
            export_pdf(rows, path)
            QMessageBox.information(self, "Exporté", "Fichier PDF créé.")

    def update_badge(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM produits WHERE quantite < seuil_min AND seuil_min > 0")
        low = c.fetchone()[0]
        conn.close()
        if low > 0:
            self.badge_low.setText(f"{low} articles en stock bas")
            self.badge_low.setProperty("class", "badge")
            self.badge_low.setStyleSheet("QLabel.badge { color: #D32F2F; font-weight: 600; }")
        else:
            self.badge_low.setText("")

    def open_affectation(self):
        if not self.selected_id:
            QMessageBox.warning(self, "Sélection", "Sélectionnez un article d'abord.")
            return

        nom = self.nom.text()
        stock = self.quantite.value()
        win = MouvementWindow(self.selected_id, nom, stock)
        win.show()
        win.closed.connect(self.load_table)  # Rafraîchir après fermeture (ajoute ça si tu as signal closed, sinon ignore)

    def ouvrir_historique_article(self):
        if not self.selected_id:
            QMessageBox.warning(self, "Sélection", "Sélectionnez un article pour voir son historique.")
            return

        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT nom FROM produits WHERE id=?", (self.selected_id,))
        nom = c.fetchone()[0]
        conn.close()

        self._ouvrir_fenetre_historique(
            title=f"Historique - {nom}",
            filtre_article=self.selected_id,
            prefiltre_article=True
        )

    def ouvrir_historique_par_destinataire(self):
        dest, ok = QInputDialog.getItem(
            self,
            "Filtrer par destinataire",
            "Choisir le bureau ou CB :",
            ["Tous"] + DESTINATAIRES,  # "Tous" pour aucun filtre
            0,
            False
        )

        if not ok:
            return

        title = f"Historique - {dest}" if dest != "Tous" else "Historique complet tous destinataires"

        self._ouvrir_fenetre_historique(
            title=title,
            filtre_destinataire=dest if dest != "Tous" else None
        )

    def _ouvrir_fenetre_historique(self, title="Historique des Mouvements", 
                                 filtre_article=None, prefiltre_article=False,
                                 filtre_destinataire=None):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(1100, 680)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(10, 10, 10, 10)

        # Barre de filtres rapide
        filtres = QHBoxLayout()

        lbl_article = QLabel("Article :")
        edit_article = QLineEdit()
        if prefiltre_article:
            edit_article.setText(title.split(" - ")[-1])  # nom déjà connu
            edit_article.setReadOnly(True)

        lbl_dest = QLabel("Destinataire :")
        combo_dest = QComboBox()
        combo_dest.addItem("Tous", "")
        for d in DESTINATAIRES:
            combo_dest.addItem(d, d)
        if filtre_destinataire:
            index = combo_dest.findData(filtre_destinataire)
            if index >= 0:
                combo_dest.setCurrentIndex(index)

        lbl_type = QLabel("Type :")
        combo_type = QComboBox()
        combo_type.addItems(["Tous", "ENTREE", "SORTIE"])

        filtres.addWidget(lbl_article)
        filtres.addWidget(edit_article)
        filtres.addSpacing(15)
        filtres.addWidget(lbl_dest)
        filtres.addWidget(combo_dest)
        filtres.addSpacing(15)
        filtres.addWidget(lbl_type)
        filtres.addWidget(combo_type)
        filtres.addStretch()

        layout.addLayout(filtres)

        # Tableau
        table = QTableWidget()
        table.setColumnCount(8)
        table.setHorizontalHeaderLabels([
            "Date", "Type", "Article", "Qté", "Destinataire", "Observation", "Stock après", "ID"
        ])
        table.setAlternatingRowColors(True)
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        table.verticalHeader().setVisible(False)
        table.setSelectionBehavior(QTableWidget.SelectRows)
        table.setEditTriggers(QTableWidget.NoEditTriggers)

        layout.addWidget(table)

        def charger():
            article_txt = edit_article.text().strip()
            dest_val = combo_dest.currentData()
            type_val = combo_type.currentText()

            conn = get_conn()
            c = conn.cursor()

            query = """
                SELECT 
                    m.date_mvt,
                    m.type,
                    p.nom,
                    m.quantite,
                    m.service,
                    m.observation,
                    p.quantite AS stock_apres,
                    m.id
                FROM mouvements m
                JOIN produits p ON m.produit_id = p.id
                WHERE 1=1
            """
            params = []

            if filtre_article is not None:
                query += " AND m.produit_id = ?"
                params.append(filtre_article)
            elif article_txt:
                query += " AND p.nom LIKE ?"
                params.append(f"%{article_txt}%")

            if dest_val:
                query += " AND m.service = ?"
                params.append(dest_val)
            elif filtre_destinataire:
                query += " AND m.service = ?"
                params.append(filtre_destinataire)

            if type_val != "Tous":
                query += " AND m.type = ?"
                params.append(type_val)

            query += " ORDER BY m.date_mvt DESC LIMIT 2000"

            c.execute(query, params)
            rows = c.fetchall()
            conn.close()

            table.setRowCount(len(rows))

            for i, row in enumerate(rows):
                for j, val in enumerate(row):
                    item = QTableWidgetItem(str(val) if val is not None else "")
                    
                    if j == 1:  # Type
                        color = "#2E7D32" if val == "ENTREE" else "#D32F2F" if val == "SORTIE" else "#555"
                        item.setForeground(QColor(color))
                    
                    if j == 3:  # Quantité
                        item.setTextAlignment(Qt.AlignCenter)
                        if row[1] == "SORTIE":
                            item.setText(f"-{val}")
                    
                    if j == 6:  # Stock après
                        item.setTextAlignment(Qt.AlignRight)
                    
                    table.setItem(i, j, item)

            table.resizeRowsToContents()

        # Connexions pour rafraîchissement automatique
        edit_article.textChanged.connect(charger)
        combo_dest.currentIndexChanged.connect(charger)
        combo_type.currentIndexChanged.connect(charger)

        # Chargement initial
        charger()

        dialog.exec_()