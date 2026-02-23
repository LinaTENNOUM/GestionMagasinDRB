from PyQt5.QtWidgets import *
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QColor
from datetime import datetime

from database import get_conn
from widgets import ModernComboBox, StyledItemDelegate
from export_utils import export_excel, export_pdf, export_history_excel, export_history_pdf

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

        # Widget central
        central = QWidget()
        self.setCentralWidget(central)
        main = QVBoxLayout(central)
        main.setContentsMargins(15, 15, 15, 15)
        main.setSpacing(12)

        # === TOOLBAR EN HAUT ===
        toolbar = QToolBar("Outils")
        toolbar.setMovable(False)
        self.addToolBar(Qt.TopToolBarArea, toolbar)

        btn_hist_article = QPushButton("Historique par article")
        btn_hist_article.clicked.connect(self.ouvrir_historique_article)
        toolbar.addWidget(btn_hist_article)

        btn_hist_dest = QPushButton("Historique par destinataire")
        btn_hist_dest.clicked.connect(self.ouvrir_historique_par_destinataire)
        toolbar.addWidget(btn_hist_dest)

        self.btn_affecter = QPushButton("Affectation")
        self.btn_affecter.setStyleSheet("""
            QPushButton { background-color: #D32F2FA4; color: white; padding: 8px 16px; border-radius: 10px; }
            QPushButton:hover { background-color: #E91E63; }
        """)
        self.btn_affecter.clicked.connect(self.open_affectation)
        self.btn_affecter.setEnabled(False)
        toolbar.addWidget(self.btn_affecter)

        # === RECHERCHE + FILTRE ===
        self.search = QLineEdit()
        self.search.setPlaceholderText("Rechercher par article...")
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

        search_bar = QHBoxLayout()
        search_bar.addWidget(QLabel("Recherche :"))
        search_bar.addWidget(self.search)
        search_bar.addSpacing(20)
        search_bar.addWidget(QLabel("Nature :"))
        search_bar.addWidget(self.filter_nature)
        search_bar.addStretch()

        # === TABLEAU ===
        self.table = QTableWidget()
        self.table.setColumnCount(7)
        self.table.setHorizontalHeaderLabels([
            "Article", "Nature", "Quantité", "Prix", "Seuil mini", "Date ajout", "Observation"
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
        self.prix = QDoubleSpinBox(); self.prix.setRange(0, 10_000_000); self.prix.setDecimals(2)
        self.seuil = QSpinBox(); self.seuil.setRange(0, 10_000_000)
        self.date = QLineEdit(datetime.now().strftime("%Y-%m-%d"))
        self.date.setMaximumWidth(150)
        self.observation = QLineEdit()

        form = QFormLayout()
        form.addRow("Article *", self.nom)
        form.addRow("Nature", self.nature)
        form.addRow("Observation", self.observation)

        hnums = QHBoxLayout()
        hnums.addWidget(self._labeled("Quantité", self.quantite))
        hnums.addWidget(self._labeled("Prix", self.prix))
        hnums.addWidget(self._labeled("Seuil mini", self.seuil))
        hnums.addWidget(self._labeled("Date ajout", self.date))

        # === BOUTONS CRUD + EXPORT ===
        self.btn_add = QPushButton("➕ Ajouter")
        self.btn_add.clicked.connect(self.add_product)

        self.btn_update = QPushButton("✏️ Modifier")
        self.btn_update.clicked.connect(self.update_product)
        self.btn_update.setEnabled(False)

        self.btn_delete = QPushButton("🗑️ Supprimer")
        self.btn_delete.clicked.connect(self.delete_product)
        self.btn_delete.setEnabled(False)

        self.btn_clear = QPushButton("🔄 Nouveau / Vider")
        self.btn_clear.clicked.connect(self.clear_form)

        self.btn_export_xlsx = QPushButton("Exporter Excel")
        self.btn_export_xlsx.clicked.connect(self.on_export_excel)

        self.btn_export_pdf = QPushButton("Exporter PDF")
        self.btn_export_pdf.clicked.connect(self.on_export_pdf)

        self.badge_low = QLabel("")

        crud = QHBoxLayout()
        crud.addWidget(self.btn_add)
        crud.addWidget(self.btn_update)
        crud.addWidget(self.btn_delete)
        crud.addWidget(self.btn_clear)
        crud.addStretch()
        crud.addWidget(self.btn_export_xlsx)
        crud.addWidget(self.btn_export_pdf)

        # === ASSEMBLAGE FINAL ===
        top = QHBoxLayout()
        lbl = QLabel("Inventaire")
        lbl.setProperty("class", "header")
        top.addWidget(lbl)
        top.addStretch()
        top.addWidget(self.badge_low)

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
        key = self.search.text().strip()
        nature_filter = self.filter_nature.currentData()

        conn = get_conn()
        c = conn.cursor()
        query = """SELECT nom, nature, quantite, prix, seuil_min, date_ajout, observation, id 
                   FROM produits WHERE 1=1"""
        params = []
        if key:
            query += " AND nom LIKE ?"
            params.append(f"%{key}%")
        if nature_filter:
            query += " AND nature = ?"
            params.append(nature_filter)
        query += " ORDER BY nom ASC"
        c.execute(query, params)
        rows = c.fetchall()
        conn.close()

        self.table.setRowCount(len(rows))
        for i, r in enumerate(rows):
            for j in range(7):
                item = QTableWidgetItem(str(r[j]) if r[j] is not None else "")
                if j in (2, 3, 4):
                    item.setTextAlignment(Qt.AlignCenter)
                if r[2] is not None and r[4] is not None and r[2] < r[4]:
                    item.setForeground(QColor("#D32F2F"))
                self.table.setItem(i, j, item)
            self.table.item(i, 0).setData(Qt.UserRole, r[7])  # ID caché

        self.table.resizeRowsToContents()
        self.update_badge()

    def on_row_click(self, index):
        row = index.row()
        self.selected_id = self.table.item(row, 0).data(Qt.UserRole)
        self.nom.setText(self.table.item(row, 0).text())
        self.nature.setCurrentText(self.table.item(row, 1).text())
        self.quantite.setValue(int(float(self.table.item(row, 2).text() or 0)))
        self.prix.setValue(float(self.table.item(row, 3).text() or 0))
        self.seuil.setValue(int(float(self.table.item(row, 4).text() or 0)))
        self.date.setText(self.table.item(row, 5).text())
        self.observation.setText(self.table.item(row, 6).text())

        self.btn_update.setEnabled(True)
        self.btn_delete.setEnabled(True)
        self.btn_affecter.setEnabled(True)
        self.btn_add.setEnabled(False)  

    def clear_form(self):
        self.selected_id = None
        self.nom.clear()
        self.nature.setCurrentIndex(0)
        self.quantite.setValue(0)
        self.prix.setValue(0.0)
        self.seuil.setValue(0)
        self.date.setText(datetime.now().strftime("%Y-%m-%d"))
        self.observation.clear()
        self.btn_update.setEnabled(False)
        self.btn_delete.setEnabled(False)
        self.btn_affecter.setEnabled(False)
        self.btn_add.setEnabled(True)           # ← RÉACTIVER AJOUTER
        self.table.clearSelection()

    def add_product(self):
        if not self.nom.text().strip():
            QMessageBox.warning(self, "Erreur", "L'article est obligatoire.")
            return
        conn = get_conn()
        c = conn.cursor()
        try:
            c.execute("""
                INSERT INTO produits (nom, nature, quantite, prix, seuil_min, date_ajout, observation)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            """, (self.nom.text().strip(), self.nature.currentText(), self.quantite.value(),
                  self.prix.value(), self.seuil.value(), self.date.text(), self.observation.text().strip()))
            conn.commit()
            QMessageBox.information(self, "Succès", "Article ajouté.")
            self.clear_form()
            self.load_table()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            conn.close()

    def update_product(self):
        if not self.selected_id:
            QMessageBox.warning(self, "Erreur", "Sélectionnez un article.")
            return
        conn = get_conn()
        c = conn.cursor()
        try:
            c.execute("""
                UPDATE produits SET nom=?, nature=?, quantite=?, prix=?, seuil_min=?, date_ajout=?, observation=?
                WHERE id=?
            """, (self.nom.text().strip(), self.nature.currentText(), self.quantite.value(),
                  self.prix.value(), self.seuil.value(), self.date.text(), self.observation.text().strip(),
                  self.selected_id))
            conn.commit()
            QMessageBox.information(self, "Succès", "Article modifié.")
            self.clear_form()
            self.load_table()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            conn.close()

    def delete_product(self):
        if not self.selected_id:
            return
        if QMessageBox.question(self, "Confirmer", "Supprimer cet article ?") != QMessageBox.Yes:
            return
        conn = get_conn()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM produits WHERE id=?", (self.selected_id,))
            conn.commit()
            QMessageBox.information(self, "Succès", "Article supprimé.")
            self.clear_form()
            self.load_table()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            conn.close()

    def update_badge(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM produits WHERE quantite < seuil_min AND seuil_min > 0")
        n = c.fetchone()[0]
        conn.close()
        self.badge_low.setText(f"{n} article(s) en alerte" if n else "")

    def on_export_excel(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT nom, nature, quantite, prix, seuil_min, date_ajout, observation FROM produits ORDER BY nom")
        rows = c.fetchall()
        conn.close()
        if not rows:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exporter Excel", "", "Excel (*.xlsx)")
        if path:
            export_excel(rows, path)

    def on_export_pdf(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT nom, nature, quantite, prix, seuil_min, date_ajout, observation FROM produits ORDER BY nom")
        rows = c.fetchall()
        conn.close()
        if not rows:
            return
        path, _ = QFileDialog.getSaveFileName(self, "Exporter PDF", "", "PDF (*.pdf)")
        if path:
            export_pdf(rows, path)

    def open_affectation(self):
        # Remplacer la méthode open_affectation par cette version améliorée :

    
        if not self.selected_id:
            QMessageBox.warning(self, "Sélection requise", "Veuillez sélectionner un article.")
            return

        row = self.table.currentRow()
        article = self.table.item(row, 0).text()
        nature = self.table.item(row, 1).text()
        try:
            stock_actuel = int(float(self.table.item(row, 2).text() or 0))
        except:
            stock_actuel = 0

        if stock_actuel <= 0:
            QMessageBox.warning(self, "Stock insuffisant", "Stock disponible insuffisant.")
            return

        dialog = QDialog(self)
        dialog.setWindowTitle("Affectation de matériel")
        dialog.setFixedSize(500, 400)
        dialog.setStyleSheet("""
            QDialog {
                background-color: #f5f5f5;
            }
            QLabel {
                font-size: 13px;
                color: #333;
            }
            QLineEdit, QComboBox, QSpinBox {
                padding: 8px;
                border: 2px solid #ddd;
                border-radius: 6px;
                background: white;
                font-size: 13px;
                min-height: 20px;
            }
            QLineEdit:focus, QComboBox:focus, QSpinBox:focus {
                border-color: #1976D2;
            }
            QPushButton {
                padding: 10px 20px;
                border-radius: 6px;
                font-size: 13px;
                font-weight: bold;
                min-width: 100px;
            }
        """)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(25, 25, 25, 25)
        layout.setSpacing(15)

        # En-tête avec informations produit
        header_frame = QFrame()
        header_frame.setStyleSheet("""
            QFrame {
                background: #1976D2;
                border-radius: 10px;
                padding: 15px;
            }
            QLabel {
                color: white;
            }
        """)
        header_layout = QVBoxLayout(header_frame)

        lbl_article = QLabel(article)
        lbl_article.setStyleSheet("font-size: 18px; font-weight: bold;")
        lbl_article.setWordWrap(True)
        header_layout.addWidget(lbl_article)

        lbl_info = QLabel(f"Catégorie : {nature} | Stock disponible : {stock_actuel}")
        lbl_info.setStyleSheet("font-size: 13px; opacity: 0.9;")
        header_layout.addWidget(lbl_info)

        layout.addWidget(header_frame)

        # Formulaire
        form_frame = QFrame()
        form_frame.setStyleSheet("""
            QFrame {
                background: white;
                border-radius: 10px;
                padding: 20px;
            }
        """)
        form_layout = QFormLayout(form_frame)
        form_layout.setSpacing(15)
        form_layout.setLabelAlignment(Qt.AlignRight)

        # Destinataire
        lbl_dest = QLabel("Destinataire :")
        lbl_dest.setStyleSheet("font-weight: bold; color: #555;")
        combo_dest = QComboBox()
        combo_dest.addItems(DESTINATAIRES)
        combo_dest.setEditable(True)
        combo_dest.setInsertPolicy(QComboBox.InsertAtTop)
        form_layout.addRow(lbl_dest, combo_dest)

        # Quantité
        lbl_qte = QLabel("Quantité à affecter :")
        lbl_qte.setStyleSheet("font-weight: bold; color: #555;")
        spin_qte = QSpinBox()
        spin_qte.setRange(1, max(1, stock_actuel))
        spin_qte.setValue(1)
        spin_qte.setSuffix(f" / {stock_actuel} disponible(s)")
        spin_qte.setStyleSheet("""
            QSpinBox::up-button, QSpinBox::down-button {
                width: 20px;
                height: 20px;
            }
        """)
        form_layout.addRow(lbl_qte, spin_qte)

        # Observation
        lbl_obs = QLabel("Observation :")
        lbl_obs.setStyleSheet("font-weight: bold; color: #555;")
        edit_obs = QLineEdit()
        edit_obs.setPlaceholderText("N° bon de sortie, remarque...")
        form_layout.addRow(lbl_obs, edit_obs)

        layout.addWidget(form_frame)

        # Récapitulatif
        recap_frame = QFrame()
        recap_frame.setStyleSheet("""
            QFrame {
                background: #E3F2FD;
                border-radius: 8px;
                padding: 12px;
            }
            QLabel {
                color: #0D47A1;
            }
        """)
        recap_layout = QHBoxLayout(recap_frame)
        
        self.lbl_recap = QLabel(f"Stock après affectation : {stock_actuel} - 1 = {stock_actuel - 1}")
        self.lbl_recap.setStyleSheet("font-size: 14px; font-weight: bold;")
        recap_layout.addWidget(self.lbl_recap)
        
        # Mise à jour du récapitulatif quand la quantité change
        def update_recap():
            qte = spin_qte.value()
            reste = stock_actuel - qte
            self.lbl_recap.setText(f"Stock après affectation : {stock_actuel} - {qte} = {reste}")
            if reste < 0:
                self.lbl_recap.setStyleSheet("font-size: 14px; font-weight: bold; color: #D32F2F;")
            else:
                self.lbl_recap.setStyleSheet("font-size: 14px; font-weight: bold; color: #0D47A1;")
        
        spin_qte.valueChanged.connect(update_recap)
        
        layout.addWidget(recap_frame)

        layout.addStretch()

        # Boutons
        btn_layout = QHBoxLayout()
        btn_layout.addStretch()

        btn_annuler = QPushButton("Annuler")
        btn_annuler.setStyleSheet("""
            QPushButton {
                background-color: #9e9e9e;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #757575;
            }
        """)

        btn_valider = QPushButton("Valider l'affectation")
        btn_valider.setStyleSheet("""
            QPushButton {
                background-color: #1976D2;
                color: white;
                border: none;
            }
            QPushButton:hover {
                background-color: #1565C0;
            }
            QPushButton:pressed {
                background-color: #0D47A1;
            }
        """)

        btn_annuler.clicked.connect(dialog.reject)
        btn_valider.clicked.connect(lambda: self.valider_affectation(
            dialog, self.selected_id, spin_qte.value(), combo_dest.currentText(), edit_obs.text()
        ))

        btn_layout.addWidget(btn_annuler)
        btn_layout.addWidget(btn_valider)
        layout.addLayout(btn_layout)

        dialog.exec_()

    def valider_affectation(self, dialog, produit_id, quantite, destinataire, observation):
        if quantite <= 0:
            QMessageBox.warning(self, "Erreur", "Quantité invalide.")
            return

        conn = get_conn()
        c = conn.cursor()
        try:
            c.execute("SELECT quantite FROM produits WHERE id = ?", (produit_id,))
            stock = c.fetchone()[0]

            if stock < quantite:
                QMessageBox.warning(self, "Stock insuffisant", f"Stock disponible : {stock}")
                return

            new_stock = stock - quantite
            c.execute("UPDATE produits SET quantite = ? WHERE id = ?", (new_stock, produit_id))

            date_mvt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            c.execute("""
                INSERT INTO mouvements (produit_id, type, quantite, date_mvt, service, observation)
                VALUES (?, 'SORTIE', ?, ?, ?, ?)
            """, (produit_id, quantite, date_mvt, destinataire, observation.strip()))

            conn.commit()
            QMessageBox.information(self, "Succès", f"Affectation enregistrée.\nStock restant : {new_stock}")
            self.load_table()
            self.update_badge()
            dialog.accept()

        except Exception as e:
            conn.rollback()
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            conn.close()

    def ouvrir_historique_article(self):
        if not self.selected_id:
            QMessageBox.warning(self, "Sélection requise", "Sélectionnez un article.")
            return

        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT nom FROM produits WHERE id=?", (self.selected_id,))
        nom = c.fetchone()[0]
        conn.close()

        self._ouvrir_fenetre_historique(
            title=f"Historique – {nom}",
            prefiltre_article=nom      # on pré-remplit avec le nom, sans bloquer l'ID
        )

    def ouvrir_historique_par_destinataire(self):
        dest, ok = QInputDialog.getItem(
            self, "Filtrer par destinataire", "Bureau / CB :", ["Tous"] + DESTINATAIRES, 0, False
        )
        if not ok:
            return

        title = "Historique complet" if dest == "Tous" else f"Historique – {dest}"
        filtre = None if dest == "Tous" else dest

        self._ouvrir_fenetre_historique(title=title, filtre_destinataire=filtre)

    def _ouvrir_fenetre_historique(self, title="Historique des Mouvements",
                               prefiltre_article=None,
                               filtre_destinataire=None):
        dialog = QDialog(self)
        dialog.setWindowTitle(title)
        dialog.resize(1100, 680)

        layout = QVBoxLayout(dialog)

        # --- Barre de filtres ---
        filtres = QHBoxLayout()

        # COMBOBOX pour les articles (avec liste déroulante)
        combo_article = QComboBox()
        combo_article.setEditable(True)                     # permet aussi de taper
        combo_article.setPlaceholderText("Choisir un article...")
        combo_article.addItem("Tous les articles", None)    # option pour tout afficher

        # Remplir la combo avec les noms d'articles existants
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT DISTINCT nom FROM produits ORDER BY nom")
        for (nom,) in c.fetchall():
            combo_article.addItem(nom, nom)
        conn.close()

        if prefiltre_article:
            idx = combo_article.findText(prefiltre_article)
            if idx >= 0:
                combo_article.setCurrentIndex(idx)
            else:
                combo_article.setCurrentIndex(0)   # "Tous les articles"
        else:
            combo_article.setCurrentIndex(0)

        # COMBOBOX pour les destinataires
        combo_dest = QComboBox()
        combo_dest.addItem("Tous", None)            # None = pas de filtre
        for d in DESTINATAIRES:
            combo_dest.addItem(d, d)
        if filtre_destinataire:
            idx = combo_dest.findData(filtre_destinataire)
            if idx >= 0:
                combo_dest.setCurrentIndex(idx)

        is_hist_par_dest = filtre_destinataire is not None

        filtres.addWidget(QLabel("Article :"))
        filtres.addWidget(combo_article)
        filtres.addSpacing(20)
        filtres.addWidget(QLabel("Destinataire :"))
        filtres.addWidget(combo_dest)
        filtres.addSpacing(20)

        if not is_hist_par_dest:
            combo_type = QComboBox()
            combo_type.addItems(["Tous", "ENTREE", "SORTIE"])
            filtres.addWidget(QLabel("Type :"))
            filtres.addWidget(combo_type)

        filtres.addStretch()
        layout.addLayout(filtres)

        # --- Tableau ---
        if is_hist_par_dest:
            table = QTableWidget()
            table.setColumnCount(6)
            table.setHorizontalHeaderLabels([
                "Date", "Article", "Quantité", "Destinataire", "Stock après", "Observation"
            ])
        else:
            table = QTableWidget()
            table.setColumnCount(7)
            table.setHorizontalHeaderLabels([
                "Date", "Article", "Type", "Quantité", "Destinataire", "Stock après", "Observation"
            ])
        table.horizontalHeader().setSectionResizeMode(QHeaderView.Stretch)
        layout.addWidget(table)

        # --- Boutons d'export ---
        bottom = QHBoxLayout()
        bottom.addStretch()
        btn_excel = QPushButton("Exporter Excel")
        btn_pdf = QPushButton("Exporter PDF")
        bottom.addWidget(btn_excel)
        bottom.addWidget(btn_pdf)
        layout.addLayout(bottom)

        # --- Fonction de chargement des données ---
        def charger():
            # Valeurs actuelles des filtres
            article_data = combo_article.currentData()   # None si "Tous"
            dest_data = combo_dest.currentData()         # None si "Tous"

            conn = get_conn()
            c = conn.cursor()

            query = """
                SELECT 
                    m.date_mvt,
                    p.nom,
                    m.type,
                    m.quantite,
                    m.service,
                    m.observation,
                    p.quantite AS stock_apres
                FROM mouvements m
                JOIN produits p ON m.produit_id = p.id
                WHERE 1=1
            """
            params = []

            # Filtre article (exact, car on a une liste de noms)
            if article_data is not None:
                query += " AND p.nom = ?"
                params.append(article_data)

            # Filtre destinataire (uniquement basé sur la combo)
            if dest_data is not None:
                query += " AND m.service = ?"
                params.append(dest_data)

            # Filtre type (sauf pour historique destinataire)
            if not is_hist_par_dest:
                type_val = combo_type.currentText()
                if type_val != "Tous":
                    query += " AND m.type = ?"
                    params.append(type_val)
            else:
                # Pour l'historique par destinataire, on ne montre que les sorties
                query += " AND m.type = 'SORTIE'"

            query += " ORDER BY m.date_mvt DESC LIMIT 1500"

            c.execute(query, params)
            rows = c.fetchall()
            conn.close()

            # Remplissage du tableau
            if is_hist_par_dest:
                table.setRowCount(len(rows))
                for i, row in enumerate(rows):
                    items = [
                        QTableWidgetItem(row[0]),               # Date
                        QTableWidgetItem(row[1]),               # Article
                        QTableWidgetItem(str(row[3])),          # Quantité
                        QTableWidgetItem(row[4] or ""),         # Destinataire
                        QTableWidgetItem(str(row[6])),          # Stock après
                        QTableWidgetItem(row[5] or "")          # Observation
                    ]
                    for j, item in enumerate(items):
                        if j in (2, 4):
                            item.setTextAlignment(Qt.AlignCenter)
                    for j in range(6):
                        table.setItem(i, j, items[j])
            else:
                table.setRowCount(len(rows))
                for i, row in enumerate(rows):
                    items = [
                        QTableWidgetItem(row[0]),               # Date
                        QTableWidgetItem(row[1]),               # Article
                        QTableWidgetItem(row[2]),               # Type
                        QTableWidgetItem(str(row[3])),          # Quantité
                        QTableWidgetItem(row[4] or ""),         # Destinataire
                        QTableWidgetItem(str(row[6])),          # Stock après
                        QTableWidgetItem(row[5] or "")          # Observation
                    ]
                    if row[2] == "SORTIE":
                        items[2].setForeground(QColor("#D32F2F"))
                    elif row[2] == "ENTREE":
                        items[2].setForeground(QColor("#2E7D32"))
                    for j, item in enumerate(items):
                        if j in (3, 5):
                            item.setTextAlignment(Qt.AlignCenter)
                    for j in range(7):
                        table.setItem(i, j, items[j])

            table.resizeRowsToContents()
            return rows

        # Connexion des signaux
        combo_article.currentIndexChanged.connect(charger)
        combo_dest.currentIndexChanged.connect(charger)
        if not is_hist_par_dest:
            combo_type.currentIndexChanged.connect(charger)

        # Gestion des exports
        def export_excel_action():
            rows = charger()
            if not rows:
                QMessageBox.warning(dialog, "Export", "Aucune donnée à exporter.")
                return
            path, _ = QFileDialog.getSaveFileName(dialog, "Exporter Excel", "", "Excel (*.xlsx)")
            if path:
                export_history_excel(rows, path)

        def export_pdf_action():
            rows = charger()
            if not rows:
                QMessageBox.warning(dialog, "Export", "Aucune donnée à exporter.")
                return
            path, _ = QFileDialog.getSaveFileName(dialog, "Exporter PDF", "", "PDF (*.pdf)")
            if path:
                export_history_pdf(rows, path)

        btn_excel.clicked.connect(export_excel_action)
        btn_pdf.clicked.connect(export_pdf_action)

        # Chargement initial
        charger()

        dialog.exec_()
    def _export_hist(self, dialog, mode, article_txt, dest_val, type_val, filtre_article, filtre_destinataire):
        conn = get_conn()
        c = conn.cursor()
        query = """
            SELECT m.date_mvt, m.type, p.nom, m.quantite, m.service, m.observation, p.quantite
            FROM mouvements m JOIN produits p ON m.produit_id = p.id WHERE 1=1
        """
        params = []
        if filtre_article:
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
        query += " ORDER BY m.date_mvt DESC"
        c.execute(query, params)
        rows = c.fetchall()
        conn.close()

        if not rows:
            QMessageBox.warning(dialog, "Export", "Aucun données à exporter.")
            return

        if mode == 'excel':
            path, _ = QFileDialog.getSaveFileName(dialog, "Exporter Excel", "", "Excel (*.xlsx)")
            if path:
                export_history_excel(rows, path)
        else:
            path, _ = QFileDialog.getSaveFileName(dialog, "Exporter PDF", "", "PDF (*.pdf)")
            if path:
                export_history_pdf(rows, path)

# Pour tester rapidement (optionnel)
if __name__ == "__main__":
    import sys
    app = QApplication(sys.argv)
    window = MagasinApp()
    window.show()
    sys.exit(app.exec_())

