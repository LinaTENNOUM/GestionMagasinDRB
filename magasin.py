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

class MagasinApp(QWidget):
    def __init__(self):
        super().__init__()
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

        # === BOUTONS ===
        self.btn_add = QPushButton("Ajouter"); self.btn_add.clicked.connect(self.add_product)
        self.btn_update = QPushButton("Modifier"); self.btn_update.clicked.connect(self.update_product)
        self.btn_delete = QPushButton("Supprimer"); self.btn_delete.clicked.connect(self.delete_product)
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
        
        self.btn_hist_article = QPushButton("Hist. Article")
        self.btn_hist_article.clicked.connect(self.ouvrir_historique_article)
        
        self.btn_hist_dest = QPushButton("Hist. par Bureau/CB")
        self.btn_hist_dest.clicked.connect(self.ouvrir_historique_par_destinataire)

        crud.addWidget(self.btn_hist_article)
        crud.addWidget(self.btn_hist_dest)


        self.btn_affecter.clicked.connect(self.open_affectation)
        crud.addWidget(self.btn_affecter)

        self.btn_historique = QPushButton("Historique Mouvements")
        self.btn_historique.clicked.connect(self.ouvrir_historique_article)
        crud.addWidget(self.btn_historique)       

        main = QVBoxLayout()
        main.addLayout(top)
        main.addLayout(search_bar)
        main.addWidget(self.table)
        main.addLayout(form)
        main.addLayout(hnums)
        main.addLayout(crud)
        main.setContentsMargins(15, 15, 15, 15)
        main.setSpacing(12)
        self.setLayout(main)

        self.current_id = None
        self.hidden_ids = []
        self.load_table()
        self.update_low_stock_badge()

    def _labeled(self, text, widget):
        box = QVBoxLayout()
        lab = QLabel(text)
        box.addWidget(lab)
        box.addWidget(widget)
        w = QWidget()
        w.setLayout(box)
        return w

    # === FONCTIONS ===
    def load_table(self):
        key = self.search.text().strip()
        nature_filter = self.filter_nature.currentData()
        conn = get_conn()
        c = conn.cursor()
        query = """ SELECT nom, nature, quantite, prix, seuil_min, date_ajout, observation, id FROM produits WHERE 1=1 """
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
        self.hidden_ids = []
        for i, r in enumerate(rows):
            self.hidden_ids.append(r[-1])
            for j, v in enumerate(r[:-1]):
                item = QTableWidgetItem(str(v) if v is not None else "")
                if j in (2, 3, 4):
                    item.setTextAlignment(Qt.AlignCenter)
                if r[2] is not None and r[4] is not None and r[2] < r[4]:
                    item.setBackground(QColor("#FFEBEE"))
                    item.setForeground(QColor("#D32F2F"))
                self.table.setItem(i, j, item)
        self.table.resizeRowsToContents()

    def on_row_click(self):
        row = self.table.currentRow()
        if row < 0: return
        self.current_id = self.hidden_ids[row]
        self.nom.setText(self.table.item(row, 0).text())
        self.nature.setCurrentText(self.table.item(row, 1).text())
        self.quantite.setValue(int(float(self.table.item(row, 2).text() or 0)))
        self.prix.setValue(float(self.table.item(row, 3).text() or 0))
        self.seuil.setValue(int(float(self.table.item(row, 4).text() or 0)))
        self.date.setText(self.table.item(row, 5).text())
        self.observation.setText(self.table.item(row, 6).text())

    def validate_inputs(self):
        if not self.nom.text().strip():
            QMessageBox.warning(self, "Champs manquants", "La désignation est obligatoire.")
            return False
        try:
            datetime.strptime(self.date.text().strip(), "%Y-%m-%d")
        except:
            QMessageBox.warning(self, "Date invalide", "Format YYYY-MM-DD requis.")
            return False
        return True

    def add_product(self):
        if not self.validate_inputs(): return
        conn = get_conn()
        c = conn.cursor()
        try:
            c.execute(""" INSERT INTO produits (nom, nature, quantite, prix, seuil_min, date_ajout, observation)
                          VALUES (?, ?, ?, ?, ?, ?, ?) """,
                      (self.nom.text().strip(), self.nature.currentText(), self.quantite.value(),
                       self.prix.value(), self.seuil.value(), self.date.text().strip(),
                       self.observation.text().strip()))
            conn.commit()
            QMessageBox.information(self, "Succès", "Article ajouté.")
            self.clear_form()
            self.load_table()
            self.update_low_stock_badge()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            conn.close()

    def update_product(self):
        if self.current_id is None:
            QMessageBox.warning(self, "Sélection", "Sélectionnez un article dans la liste.")
            return
        if not self.validate_inputs(): return
        conn = get_conn()
        c = conn.cursor()
        try:
            c.execute(""" UPDATE produits SET nom=?, nature=?, quantite=?, prix=?, seuil_min=?, date_ajout=?, observation=?
                          WHERE id=? """,
                      (self.nom.text().strip(), self.nature.currentText(), self.quantite.value(),
                       self.prix.value(), self.seuil.value(), self.date.text().strip(),
                       self.observation.text().strip(), self.current_id))
            conn.commit()
            QMessageBox.information(self, "Succès", "Article modifié.")
            self.clear_form()
            self.load_table()
            self.update_low_stock_badge()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            conn.close()

    def delete_product(self):
        if self.current_id is None:
            QMessageBox.warning(self, "Sélection", "Sélectionnez un article dans la liste.")
            return
        if QMessageBox.question(self, "Confirmation", "Supprimer cet article ?", QMessageBox.Yes | QMessageBox.No) == QMessageBox.No:
            return
        conn = get_conn()
        c = conn.cursor()
        try:
            c.execute("DELETE FROM produits WHERE id=?", (self.current_id,))
            conn.commit()
            QMessageBox.information(self, "Succès", "Article supprimé.")
            self.clear_form()
            self.load_table()
            self.update_low_stock_badge()
        except Exception as e:
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            conn.close()

    def clear_form(self):
        self.current_id = None
        self.nom.clear()
        self.nature.setCurrentIndex(0)
        self.quantite.setValue(0)
        self.prix.setValue(0.0)
        self.seuil.setValue(0)
        self.date.setText(datetime.now().strftime("%Y-%m-%d"))
        self.observation.clear()
        self.table.clearSelection()

    def update_low_stock_badge(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT COUNT(*) FROM produits WHERE quantite < seuil_min AND seuil_min > 0")
        n = c.fetchone()[0]
        conn.close()
        self.badge_low.setText(f"{n} article(s) sous seuil" if n else "")

    def fetch_all_rows(self):
        conn = get_conn()
        c = conn.cursor()
        c.execute("SELECT nom, nature, quantite, prix, seuil_min, date_ajout, observation FROM produits ORDER BY nom ASC")
        rows = c.fetchall()
        conn.close()
        return rows

    def on_export_excel(self):
        rows = self.fetch_all_rows()
        if not rows:
            QMessageBox.information(self, "Export", "Aucun article à exporter.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Enregistrer Excel", f"inventaire_{datetime.now().strftime('%Y%m%d')}.xlsx", "Fichiers Excel (*.xlsx)")
        if not path: return
        try:
            export_excel(rows, path)
            QMessageBox.information(self, "Export", f"Excel exporté :\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur export", str(e))

    def on_export_pdf(self):
        rows = self.fetch_all_rows()
        if not rows:
            QMessageBox.information(self, "Export", "Aucun article à exporter.")
            return
        path, _ = QFileDialog.getSaveFileName(self, "Enregistrer PDF", f"inventaire_{datetime.now().strftime('%Y%m%d')}.pdf", "PDF (*.pdf)")
        if not path: return
        try:
            export_pdf(rows, path)
            QMessageBox.information(self, "Export", f"PDF exporté :\n{path}")
        except Exception as e:
            QMessageBox.critical(self, "Erreur export", str(e))
    def open_mouvements(self):
        if self.current_id is None:
            QMessageBox.warning(self, "Sélection", "Sélectionnez un article.")
            return
    
        row = self.table.currentRow()
        stock = int(self.table.item(row, 2).text())
    
        self.mvt_win = MouvementWindow(
            self.current_id,
            self.nom.text(),
            stock
        )
        self.mvt_win.show()
    
    def open_affectation(self):
        if self.current_id is None:
            QMessageBox.warning(self, "Sélection requise", "Veuillez sélectionner un article dans la liste.")
            return

        row = self.table.currentRow()
        if row < 0:
            return

        designation = self.table.item(row, 0).text()
        nature = self.table.item(row, 1).text()
        stock_actuel = self.table.item(row, 2).text()
        try:
            stock_actuel = int(float(stock_actuel))
        except:
            stock_actuel = 0

        if stock_actuel <= 0:
            QMessageBox.warning(self, "Stock insuffisant", "Cet article n'a plus de stock disponible.")
            return

        # Petite fenêtre de dialogue simple
        dialog = QDialog(self)
        dialog.setWindowTitle("Affectation / Sortie")
        dialog.setFixedSize(480, 320)
        dialog.setWindowFlags(dialog.windowFlags() & ~Qt.WindowContextHelpButtonHint)

        layout = QVBoxLayout(dialog)
        layout.setContentsMargins(20, 20, 20, 20)
        layout.setSpacing(16)

        # Infos article
        lbl_info = QLabel(f"<b>{designation}</b><br><small>{nature} • Stock actuel : {stock_actuel}</small>")
        lbl_info.setStyleSheet("font-size:14px; color:#1A237E;")
        layout.addWidget(lbl_info)



        # Choix destinataire
        combo_dest = QComboBox()
        combo_dest.addItems(DESTINATAIRES)
        layout.addWidget(QLabel("Destinataire :"))
        layout.addWidget(combo_dest)

        # Quantité
        spin_qte = QSpinBox()
        spin_qte.setRange(1, max(1, stock_actuel))
        spin_qte.setValue(1)
        layout.addWidget(QLabel("Quantité à affecter :"))
        layout.addWidget(spin_qte)

        # Observation
        edit_obs = QLineEdit()
        edit_obs.setPlaceholderText("Observation / N° bon / référence... (facultatif)")
        layout.addWidget(QLabel("Observation :"))
        layout.addWidget(edit_obs)

        layout.addStretch()

        # Boutons
        btn_layout = QHBoxLayout()
        btn_annuler = QPushButton("Annuler")
        btn_valider = QPushButton("Valider l'affectation")
        btn_valider.setStyleSheet("background:#D81B60; color:white;")

        btn_annuler.clicked.connect(dialog.reject)
        btn_valider.clicked.connect(lambda: self.valider_affectation(
            dialog, self.current_id, spin_qte.value(), combo_dest.currentText(), edit_obs.text()
        ))

        btn_layout.addStretch()
        btn_layout.addWidget(btn_annuler)
        btn_layout.addWidget(btn_valider)
        layout.addLayout(btn_layout)

        dialog.exec_()
    def valider_affectation(self, dialog, produit_id, quantite, destinataire, observation):
        if quantite <= 0:
            QMessageBox.warning(self, "Erreur", "La quantité doit être supérieure à 0.")
            return

        conn = get_conn()
        c = conn.cursor()
        try:
            # Vérifier stock actuel
            c.execute("SELECT quantite FROM produits WHERE id = ?", (produit_id,))
            stock = c.fetchone()[0]

            if stock < quantite:
                QMessageBox.warning(self, "Stock insuffisant", f"Stock disponible : {stock}")
                return

            # Mise à jour du stock
            new_stock = stock - quantite
            c.execute("UPDATE produits SET quantite = ? WHERE id = ?", (new_stock, produit_id))

            # Enregistrement du mouvement
            date_mvt = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            c.execute("""
                INSERT INTO mouvements 
                (produit_id, type, quantite, date_mvt, service, observation)
                VALUES (?, ?, ?, ?, ?, ?)
            """, (produit_id, "SORTIE", quantite, date_mvt, destinataire, observation.strip()))

            conn.commit()

            QMessageBox.information(self, "Succès", f"Affectation enregistrée.\nStock restant : {new_stock}")
            self.load_table()               # Rafraîchir tableau
            self.update_low_stock_badge()   # MàJ badge
            dialog.accept()

        except Exception as e:
            conn.rollback()
            QMessageBox.critical(self, "Erreur", str(e))
        finally:
            conn.close()    
    def ouvrir_historique_article(self):
        """ Historique complet d'un article sélectionné """
        if self.current_id is None:
            QMessageBox.warning(self, "Sélection requise", "Veuillez sélectionner un article dans la liste.")
            return

        row = self.table.currentRow()
        nom_article = self.table.item(row, 0).text()
        id_article = self.current_id

        self._ouvrir_fenetre_historique(
            title=f"Historique - {nom_article}",
            filtre_article=id_article,
            prefiltre_article=True
        )


    def ouvrir_historique_par_destinataire(self):
        """ Historique de tous les mouvements vers un bureau / CB donné """
        from PyQt5.QtWidgets import QInputDialog

        dest, ok = QInputDialog.getItem(
            self,
            "Filtrer par destinataire",
            "Choisir le bureau ou CB :",
            [""] + DESTINATAIRES,  # "" pour "tous"
            0,
            False
        )

        if not ok or not dest:
            return

        title = f"Historique - {dest}" if dest else "Historique complet tous destinataires"

        self._ouvrir_fenetre_historique(
            title=title,
            filtre_destinataire=dest if dest else None
        )

    def _ouvrir_fenetre_historique(self, title="Historique des Mouvements", 
                                 filtre_article=None, prefiltre_article=False,
                                 filtre_destinataire=None):
        from PyQt5.QtWidgets import QDialog, QVBoxLayout, QHBoxLayout, QTableWidget, QHeaderView, QLabel, QLineEdit, QComboBox

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
        combo_dest.addItems(DESTINATAIRES)
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
    
    def charger_historique(self, table):
        article_filter = self.filter_article_hist.text().strip()
        dest_filter = self.filter_dest_hist.currentData()
        type_filter = self.filter_type_hist.currentText()

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
                p.quantite as stock_apres,
                m.id
            FROM mouvements m
            JOIN produits p ON m.produit_id = p.id
            WHERE 1=1
        """
        params = []

        if article_filter:
            query += " AND p.nom LIKE ?"
            params.append(f"%{article_filter}%")

        if dest_filter:
            query += " AND m.service = ?"
            params.append(dest_filter)

        if type_filter != "Tous":
            query += " AND m.type = ?"
            params.append(type_filter)

        query += " ORDER BY m.date_mvt DESC LIMIT 1500"

        c.execute(query, params)
        rows = c.fetchall()
        conn.close()

        table.setRowCount(len(rows))

        for i, row in enumerate(rows):
            for j, value in enumerate(row):
                item = QTableWidgetItem(str(value) if value is not None else "")
                
                # Mise en forme selon type
                if j == 1:  # Type
                    if value == "SORTIE":
                        item.setForeground(QColor("#D32F2F"))
                    elif value == "ENTREE":
                        item.setForeground(QColor("#2E7D32"))
                
                if j == 3:  # Quantité
                    item.setTextAlignment(Qt.AlignCenter)
                    if row[1] == "SORTIE":
                        item.setText(f"-{value}")
                
                if j == 6:  # Stock après
                    item.setTextAlignment(Qt.AlignRight)
                    item.setForeground(QColor("#1976D2"))

                table.setItem(i, j, item)

        table.resizeRowsToContents()