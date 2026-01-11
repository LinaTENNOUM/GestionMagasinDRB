import os, sys, sqlite3

def app_dir():
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

DB_PATH = os.path.join(app_dir(), "magasin.db")

def get_conn():
    return sqlite3.connect(DB_PATH)

def init_db():
    conn = get_conn()
    c = conn.cursor()
    c.execute("""
        CREATE TABLE IF NOT EXISTS produits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nom TEXT NOT NULL,
            nature TEXT,
            quantite INTEGER NOT NULL DEFAULT 0,
            prix REAL NOT NULL DEFAULT 0,
            seuil_min INTEGER NOT NULL DEFAULT 0,
            date_ajout TEXT NOT NULL,
            observation TEXT
        )
    """)
    c.execute("""
        CREATE TABLE IF NOT EXISTS mouvements (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            produit_id INTEGER NOT NULL,
            type TEXT NOT NULL,            -- ENTREE / SORTIE
            quantite INTEGER NOT NULL,
            date_mvt TEXT NOT NULL,
            service TEXT,
            observation TEXT,
            FOREIGN KEY (produit_id) REFERENCES produits(id)
        )
    """)


    conn.commit()
    conn.close()

