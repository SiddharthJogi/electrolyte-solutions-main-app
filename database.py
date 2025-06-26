import sqlite3
from datetime import datetime
DB_PATH = "conversion_logs.db"

def setup_database():
    conn = sqlite3.connect(DB_PATH)
    conn.execute('''CREATE TABLE IF NOT EXISTS logs (
        id INTEGER PRIMARY KEY AUTOINCREMENT,
        filename TEXT,
        converted_at TEXT,
        status TEXT,
        rows INTEGER
    );''')
    conn.commit()
    conn.close()

def log_conversion(filename, status, rows):
    conn = sqlite3.connect(DB_PATH)
    conn.execute("INSERT INTO logs (filename, converted_at, status, rows) VALUES (?, ?, ?, ?)",
                 (filename, datetime.now().isoformat(), status, rows))
    conn.commit()
    conn.close() 