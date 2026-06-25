"""
SeePlaces — Run history (SQLite)
Stores a summary record every time a guide report is generated.
"""
import sqlite3, json, os
from datetime import datetime

DB_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), "history.db")


def _init():
    conn = sqlite3.connect(DB_PATH)
    conn.execute("""
        CREATE TABLE IF NOT EXISTS runs (
            id          INTEGER PRIMARY KEY AUTOINCREMENT,
            created_at  TEXT NOT NULL,
            date_from   TEXT,
            date_to     TEXT,
            week_label  TEXT,
            summary     TEXT NOT NULL
        )
    """)
    conn.commit()
    conn.close()


def save_run(summary: dict, date_from: str = None, date_to: str = None):
    _init()
    now        = datetime.now()
    week_label = now.strftime("W%W %Y")
    conn       = sqlite3.connect(DB_PATH)
    conn.execute(
        "INSERT INTO runs (created_at, date_from, date_to, week_label, summary) VALUES (?,?,?,?,?)",
        (now.isoformat(), date_from or "", date_to or "", week_label, json.dumps(summary)),
    )
    conn.commit()
    conn.close()


def get_history(limit: int = 16) -> list:
    _init()
    conn = sqlite3.connect(DB_PATH)
    rows = conn.execute(
        "SELECT id, created_at, date_from, date_to, week_label, summary "
        "FROM runs ORDER BY id DESC LIMIT ?",
        (limit,),
    ).fetchall()
    conn.close()
    result = []
    for r in rows:
        entry = json.loads(r[5])
        entry.update({"id": r[0], "created_at": r[1],
                      "date_from": r[2], "date_to": r[3], "week_label": r[4]})
        result.append(entry)
    return list(reversed(result))  # chronological order
