"""
38DN Excel Macro Runner — SQLite Storage
"""
import json
import sqlite3
from datetime import datetime
from pathlib import Path

from config import DB_PATH


def get_connection(db_path: Path = DB_PATH) -> sqlite3.Connection:
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute("""
        CREATE TABLE IF NOT EXISTS macro_runs (
            id              INTEGER PRIMARY KEY AUTOINCREMENT,
            workbook_name   TEXT NOT NULL,
            run_timestamp   TEXT NOT NULL,
            macro_name      TEXT NOT NULL,
            project_name    TEXT,
            project_col     INTEGER,
            npp_per_w       REAL,
            npp_total       REAL,
            fmv_per_w       REAL,
            dev_fee_per_w   REAL,
            target_irr      REAL,
            live_irr        REAL,
            status          TEXT NOT NULL DEFAULT 'success',
            duration_sec    REAL,
            raw_outputs     TEXT,
            error_message   TEXT,
            batch_id        TEXT
        )
    """)
    # Add batch_id column to existing databases that lack it
    try:
        conn.execute("ALTER TABLE macro_runs ADD COLUMN batch_id TEXT")
        conn.commit()
    except sqlite3.OperationalError:
        pass  # column already exists
    conn.commit()
    return conn


def save_run(conn: sqlite3.Connection, *, workbook_name: str, macro_name: str,
             project_name: str = None, project_col: int = None,
             npp_per_w: float = None, npp_total: float = None,
             fmv_per_w: float = None, dev_fee_per_w: float = None,
             target_irr: float = None, live_irr: float = None,
             status: str = "success", duration_sec: float = None,
             raw_outputs: dict = None, error_message: str = None,
             batch_id: str = None):
    conn.execute("""
        INSERT INTO macro_runs
            (workbook_name, run_timestamp, macro_name, project_name, project_col,
             npp_per_w, npp_total, fmv_per_w, dev_fee_per_w, target_irr, live_irr,
             status, duration_sec, raw_outputs, error_message, batch_id)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        workbook_name, datetime.now().isoformat(), macro_name,
        project_name, project_col,
        npp_per_w, npp_total, fmv_per_w, dev_fee_per_w,
        target_irr, live_irr,
        status, duration_sec,
        json.dumps(raw_outputs) if raw_outputs else None,
        error_message,
        batch_id,
    ))
    conn.commit()


def get_runs(conn: sqlite3.Connection, limit: int = 50):
    return conn.execute(
        "SELECT * FROM macro_runs ORDER BY run_timestamp DESC LIMIT ?", (limit,)
    ).fetchall()


def get_batch_runs(conn: sqlite3.Connection, batch_id: str):
    """Retrieve all runs belonging to a specific batch."""
    return conn.execute(
        "SELECT * FROM macro_runs WHERE batch_id = ? ORDER BY run_timestamp ASC",
        (batch_id,)
    ).fetchall()


def get_latest_run(conn: sqlite3.Connection, workbook_name: str = None):
    if workbook_name:
        return conn.execute(
            "SELECT * FROM macro_runs WHERE workbook_name = ? ORDER BY run_timestamp DESC LIMIT 1",
            (workbook_name,)
        ).fetchone()
    return conn.execute(
        "SELECT * FROM macro_runs ORDER BY run_timestamp DESC LIMIT 1"
    ).fetchone()
