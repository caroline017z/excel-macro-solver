"""dn38_solver.storage.database — SQLite persistence with typed records.

All data flows through RunRecord and ProjectResult structs.
No raw dicts crossing the boundary.
"""
from __future__ import annotations

import json
import logging
import sqlite3
from datetime import datetime, timezone
from pathlib import Path

import msgspec

from dn38_solver.config import DB_PATH
from dn38_solver.types import ProjectResult, RunRecord

log = logging.getLogger(__name__)

_SCHEMA_VERSION = 2

_CREATE_RUNS = """\
CREATE TABLE IF NOT EXISTS solver_runs (
    id               INTEGER PRIMARY KEY AUTOINCREMENT,
    workbook_name    TEXT NOT NULL,
    run_timestamp    TEXT NOT NULL,
    batch_id         TEXT NOT NULL,
    solver_mode      TEXT NOT NULL,
    total_duration   REAL NOT NULL,
    status           TEXT NOT NULL,
    error            TEXT,
    projects_json    TEXT NOT NULL
)
"""

_CREATE_META = """\
CREATE TABLE IF NOT EXISTS _meta (
    key   TEXT PRIMARY KEY,
    value TEXT NOT NULL
)
"""


def get_connection(db_path: Path = DB_PATH) -> sqlite3.Connection:
    """Open (or create) the SQLite database with WAL mode."""
    conn = sqlite3.connect(str(db_path))
    conn.row_factory = sqlite3.Row
    conn.execute("PRAGMA journal_mode=WAL")
    conn.execute(_CREATE_RUNS)
    conn.execute(_CREATE_META)
    conn.execute(
        "INSERT OR IGNORE INTO _meta (key, value) VALUES (?, ?)",
        ("schema_version", str(_SCHEMA_VERSION)),
    )
    conn.commit()
    log.debug("Database connected: %s", db_path)
    return conn


def save_run(conn: sqlite3.Connection, record: RunRecord) -> int:
    """Persist a RunRecord. Returns the new row id."""
    projects_json = msgspec.json.encode(record.projects).decode("utf-8")
    cursor = conn.execute(
        """\
        INSERT INTO solver_runs
            (workbook_name, run_timestamp, batch_id, solver_mode,
             total_duration, status, error, projects_json)
        VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        """,
        (
            record.workbook_name,
            record.run_timestamp,
            record.batch_id,
            record.solver_mode,
            record.total_duration_sec,
            record.status,
            record.error,
            projects_json,
        ),
    )
    conn.commit()
    row_id = cursor.lastrowid or 0
    log.info("Saved run id=%d (%s, %d projects)", row_id, record.status, len(record.projects))
    return row_id


def _row_to_record(row: sqlite3.Row) -> RunRecord:
    """Convert a sqlite3.Row to a RunRecord struct."""
    projects = msgspec.json.decode(
        row["projects_json"].encode("utf-8"),
        type=tuple[ProjectResult, ...],
    )
    return RunRecord(
        id=row["id"],
        workbook_name=row["workbook_name"],
        run_timestamp=row["run_timestamp"],
        batch_id=row["batch_id"],
        solver_mode=row["solver_mode"],
        total_duration_sec=row["total_duration"],
        status=row["status"],
        error=row["error"],
        projects=projects,
    )


def get_runs(conn: sqlite3.Connection, limit: int = 50) -> tuple[RunRecord, ...]:
    """Fetch recent runs, newest first."""
    rows = conn.execute(
        "SELECT * FROM solver_runs ORDER BY run_timestamp DESC LIMIT ?",
        (limit,),
    ).fetchall()
    return tuple(_row_to_record(r) for r in rows)


def get_run_by_id(conn: sqlite3.Connection, run_id: int) -> RunRecord | None:
    """Fetch a single run by id."""
    row = conn.execute(
        "SELECT * FROM solver_runs WHERE id = ?", (run_id,)
    ).fetchone()
    return _row_to_record(row) if row else None


def get_batch_runs(conn: sqlite3.Connection, batch_id: str) -> tuple[RunRecord, ...]:
    """Fetch all runs in a batch."""
    rows = conn.execute(
        "SELECT * FROM solver_runs WHERE batch_id = ? ORDER BY run_timestamp ASC",
        (batch_id,),
    ).fetchall()
    return tuple(_row_to_record(r) for r in rows)


def get_latest_run(
    conn: sqlite3.Connection,
    workbook_name: str | None = None,
) -> RunRecord | None:
    """Fetch the most recent run, optionally filtered by workbook."""
    if workbook_name:
        row = conn.execute(
            "SELECT * FROM solver_runs WHERE workbook_name = ? ORDER BY run_timestamp DESC LIMIT 1",
            (workbook_name,),
        ).fetchone()
    else:
        row = conn.execute(
            "SELECT * FROM solver_runs ORDER BY run_timestamp DESC LIMIT 1",
        ).fetchone()
    return _row_to_record(row) if row else None


def now_iso() -> str:
    """Current timestamp in ISO-8601 UTC."""
    return datetime.now(timezone.utc).isoformat()
