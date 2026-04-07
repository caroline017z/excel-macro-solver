"""
38DN Excel Macro Runner — VP Review App Bridge
Reads macro run results from SQLite and returns data in the same format
as the VP Review App's load_pricing_model().
"""
import json
import sqlite3
from pathlib import Path

from db import get_connection


# Row-label mapping used by the macro runner's OUTPUT_ROWS (config.py)
_MACRO_OUTPUT_LABEL_TO_ROW = {
    "Dev Fee ($/W)": 32,
    "FMV Calculated ($/W)": 33,
    "NPP ($/W)": 38,
    "NPP ($)": 39,
    "FMV WACC (Target)": 30,
    "Live Appraisal IRR": 31,
    "Target IRR": 36,
    "Live Levered Pre-Tax IRR": 37,
}


def list_available_runs(db_path):
    """Return a list of runs for a selectbox.

    Each entry is a dict with:
        id, workbook_name, run_timestamp, project_name, project_col,
        status, batch_id, display_label
    """
    conn = get_connection(Path(db_path))
    rows = conn.execute(
        "SELECT id, workbook_name, run_timestamp, project_name, project_col, "
        "       status, batch_id "
        "FROM macro_runs "
        "WHERE status = 'success' "
        "ORDER BY run_timestamp DESC "
        "LIMIT 200"
    ).fetchall()
    conn.close()

    results = []
    for r in rows:
        ts_short = (r["run_timestamp"] or "")[:19]
        wb = r["workbook_name"] or "unknown"
        proj = r["project_name"] or "—"
        label = f"{ts_short}  |  {wb}  |  {proj}"
        results.append({
            "id": r["id"],
            "workbook_name": r["workbook_name"],
            "run_timestamp": r["run_timestamp"],
            "project_name": r["project_name"],
            "project_col": r["project_col"],
            "status": r["status"],
            "batch_id": r["batch_id"],
            "display_label": label,
        })
    return results


def list_available_batches(db_path):
    """Return a list of distinct batches (timestamp + workbook combo).

    Each batch may contain multiple projects solved in one run.
    """
    conn = get_connection(Path(db_path))
    rows = conn.execute(
        "SELECT batch_id, workbook_name, MIN(run_timestamp) as first_ts, "
        "       COUNT(*) as project_count "
        "FROM macro_runs "
        "WHERE status = 'success' AND batch_id IS NOT NULL "
        "GROUP BY batch_id "
        "ORDER BY first_ts DESC "
        "LIMIT 100"
    ).fetchall()
    conn.close()

    results = []
    for r in rows:
        ts_short = (r["first_ts"] or "")[:19]
        wb = r["workbook_name"] or "unknown"
        n = r["project_count"]
        label = f"{ts_short}  |  {wb}  |  {n} project(s)"
        results.append({
            "batch_id": r["batch_id"],
            "workbook_name": r["workbook_name"],
            "first_ts": r["first_ts"],
            "project_count": r["project_count"],
            "display_label": label,
        })
    return results


def _row_from_run(run_row):
    """Convert a single macro_runs DB row into a VP-app project dict.

    The VP app's load_pricing_model() returns per-project dicts shaped like:
        {
            "name": str,
            "toggle": bool,
            "col_letter": str,
            "data": {row_int: value, ...},
            "rate_comps": {1: {...}, ...},
            "dscr_label": ...,
            "dscr_schedule": {},
        }
    We populate `data` with the output rows we have from the macro run, and
    leave the rest empty/default so the VP app can at least render the
    portfolio, comparison, and review tabs for the solved values.
    """
    import openpyxl.utils  # lightweight import for column letter

    raw = run_row["raw_outputs"]
    raw_dict = json.loads(raw) if raw else {}
    project_outputs = raw_dict.get("project_outputs", {})

    # Build the data dict keyed by row number
    data = {}
    for label, value in project_outputs.items():
        row_num = _MACRO_OUTPUT_LABEL_TO_ROW.get(label)
        if row_num is not None:
            data[row_num] = value

    # Also populate from top-level columns stored directly in the DB row
    if run_row["npp_per_w"] is not None:
        data.setdefault(38, run_row["npp_per_w"])
    if run_row["npp_total"] is not None:
        data.setdefault(39, run_row["npp_total"])
    if run_row["fmv_per_w"] is not None:
        data.setdefault(33, run_row["fmv_per_w"])
    if run_row["dev_fee_per_w"] is not None:
        data.setdefault(32, run_row["dev_fee_per_w"])
    if run_row["target_irr"] is not None:
        data.setdefault(36, run_row["target_irr"])
    if run_row["live_irr"] is not None:
        data.setdefault(37, run_row["live_irr"])

    col_idx = run_row["project_col"] or 6
    try:
        col_letter = openpyxl.utils.get_column_letter(col_idx)
    except Exception:
        col_letter = "?"

    name = run_row["project_name"] or "Unknown Project"
    # Clean multiline names the same way the VP app does
    clean_name = " | ".join(
        line.strip() for line in str(name).strip().splitlines() if line.strip()
    )

    return {
        "name": clean_name,
        "toggle": True,  # macro runner only saves active (toggled-on) projects
        "col_letter": col_letter,
        "data": data,
        "rate_comps": {},
        "dscr_label": None,
        "dscr_schedule": {},
    }


def load_run_as_model(db_path, run_id):
    """Load a single macro run from SQLite and return data in load_pricing_model() format.

    Returns:
        {"projects": {col_int: project_dict, ...}, "ops_sandbox": {}}
    """
    conn = get_connection(Path(db_path))
    row = conn.execute(
        "SELECT * FROM macro_runs WHERE id = ?", (run_id,)
    ).fetchone()
    conn.close()

    if row is None:
        return {"projects": {}, "ops_sandbox": {}}

    col_idx = row["project_col"] or 6
    project = _row_from_run(row)

    return {
        "projects": {col_idx: project},
        "ops_sandbox": {},
    }


def load_batch_as_model(db_path, batch_id):
    """Load all runs in a batch and return combined data in load_pricing_model() format.

    Returns:
        {"projects": {col_int: project_dict, ...}, "ops_sandbox": {}}
    """
    conn = get_connection(Path(db_path))
    rows = conn.execute(
        "SELECT * FROM macro_runs WHERE batch_id = ? AND status = 'success' "
        "ORDER BY project_col ASC",
        (batch_id,)
    ).fetchall()
    conn.close()

    projects = {}
    for row in rows:
        col_idx = row["project_col"] or 6
        projects[col_idx] = _row_from_run(row)

    return {
        "projects": projects,
        "ops_sandbox": {},
    }


def load_workbook_latest_as_model(db_path, workbook_name):
    """Load the latest runs for a given workbook name.

    Groups by the most recent batch_id (or timestamp if no batch) and returns
    all projects from that run in load_pricing_model() format.
    """
    conn = get_connection(Path(db_path))

    # First try to find the latest batch for this workbook
    batch_row = conn.execute(
        "SELECT batch_id FROM macro_runs "
        "WHERE workbook_name = ? AND status = 'success' AND batch_id IS NOT NULL "
        "ORDER BY run_timestamp DESC LIMIT 1",
        (workbook_name,)
    ).fetchone()

    if batch_row and batch_row["batch_id"]:
        rows = conn.execute(
            "SELECT * FROM macro_runs WHERE batch_id = ? AND status = 'success' "
            "ORDER BY project_col ASC",
            (batch_row["batch_id"],)
        ).fetchall()
    else:
        # Fall back to latest single run
        rows = conn.execute(
            "SELECT * FROM macro_runs "
            "WHERE workbook_name = ? AND status = 'success' "
            "ORDER BY run_timestamp DESC LIMIT 1",
            (workbook_name,)
        ).fetchall()

    conn.close()

    projects = {}
    for row in rows:
        col_idx = row["project_col"] or 6
        projects[col_idx] = _row_from_run(row)

    return {
        "projects": projects,
        "ops_sandbox": {},
    }
