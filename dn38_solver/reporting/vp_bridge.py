"""dn38_solver.reporting.vp_bridge — VP Review App data adapter.

Translates RunRecord/ProjectResult into the dict shape expected by
the VP Review App's load_pricing_model() return format.
"""
from __future__ import annotations

import logging
from pathlib import Path

from dn38_solver.types import ProjectResult, RunRecord
from dn38_solver.config import DB_PATH
from dn38_solver.storage.database import get_connection, get_run_by_id, get_runs

log = logging.getLogger(__name__)


def _project_to_vp_data(proj: ProjectResult) -> dict[str, object]:
    """Convert a ProjectResult to the VP app's per-project dict shape."""
    data: dict[int, float | str | None] = {}

    # Map solved values to row numbers
    field_map: dict[str, int] = {
        "npp_per_w": 38,
        "npp_total": 39,
        "dev_fee_per_w": 32,
        "fmv_per_w": 33,
        "target_irr": 36,
        "live_irr": 37,
        "appraisal_live": 31,
        "wacc_target": 30,
    }

    for attr, row in field_map.items():
        val = getattr(proj, attr, None)
        if val is not None:
            data[row] = val

    return {
        "name": proj.name,
        "toggle": True,
        "col_letter": proj.col_letter,
        "data": data,
        "rate_comps": {},
        "dscr_label": None,
        "dscr_schedule": {},
    }


def run_to_vp_model(record: RunRecord) -> dict[str, object]:
    """Convert a RunRecord to the VP app's load_pricing_model() format.

    Returns {"projects": {col: {...}}, "ops_sandbox": {}}.
    """
    projects: dict[int, dict[str, object]] = {}
    for proj in record.projects:
        projects[proj.col] = _project_to_vp_data(proj)

    return {
        "projects": projects,
        "ops_sandbox": {},
    }


def list_available_runs(db_path: Path = DB_PATH, limit: int = 50) -> list[dict[str, str | int]]:
    """List runs for a selectbox in the VP app sidebar."""
    conn = get_connection(db_path)
    runs = get_runs(conn, limit)
    conn.close()

    return [
        {
            "id": r.id,
            "label": f"{r.run_timestamp[:19]} | {r.workbook_name} | {len(r.projects)} proj | {r.status}",
            "status": r.status,
        }
        for r in runs
        if r.id is not None
    ]


def load_run_for_vp(
    run_id: int,
    db_path: Path = DB_PATH,
) -> dict[str, object] | None:
    """Load a specific run and return it in VP app format."""
    conn = get_connection(db_path)
    record = get_run_by_id(conn, run_id)
    conn.close()

    if record is None:
        return None
    return run_to_vp_model(record)
