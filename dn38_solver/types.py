"""dn38_solver.types — All shared type definitions.

Every struct uses frozen=True (immutable) and kw_only=True per
python-modern-performance-standards skill. msgspec only, never pydantic.
"""
from __future__ import annotations

import enum
import msgspec


# ---------------------------------------------------------------------------
# Cell references
# ---------------------------------------------------------------------------

class CellAddress(msgspec.Struct, frozen=True, kw_only=True):
    """Cell identified by sheet name + A1-style address."""
    sheet: str
    address: str


class CellRef(msgspec.Struct, frozen=True, kw_only=True):
    """Cell identified by sheet name + numeric row/col (1-based)."""
    sheet: str
    row: int
    col: int
    label: str = ""


# ---------------------------------------------------------------------------
# Project discovery
# ---------------------------------------------------------------------------

class ProjectInfo(msgspec.Struct, frozen=True, kw_only=True):
    """A single project discovered in Project Inputs."""
    name: str
    col: int
    col_letter: str
    offset: int       # col - BASE_COL, written to F2 for OFFSET routing
    toggle: bool


# ---------------------------------------------------------------------------
# Solve task (COM worker contract)
# ---------------------------------------------------------------------------

class SolveTask(msgspec.Struct, frozen=True, kw_only=True):
    """Per-project solve metadata. Built by sequence.build_solve_task; consumed
    by direct_runner.run_direct.

    Trimmed to the fields the runner actually reads. The convergence
    constants and GoalSeek templates that used to ride on this struct are
    owned by SolveHeadless.bas now -- VBA reads them directly as Private
    Const, so passing them across the COM boundary was always inert.
    """
    workbook_path: str
    project_offset: int
    project_col_letter: str
    project_name: str
    project_index_cell: CellAddress
    read_cells: tuple[CellAddress, ...]


class SolveStatus(str, enum.Enum):
    """Solve outcome — str enum for JSON serialization."""
    CONVERGED = "converged"
    NOT_CONVERGED = "not_converged"
    ERROR = "error"
    TIMEOUT = "timeout"
    DRY_RUN = "dry_run"


# ---------------------------------------------------------------------------
# Project results & run records (SQLite persistence)
# ---------------------------------------------------------------------------

class ProjectResult(msgspec.Struct, frozen=True, kw_only=True):
    """Solved result for one project."""
    name: str
    col: int
    col_letter: str
    npp_per_w: float | None = None
    npp_total: float | None = None
    dev_fee_per_w: float | None = None
    fmv_per_w: float | None = None
    target_irr: float | None = None
    live_irr: float | None = None
    appraisal_live: float | None = None
    wacc_target: float | None = None
    dscr_multiple: float | None = None
    equity_pct: float | None = None
    converged: bool = False
    convergence_tier: str = "none"   # "strict" | "relaxed" | "none" | "not_attempted"
    iterations: int = 0


class RunRecord(msgspec.Struct, frozen=True, kw_only=True):
    """Complete record of a solver run."""
    workbook_name: str
    run_timestamp: str
    batch_id: str
    solver_mode: str
    projects: tuple[ProjectResult, ...]
    total_duration_sec: float
    status: str
    error: str | None = None
    id: int | None = None


def convergence_label(p: ProjectResult) -> str:
    """Render a project's convergence outcome for terminal tables.

    Strict converged -> OK   (equity within +/-0.25pp of 10% target)
    Relaxed-tier     -> OK*  (within +/-0.5pp / 5x gap tol, --allow-relaxed-eligible)
    Not attempted    -> SKIP (worker crashed before reaching this project)
    Otherwise        -> CHECK

    Callers that print a table should follow with a one-line legend below
    when any OK* labels appear so the asterisk is self-documenting.
    """
    if p.convergence_tier == "strict":
        return "OK"
    if p.convergence_tier == "relaxed":
        return "OK*"
    if p.convergence_tier == "not_attempted":
        return "SKIP"
    return "CHECK"


RELAXED_LEGEND = (
    "OK* = relaxed tier (equity +/-0.5pp, gaps <= 5x tol; "
    "--allow-relaxed-eligible). Strict band is +/-0.25pp."
)


# ---------------------------------------------------------------------------
# Diff reporting
# ---------------------------------------------------------------------------

class CellChange(msgspec.Struct, frozen=True, kw_only=True):
    """One cell that changed between pre and post snapshots."""
    sheet: str
    row: int
    col: int
    label: str
    before: float | str | None
    after: float | str | None
    delta: float | None
    pct_change: float | None
