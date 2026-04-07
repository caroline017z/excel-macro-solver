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
# GoalSeek operations (COM worker contract)
# ---------------------------------------------------------------------------

class GoalSeekOp(msgspec.Struct, frozen=True, kw_only=True):
    """One GoalSeek operation for the COM worker."""
    target_sheet: str
    target_cell: str      # A1-style — the cell whose value should match goal
    goal_sheet: str
    goal_cell: str        # A1-style — cell holding the target value
    changing_sheet: str
    changing_cell: str    # A1-style — cell that GoalSeek adjusts
    lower_bound: float
    upper_bound: float


class SolveTask(msgspec.Struct, frozen=True, kw_only=True):
    """Full payload sent to com_worker.py via JSON stdin."""
    workbook_path: str
    project_offset: int
    project_col_letter: str
    project_name: str
    calc_core_sheets: tuple[str, ...]
    project_index_cell: CellAddress
    holdco_cell: CellAddress
    equity_cell: CellAddress
    equity_target_cell: CellAddress
    total_uses_cell: CellAddress
    goal_seeks_phase1: tuple[GoalSeekOp, ...]
    goal_seeks_phase2: tuple[GoalSeekOp, ...]
    max_iterations: int
    max_gs_retries: int
    irr_tolerance: float
    equity_tolerance: float
    gs_max_change: float
    gs_max_iterations: int
    read_cells: tuple[CellAddress, ...]
    saved_workbook_suffix: str = "_SOLVED"


class SolveStatus(str, enum.Enum):
    """Solve outcome — str enum for JSON serialization."""
    CONVERGED = "converged"
    NOT_CONVERGED = "not_converged"
    ERROR = "error"
    TIMEOUT = "timeout"
    DRY_RUN = "dry_run"


class SolveResult(msgspec.Struct, frozen=True, kw_only=True):
    """Result returned from com_worker.py via JSON stdout."""
    status: str               # SolveStatus value
    solved_values: dict[str, float | str | None]
    iterations_used: int
    duration_sec: float
    saved_to: str | None = None
    error: str | None = None


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
