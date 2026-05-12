"""dn38_solver.types — All shared type definitions.

Every struct uses frozen=True (immutable) and kw_only=True per
python-modern-performance-standards skill. msgspec only, never pydantic.
"""
from __future__ import annotations

import enum
from typing import Literal

import msgspec


# Convergence-tier constants. Keep this Literal in sync with
# convergence_label() and orchestrator._parse_project_result(). msgspec
# validates Literal values at decode time, so a SQLite row with a stale
# tier string surfaces immediately rather than slipping through to the
# rollup as a silently-mis-categorized row.
ConvergenceTier = Literal["strict", "relaxed", "none", "not_attempted"]


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
    convergence_tier: ConvergenceTier = "none"
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


class RunMetrics(msgspec.Struct, frozen=True, kw_only=True):
    """Sidecar metrics for a solver run.

    Holds the non-RunRecord values surfaced in the end-of-run summary
    (parallel speedup, merge path, worker count). Persisted to the
    `solver_run_metrics` table by `dn38_solver.storage.database`, keyed
    by `run_id` (FK to `solver_runs.id`).

    Why a sidecar struct instead of fields on RunRecord: every run-level
    signal added since v0.1 has faced the same false choice — bolt onto
    RunRecord and break the SQLite schema, or leave it in `batch_result`
    and lose it after the log line. The sidecar pattern lets the stable
    persistence shape stay stable while runtime-only metrics get their
    own home and can grow without rippling through every consumer of
    RunRecord (Streamlit dashboard, --show-checkpoints CLI, etc.).

    `merge_path`: "openpyxl" | "vba_fallback" | "copy_master" | None.
        None for single-worker runs (run_direct doesn't merge).
    `workers_used`: clamped count of workers actually spawned.
    `estimated_sequential_sec`: sum of every attempted project's
        meta.solve_seconds. Comparing wall time to this gives the actual
        parallel speedup; > 1 means parallel paid off.
    """
    run_id: int
    workers_used: int
    merge_path: str | None = None
    estimated_sequential_sec: float | None = None
    wall_time_sec: float | None = None


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
