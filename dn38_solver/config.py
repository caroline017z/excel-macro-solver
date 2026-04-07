"""dn38_solver.config — Single source of truth for all Excel references.

When the Excel model changes (new rows, new sheets), ONLY this file updates.
Every other module reads from here. No magic numbers elsewhere.
"""
from __future__ import annotations

from pathlib import Path

from dn38_solver.types import CellAddress, GoalSeekOp

# ---------------------------------------------------------------------------
# Paths
# ---------------------------------------------------------------------------

PROJECT_DIR: Path = Path(__file__).resolve().parent.parent
DB_PATH: Path = PROJECT_DIR / "results.db"
DEFAULT_WORKBOOK: Path = Path(
    r"C:\Users\CarolineZepecki\Box\DF\DevEngine"
    r"\38DN-IL_DevEngine_PricingModel_2026.04.06.xlsm"
)
DEFAULT_BATCH_DIR: Path = Path(
    r"C:\Users\CarolineZepecki\Box\DF\DevEngine"
)

# ---------------------------------------------------------------------------
# Project Inputs sheet layout
# ---------------------------------------------------------------------------

PROJECT_NAME_ROW: int = 4
PROJECT_TOGGLE_ROW: int = 7
PROJECT_COL_RANGE: range = range(8, 68)   # columns H through BQ (60 cols)
BASE_COL: int = 7                          # column G (OFFSET base for F2)

# ---------------------------------------------------------------------------
# Output rows — THE one canonical mapping (row_num -> label)
# ---------------------------------------------------------------------------

OUTPUT_ROWS: dict[int, str] = {
    30: "FMV WACC (Target)",
    31: "Live Appraisal IRR",
    32: "Dev Fee ($/W)",
    33: "FMV Calculated ($/W)",
    36: "Target IRR",
    37: "Live Levered Pre-Tax IRR",
    38: "NPP ($/W)",
    39: "NPP ($)",
}

# Computed inverse — derived, never manually maintained
LABEL_TO_ROW: dict[str, int] = {v: k for k, v in OUTPUT_ROWS.items()}

# ---------------------------------------------------------------------------
# CalcModelCore — sheet recalc order (exact match from VBA)
# ---------------------------------------------------------------------------

CALC_CORE_SHEETS: tuple[str, ...] = (
    "Project Inputs",
    "Rate Curves",
    "Ops Sandbox",
    "Global",
    "Operations",
    "Capex",
    "Safe Harbor",
    "CL",
    "Perm Debt",
    "Tax Equity",
    "Appraisal",
    "NPP Calc",
    "PT Returns",
)

CALC_OUTPUT_SHEETS: tuple[str, ...] = (
    "Portfolio",
    "AT Returns_WIP",
    "Corp Model Output",
    "Cust Prop",
    "Dashboard",
    "Table",
    "Waterfall Sensitivity",
)

# ---------------------------------------------------------------------------
# Named cell addresses (no magic strings anywhere else)
# ---------------------------------------------------------------------------

CELL_PROJECT_INDEX = CellAddress(sheet="Project Inputs", address="F2")
CELL_HOLDCO_TOGGLE = CellAddress(sheet="PT Returns", address="C134")
CELL_EQUITY = CellAddress(sheet="PT Returns", address="C128")
CELL_EQUITY_TARGET = CellAddress(sheet="PT Returns", address="F128")
CELL_DSCR_MULTIPLE = CellAddress(sheet="PT Returns", address="F129")
CELL_TOTAL_USES = CellAddress(sheet="PT Returns", address="C130")
CELL_IRR_LIVE = CellAddress(sheet="Project Inputs", address="F37")
CELL_IRR_TARGET = CellAddress(sheet="Project Inputs", address="F36")
CELL_APPRAISAL_LIVE = CellAddress(sheet="Project Inputs", address="F31")
CELL_WACC_TARGET = CellAddress(sheet="Project Inputs", address="F30")

# ---------------------------------------------------------------------------
# Solver constants (exact match from VBA SolveMinEquityWithHoldCo)
# ---------------------------------------------------------------------------

MAX_ITERATIONS: int = 8
MAX_GS_RETRIES: int = 6
EQUITY_TOLERANCE: float = 0.005     # +/- 0.5pp = 5% of 10% target
IRR_TOLERANCE: float = 0.0003       # 0.03%
APPRAISAL_TOLERANCE: float = 0.0003 # 0.03%
DSCR_BOUNDS: tuple[float, float] = (0.5, 5.0)
NPP_BOUNDS: tuple[float, float] = (-2.0, 5.0)
DEV_FEE_BOUNDS: tuple[float, float] = (0.0, 10.0)

# GoalSeek precision (mirrors VBA SetGoalSeekPrecision)
GS_MAX_CHANGE: float = 0.00001
GS_MAX_ITERATIONS: int = 1000

# ---------------------------------------------------------------------------
# GoalSeek operation templates
# Per-project cells use "{col}" placeholder, replaced at runtime
# ---------------------------------------------------------------------------

GOALSEEK_PHASE1: tuple[GoalSeekOp, ...] = (
    GoalSeekOp(
        target_sheet="PT Returns", target_cell="C128",
        goal_sheet="PT Returns", goal_cell="F128",
        changing_sheet="PT Returns", changing_cell="F129",
        lower_bound=DSCR_BOUNDS[0], upper_bound=DSCR_BOUNDS[1],
    ),
)

GOALSEEK_PHASE2_TEMPLATES: tuple[GoalSeekOp, ...] = (
    GoalSeekOp(
        target_sheet="Project Inputs", target_cell="F37",
        goal_sheet="Project Inputs", goal_cell="F36",
        changing_sheet="Project Inputs", changing_cell="{col}38",
        lower_bound=NPP_BOUNDS[0], upper_bound=NPP_BOUNDS[1],
    ),
    GoalSeekOp(
        target_sheet="Project Inputs", target_cell="F31",
        goal_sheet="Project Inputs", goal_cell="F30",
        changing_sheet="Project Inputs", changing_cell="{col}32",
        lower_bound=DEV_FEE_BOUNDS[0], upper_bound=DEV_FEE_BOUNDS[1],
    ),
)

# ---------------------------------------------------------------------------
# Cells to read back after solve (per-project, {col} templated)
# ---------------------------------------------------------------------------

READ_CELLS_TEMPLATES: tuple[CellAddress, ...] = (
    CellAddress(sheet="Project Inputs", address="F37"),
    CellAddress(sheet="Project Inputs", address="F31"),
    CellAddress(sheet="Project Inputs", address="F30"),
    CellAddress(sheet="Project Inputs", address="F36"),
    CellAddress(sheet="PT Returns", address="F129"),
    CellAddress(sheet="PT Returns", address="C128"),
    CellAddress(sheet="PT Returns", address="C130"),
    CellAddress(sheet="Project Inputs", address="{col}38"),
    CellAddress(sheet="Project Inputs", address="{col}32"),
    CellAddress(sheet="Project Inputs", address="{col}33"),
    CellAddress(sheet="Project Inputs", address="{col}39"),
)

# ---------------------------------------------------------------------------
# Snapshot sheets for diff reporting
# ---------------------------------------------------------------------------

SNAPSHOT_RANGES: dict[str, tuple[int, int, int, int]] = {
    # sheet_name: (min_row, max_row, min_col, max_col)
    "Dashboard": (1, 50, 1, 10),
    "NPP Calc": (1, 30, 1, 6),
    "Corp Model Output": (1, 50, 1, 10),
}
