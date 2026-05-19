"""dn38_solver.config — Single source of truth for all Excel references.

When the Excel model changes (new rows, new sheets), ONLY this file updates.
Every other module reads from here. No magic numbers elsewhere.
"""
from __future__ import annotations

from pathlib import Path

from dn38_solver.types import CellAddress

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
    # Per-project DSCR Multiple input (Tranche 7.12). Default cell
    # formula is ='PT Returns'!$F$129; macro replaces with the converged
    # numeric so rows 31/37 can stay as sticky-IF formulas with the
    # right per-project cached value. Restore the formula to put Min
    # Equity back into fully dynamic solve.
    371: "Min Equity DSCR Multiple",
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
# Solver constants
#
# Convergence thresholds (MAX_ITER, EQUITY_FINAL_TOL, IRR_TOLERANCE,
# APPR_TOLERANCE, GS_MAXITER_*, etc.) live in SolveHeadless.bas as Private
# Const. The macro reads them directly; Python no longer mirrors them
# because no live module passes them across the COM boundary. Edit the
# .bas file when these need to move.
# ---------------------------------------------------------------------------

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
