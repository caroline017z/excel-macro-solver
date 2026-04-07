"""
38DN Excel Macro Runner — Configuration
"""
from pathlib import Path

# Project paths
PROJECT_DIR = Path(__file__).parent
DB_PATH = PROJECT_DIR / "results.db"

# Default batch directory
DEFAULT_BATCH_DIR = Path(
    r"C:\Users\CarolineZepecki\Box\DF\DevEngine"
)

# Default workbook (override via CLI)
DEFAULT_WORKBOOK = Path(
    r"C:\Users\CarolineZepecki\Box\DF\DevEngine"
    r"\38DN-IL_DevEngine_PricingModel_2026.04.06.xlsm"
)

# Macro name variants to try (in order)
# VBA modules found: CalcModelCore, Module2, Project_Toggle, Project, Calc
MACRO_VARIANTS = [
    "SolveMinEquityHoldco",
    "Solve_Min_Equity_Holdco",
    "SolveMinEquityHoldCo",
    "CalcModelCore.SolveMinEquityHoldco",
    "CalcModelCore.Solve_Min_Equity_Holdco",
    "Module2.SolveMinEquityHoldco",
    "Module2.Solve_Min_Equity_Holdco",
]

# Output rows to extract from "Project Inputs" sheet after macro runs
OUTPUT_ROWS = {
    32: "Dev Fee ($/W)",
    33: "FMV Calculated ($/W)",
    38: "NPP ($/W)",
    39: "NPP ($)",
    30: "FMV WACC (Target)",
    31: "Live Appraisal IRR",
    36: "Target IRR",
    37: "Live Levered Pre-Tax IRR",
}

# Sheets to capture full snapshots from
SNAPSHOT_SHEETS = ["Dashboard", "Corp Model Output", "NPP Calc"]

# Project Inputs layout
PROJECT_NAME_ROW = 4
PROJECT_TOGGLE_ROW = 7
PROJECT_COL_START = 6
PROJECT_COL_END = 88
