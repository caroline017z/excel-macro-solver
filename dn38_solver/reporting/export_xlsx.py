"""dn38_solver.reporting.export_xlsx — Clean summary .xlsx for sharing.

Generates a branded one-pager with solve results, suitable for email.
Includes NPP, Dev Fee, DSCR Multiple, and all convergence details
so the run is fully auditable and replicable.
"""
from __future__ import annotations

import io
import logging
from datetime import datetime

import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from dn38_solver.types import ProjectResult, RunRecord

log = logging.getLogger(__name__)

# ---------------------------------------------------------------------------
# 38DN brand styles (Aptos Narrow / navy / teal)
# ---------------------------------------------------------------------------

_FONT_TITLE = Font(name="Aptos Display", size=14, bold=True, color="FFFFFF")
_FONT_SUBTITLE = Font(name="Aptos Narrow", size=10, color="A0A8C0")
_FONT_HEADER = Font(name="Aptos Display", size=10, bold=True, color="FFFFFF")
_FONT_LABEL = Font(name="Aptos Narrow", size=10, color="050D25")
_FONT_DATA = Font(name="Aptos Narrow", size=10, color="050D25")
_FONT_DATA_BOLD = Font(name="Aptos Narrow", size=10, bold=True, color="050D25")
_FONT_PASS = Font(name="Aptos Narrow", size=10, bold=True, color="3A7D44")
_FONT_FAIL = Font(name="Aptos Narrow", size=10, bold=True, color="B83230")

_FILL_NAVY = PatternFill(start_color="050D25", end_color="050D25", fill_type="solid")
_FILL_NAVY2 = PatternFill(start_color="212B48", end_color="212B48", fill_type="solid")
_FILL_LIGHT = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
_FILL_WHITE = PatternFill(fill_type=None)

_THIN = Side(style="thin", color="D0D4DC")
_BORDER = Border(bottom=_THIN)
_ALIGN_LEFT = Alignment(horizontal="left", vertical="center")
_ALIGN_CENTER = Alignment(horizontal="center", vertical="center")
_ALIGN_RIGHT = Alignment(horizontal="right", vertical="center")

# Number formats
_NF_DPW = '"$"0.0000'
_NF_PCT = "0.00%"
_NF_DSCR = "0.0000\"x\""
_NF_DOLLAR = '"$"#,##0'
_NF_SEC = '0.0"s"'


def generate_summary_xlsx(record: RunRecord) -> io.BytesIO:
    """Generate a branded summary .xlsx from a RunRecord.

    Layout:
      Row 1-2: Navy banner with title + timestamp
      Row 3: Blank spacer
      Row 4: Column headers (navy)
      Row 5+: One row per project with key metrics
      Bottom: Run metadata (batch ID, solver mode, total duration)

    Columns:
      Project | State | MW(DC) | NPP ($/W) | Dev Fee ($/W) | FMV ($/W) |
      DSCR Multiple | Target IRR | Live IRR | Appraisal IRR | WACC Target |
      Equity % | Iterations | Status
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Solve Summary"

    # --- Column definitions ---
    columns = [
        ("Project", 30, _ALIGN_LEFT, None),
        ("NPP ($/W)", 14, _ALIGN_RIGHT, _NF_DPW),
        ("Step-Up Dev Fee ($/W)", 20, _ALIGN_RIGHT, _NF_DPW),
        ("FMV ($/W)", 14, _ALIGN_RIGHT, _NF_DPW),
        ("NPP ($)", 16, _ALIGN_RIGHT, _NF_DOLLAR),
        ("DSCR Multiple", 16, _ALIGN_RIGHT, _NF_DSCR),
        ("Target IRR", 13, _ALIGN_RIGHT, _NF_PCT),
        ("Live IRR", 13, _ALIGN_RIGHT, _NF_PCT),
        ("IRR Gap", 11, _ALIGN_RIGHT, _NF_PCT),
        ("Appraisal IRR", 15, _ALIGN_RIGHT, _NF_PCT),
        ("WACC Target", 13, _ALIGN_RIGHT, _NF_PCT),
        ("Appraisal Gap", 14, _ALIGN_RIGHT, _NF_PCT),
        ("Equity %", 11, _ALIGN_RIGHT, _NF_PCT),
        ("Iterations", 11, _ALIGN_CENTER, None),
        ("Status", 14, _ALIGN_CENTER, None),
    ]

    n_cols = len(columns)

    # Set column widths
    for i, (_, width, _, _) in enumerate(columns, 1):
        ws.column_dimensions[get_column_letter(i)].width = width

    # --- Row 1: Navy banner ---
    ws.merge_cells(start_row=1, start_column=1, end_row=1, end_column=n_cols)
    title_cell = ws.cell(row=1, column=1)
    title_cell.value = "38DN Pricing Model — Solve Summary"
    title_cell.font = _FONT_TITLE
    title_cell.fill = _FILL_NAVY
    title_cell.alignment = _ALIGN_LEFT
    ws.row_dimensions[1].height = 32

    for c in range(2, n_cols + 1):
        ws.cell(row=1, column=c).fill = _FILL_NAVY

    # --- Row 2: Subtitle with metadata ---
    ws.merge_cells(start_row=2, start_column=1, end_row=2, end_column=n_cols)
    sub_cell = ws.cell(row=2, column=1)
    ts_display = record.run_timestamp[:19].replace("T", " ")
    sub_cell.value = (
        f"Workbook: {record.workbook_name}  |  "
        f"Run: {ts_display} UTC  |  "
        f"Batch: {record.batch_id}  |  "
        f"Mode: {record.solver_mode}  |  "
        f"Duration: {record.total_duration_sec:.1f}s"
    )
    sub_cell.font = _FONT_SUBTITLE
    sub_cell.fill = _FILL_NAVY2
    sub_cell.alignment = _ALIGN_LEFT
    ws.row_dimensions[2].height = 22

    for c in range(2, n_cols + 1):
        ws.cell(row=2, column=c).fill = _FILL_NAVY2

    # --- Row 3: Spacer ---
    ws.row_dimensions[3].height = 6

    # --- Row 4: Column headers ---
    for i, (label, _, align, _) in enumerate(columns, 1):
        cell = ws.cell(row=4, column=i, value=label)
        cell.font = _FONT_HEADER
        cell.fill = _FILL_NAVY
        cell.alignment = align
        cell.border = _BORDER
    ws.row_dimensions[4].height = 24

    # --- Data rows (one per project) ---
    for row_idx, proj in enumerate(record.projects, 5):
        irr_gap = (
            abs((proj.live_irr or 0) - (proj.target_irr or 0))
            if proj.live_irr is not None and proj.target_irr is not None
            else None
        )
        appr_gap = (
            abs((proj.appraisal_live or 0) - (proj.wacc_target or 0))
            if proj.appraisal_live is not None and proj.wacc_target is not None
            else None
        )

        values = [
            proj.name,
            proj.npp_per_w,
            proj.dev_fee_per_w,
            proj.fmv_per_w,
            proj.npp_total,
            proj.dscr_multiple,
            proj.target_irr,
            proj.live_irr,
            irr_gap,
            proj.appraisal_live,
            proj.wacc_target,
            appr_gap,
            proj.equity_pct,
            proj.iterations,
            "CONVERGED" if proj.converged else "CHECK",
        ]

        fill = _FILL_LIGHT if row_idx % 2 == 1 else _FILL_WHITE
        for col_idx, val in enumerate(values, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            _, _, align, nf = columns[col_idx - 1]
            cell.font = _FONT_DATA
            cell.alignment = align
            cell.fill = fill
            cell.border = _BORDER
            if nf and val is not None:
                cell.number_format = nf

        # Style the status column
        status_cell = ws.cell(row=row_idx, column=n_cols)
        status_cell.font = _FONT_PASS if proj.converged else _FONT_FAIL

        # Bold the key solve outputs
        for bold_col in (2, 3, 6):  # NPP, Dev Fee, DSCR
            ws.cell(row=row_idx, column=bold_col).font = _FONT_DATA_BOLD

    # --- Bottom: Convergence reference ---
    bottom_row = 5 + len(record.projects) + 1
    ws.merge_cells(start_row=bottom_row, start_column=1, end_row=bottom_row, end_column=n_cols)
    ref_cell = ws.cell(row=bottom_row, column=1)
    ref_cell.value = (
        "Tolerances — IRR: 0.03% (3 bps)  |  "
        "Equity: +/-0.50pp of 10%  |  "
        "DSCR bounds: 0.5x-5.0x  |  "
        "Max iterations: 8 outer x 6 inner"
    )
    ref_cell.font = Font(name="Aptos Narrow", size=9, italic=True, color="7A8291")
    ref_cell.alignment = _ALIGN_LEFT

    # Freeze panes (header row)
    ws.freeze_panes = "A5"

    # Print settings
    ws.sheet_properties.pageSetUpPr = openpyxl.worksheet.properties.PageSetupProperties(fitToPage=True)
    ws.page_setup.orientation = "landscape"
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 0

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    log.info("Generated summary XLSX (%d projects)", len(record.projects))
    return buf
