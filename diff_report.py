"""
38DN Excel Macro Runner — Pre/Post Macro Diff Reporting

Captures workbook state before and after a macro run, computes
cell-level diffs, and produces console + JSON reports.
"""
import json
from datetime import datetime
from pathlib import Path

from config import (
    OUTPUT_ROWS, PROJECT_COL_END, PROJECT_COL_START,
    PROJECT_NAME_ROW, PROJECT_TOGGLE_ROW,
)


def _safe_value(v):
    """Normalize a COM cell value to a JSON-friendly Python type."""
    if v is None:
        return None
    try:
        f = float(v)
        return f
    except (ValueError, TypeError):
        return str(v)


def capture_state(wb, config=None):
    """
    Read key output cells from the workbook and return a dict keyed by
    (sheet_name, row, col) -> value.

    Captures:
      - "Project Inputs": OUTPUT_ROWS for each active project column
      - "Dashboard": rows 1-50, columns A-J (1-10)
      - "NPP Calc": rows 1-30, columns A-F (1-6)
    """
    state = {}

    # --- Project Inputs sheet: OUTPUT_ROWS for active project columns ---
    try:
        ws_pi = wb.Sheets("Project Inputs")
        for col in range(PROJECT_COL_START, PROJECT_COL_END + 1):
            name = ws_pi.Cells(PROJECT_NAME_ROW, col).Value
            if not name or not str(name).strip():
                continue
            toggle = ws_pi.Cells(PROJECT_TOGGLE_ROW, col).Value
            is_on = (
                str(toggle).strip().lower() in ("1", "on", "true")
                if toggle
                else False
            )
            if not is_on:
                continue
            for row in OUTPUT_ROWS:
                state[("Project Inputs", row, col)] = _safe_value(
                    ws_pi.Cells(row, col).Value
                )
    except Exception:
        pass

    # --- Dashboard sheet: rows 1-50, columns A-J (1-10) ---
    try:
        ws_dash = wb.Sheets("Dashboard")
        for r in range(1, 51):
            for c in range(1, 11):
                v = ws_dash.Cells(r, c).Value
                if v is not None:
                    state[("Dashboard", r, c)] = _safe_value(v)
    except Exception:
        pass

    # --- NPP Calc sheet: rows 1-30, columns A-F (1-6) ---
    try:
        ws_npp = wb.Sheets("NPP Calc")
        for r in range(1, 31):
            for c in range(1, 7):
                v = ws_npp.Cells(r, c).Value
                if v is not None:
                    state[("NPP Calc", r, c)] = _safe_value(v)
    except Exception:
        pass

    return state


def _col_letter(col_num):
    """Convert 1-based column number to Excel letter(s), e.g. 1->A, 27->AA."""
    result = ""
    while col_num > 0:
        col_num, remainder = divmod(col_num - 1, 26)
        result = chr(65 + remainder) + result
    return result


def _cell_label(sheet, row, col):
    """Human-readable label for a cell, including OUTPUT_ROWS label if applicable."""
    col_letter = _col_letter(col)
    base = f"{sheet}!{col_letter}{row}"
    if sheet == "Project Inputs" and row in OUTPUT_ROWS:
        return f"{base} ({OUTPUT_ROWS[row]})"
    return base


def compute_diff(before, after):
    """
    Compare before/after state dicts and return a list of change records.

    Each record: {sheet, row, col, label, before, after, delta, pct_change}
    """
    all_keys = set(before.keys()) | set(after.keys())
    changes = []

    for key in sorted(all_keys):
        bv = before.get(key)
        av = after.get(key)

        if bv == av:
            continue

        sheet, row, col = key
        label = _cell_label(sheet, row, col)

        delta = None
        pct_change = None
        if isinstance(bv, (int, float)) and isinstance(av, (int, float)):
            delta = av - bv
            if bv != 0:
                pct_change = (delta / abs(bv)) * 100

        changes.append({
            "sheet": sheet,
            "row": row,
            "col": col,
            "label": label,
            "before": bv,
            "after": av,
            "delta": delta,
            "pct_change": pct_change,
        })

    return changes


def print_diff_report(changes, workbook_name):
    """Print a formatted console report of cell-level changes."""
    print(f"\n{'='*72}")
    print(f"  DIFF REPORT: {workbook_name}")
    print(f"  {len(changes)} cell(s) changed")
    print(f"{'='*72}")

    if not changes:
        print("  No changes detected.")
        print(f"{'='*72}\n")
        return

    # Group by sheet
    by_sheet = {}
    for c in changes:
        by_sheet.setdefault(c["sheet"], []).append(c)

    for sheet, sheet_changes in by_sheet.items():
        print(f"\n  --- {sheet} ({len(sheet_changes)} changes) ---")
        print(f"  {'Cell':<40} {'Before':>14} {'After':>14} {'Delta':>14} {'%':>8}")
        print(f"  {'-'*40} {'-'*14} {'-'*14} {'-'*14} {'-'*8}")

        for c in sheet_changes:
            bv_str = _fmt_val(c["before"])
            av_str = _fmt_val(c["after"])
            d_str = _fmt_val(c["delta"]) if c["delta"] is not None else ""
            pct_str = (
                f"{c['pct_change']:+.1f}%"
                if c["pct_change"] is not None
                else ""
            )
            print(f"  {c['label']:<40} {bv_str:>14} {av_str:>14} {d_str:>14} {pct_str:>8}")

    print(f"\n{'='*72}\n")


def _fmt_val(v):
    """Format a value for display."""
    if v is None:
        return "—"
    if isinstance(v, float):
        if abs(v) >= 1_000_000:
            return f"{v:,.0f}"
        if abs(v) >= 1:
            return f"{v:,.4f}"
        return f"{v:.6f}"
    return str(v)


def save_diff_to_json(changes, output_path):
    """Save the diff as a JSON file."""
    output_path = Path(output_path)
    payload = {
        "generated_at": datetime.now().isoformat(),
        "total_changes": len(changes),
        "changes": changes,
    }
    with open(output_path, "w", encoding="utf-8") as f:
        json.dump(payload, f, indent=2, default=str)
    print(f"  Diff JSON saved: {output_path.name}")
