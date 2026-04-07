"""dn38_solver.reporting.diff — Pre/post cell-level comparison.

Computes which cells changed between two snapshots (before/after solve).
Returns typed CellChange structs, not raw dicts.
"""
from __future__ import annotations

import json
import logging
from pathlib import Path

import msgspec

from dn38_solver.types import CellChange
from dn38_solver.config import OUTPUT_ROWS
from dn38_solver.convert import col_letter, safe_float

log = logging.getLogger(__name__)

SnapshotT = dict[tuple[str, int, int], float | str | None]


def compute_diff(
    before: SnapshotT,
    after: SnapshotT,
) -> tuple[CellChange, ...]:
    """Compare two cell snapshots and return changes."""
    all_keys = set(before) | set(after)
    changes: list[CellChange] = []

    for key in sorted(all_keys):
        b = before.get(key)
        a = after.get(key)

        if b == a:
            continue

        sheet, row, col = key
        label = OUTPUT_ROWS.get(row, f"R{row}C{col_letter(col)}")

        bf = safe_float(b)
        af = safe_float(a)
        delta = af - bf if bf is not None and af is not None else None
        pct = delta / abs(bf) if delta is not None and bf and bf != 0 else None

        changes.append(CellChange(
            sheet=sheet,
            row=row,
            col=col,
            label=label,
            before=b,
            after=a,
            delta=delta,
            pct_change=pct,
        ))

    return tuple(changes)


def print_diff_report(
    changes: tuple[CellChange, ...],
    workbook_name: str,
) -> None:
    """Print a formatted diff report to the log."""
    if not changes:
        log.info("No changes detected in %s", workbook_name)
        return

    log.info("Diff report for %s (%d changes):", workbook_name, len(changes))
    log.info("  %-20s %-6s %-6s %-15s %-15s %-12s %-10s", "Sheet", "Row", "Col", "Before", "After", "Delta", "% Change")
    log.info("  %s", "-" * 85)

    for c in changes:
        before_s = f"{c.before:.4f}" if isinstance(c.before, float) else str(c.before or "—")
        after_s = f"{c.after:.4f}" if isinstance(c.after, float) else str(c.after or "—")
        delta_s = f"{c.delta:+.4f}" if c.delta is not None else "—"
        pct_s = f"{c.pct_change:+.2%}" if c.pct_change is not None else "—"
        log.info("  %-20s %-6d %-6s %-15s %-15s %-12s %-10s",
                 c.sheet, c.row, col_letter(c.col), before_s, after_s, delta_s, pct_s)


def save_diff_json(
    changes: tuple[CellChange, ...],
    output_path: Path,
) -> None:
    """Save the diff as a JSON file."""
    data = msgspec.json.encode(changes)
    output_path.write_bytes(data)
    log.info("Diff saved to %s (%d changes)", output_path, len(changes))
