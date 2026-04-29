"""dn38_solver.convert — Canonical value conversion utilities.

ONE place for safe_float, safe_value, col_letter. No duplication.
Every module in the package imports from here.
"""
from __future__ import annotations

import contextlib

from openpyxl.utils import get_column_letter as _openpyxl_col_letter


def safe_float(v: object) -> float | None:
    """Convert any value to float, returning None on failure.

    NEVER returns str. This is the canonical version — replaces the 3
    inconsistent implementations that existed before.
    """
    if v is None:
        return None
    with contextlib.suppress(ValueError, TypeError):
        return float(v)
    return None


def safe_value(v: object) -> float | str | None:
    """Normalize a cell value to float (preferred) or str.

    For snapshot/diff use where we want to preserve text labels.
    Returns None only for None input.
    """
    if v is None:
        return None
    with contextlib.suppress(ValueError, TypeError):
        return float(v)
    return str(v)


def safe_str_or_float(v: object) -> float | str | None:
    """Like safe_value, but also collapses empty strings to None.

    For COM telemetry reads where blank cells surface as "" rather than None.
    """
    if v is None or v == "":
        return None
    with contextlib.suppress(ValueError, TypeError):
        return float(v)
    return str(v)


def col_letter(col_num: int) -> str:
    """1-based column number to Excel letter(s). Delegates to openpyxl."""
    return _openpyxl_col_letter(col_num)
