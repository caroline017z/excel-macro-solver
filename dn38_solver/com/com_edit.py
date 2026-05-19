"""dn38_solver.com.com_edit — Excel COM helper for safe .xlsm mutation.

This is the CANONICAL way to mutate a .xlsm file in this codebase. Use it
for any change that needs to land in a macro-enabled workbook between
a macro re-import and a solver run (or before a solver run on a workbook
whose macro state we need to preserve cleanly).

WHY THIS EXISTS — the openpyxl .xlsm hazard
============================================
openpyxl with `keep_vba=True` reads + writes the macro blob (vbaProject
.bin) faithfully on a save round-trip, but it does NOT model every
Excel-specific XML part. openpyxl's save has been observed to strip or
rewrite:

  - data-validation extensions (x14:extLst inside xl/worksheets/*.xml)
  - conditional-formatting state
  - calculation chain caches (xl/calcChain.xml)
  - some doc-property parts

When any of those are rewritten between a macro re-import and the next
solver run, the embedded macro tolerates immediate cell writes (e.g.
F2 = project index) but throws a generic `Exception occurred.`
(HRESULT 0x80048028) inside `SolveOneProjectByColHL` — typically during
the first GoalSeek or full recalc. The macro itself is fine; the
WORKBOOK STATE is corrupted relative to what the macro assumes.

The verified fix is to fully rewrite the file via Excel COM SaveAs,
which re-serializes every part of the .xlsm package. That's what this
module does. Auto-recovery (`dn38_solver.com.auto_recovery`) builds on
the same principle: when a solver throws a recoverable COM error, the
recovery path runs `import_vba_module.py`, which itself calls
`wb.SaveAs(FileFormat=52)` and clears the corruption.

CONTRACT
========
- `edit_xlsm` opens via DispatchEx (fresh COM session, no zombie
  attachment), mutates the cells, SaveAs with FileFormat=52
  (xlOpenXMLWorkbookMacroEnabled), and quits.
- Never modifies the source file unless `dst` matches `src` (overwrite
  is explicit).
- Failure during any phase quits Excel cleanly before re-raising — no
  zombie EXCEL.EXE on error paths.

ANTI-PATTERNS — DO NOT
======================
- `openpyxl.load_workbook(xlsm, keep_vba=True); wb.save(xlsm)` between
  a macro re-import and a solver run. This is the exact pattern that
  caused RP Puma. The strip_sheets path in `direct_runner.py` is a
  legacy exception (startup-time optimization, mitigated by auto-
  recovery downstream) — do not add new sites.
- pywin32 `Dispatch` (vs `DispatchEx`). `Dispatch` reuses an existing
  Excel COM server if one is running, which is exactly what we don't
  want — a stale process state on the user's interactive Excel can
  poison our mutation.
- Setting cells via Excel COM but skipping the SaveAs. The corruption-
  clearing property is in the SaveAs, not in the cell write.
"""
from __future__ import annotations

import contextlib
import logging
from pathlib import Path
from typing import Iterable

log = logging.getLogger(__name__)

# Excel FileFormat constants. Pin them as named constants so call sites
# can read intent without needing to remember the magic numbers.
XL_OPEN_XML_MACRO_ENABLED = 52   # .xlsm
XL_OPEN_XML_WORKBOOK = 51        # .xlsx (used by Excel; not for macro-bearing files)


CellEdit = tuple[str, str, object]
"""(sheet_name, cell_address, value). value may be int/float/str/bool/None."""


def edit_xlsm(
    src: Path | str,
    edits: Iterable[CellEdit],
    *,
    dst: Path | str | None = None,
    update_links: int = 0,
) -> Path:
    """Apply a batch of cell edits to a .xlsm file via Excel COM.

    Opens `src`, writes each (sheet, address, value), SaveAs to `dst`
    (defaults to overwriting `src`), quits cleanly. Returns the path
    that was written.

    Use this for any mid-workflow mutation of an .xlsm where:
      - the embedded macro will run after the mutation, OR
      - a downstream consumer reads workbook state beyond raw cell
        values (data validation, conditional formatting, calc cache).

    Args:
        src: Source .xlsm to open.
        edits: Iterable of (sheet, address, value). Sheet must exist;
               address is A1-style ("F2", "L7", "AB123"). Value None
               clears the cell.
        dst:   Destination path; defaults to overwriting `src`. Both
               .xlsm; SaveAs always uses xlOpenXMLWorkbookMacroEnabled.
        update_links: Excel UpdateLinks param. 0 = don't update on
               open; matches the convention in direct_runner.

    Raises:
        FileNotFoundError: src does not exist.
        RuntimeError:      a sheet name in edits is not in the workbook.
        Any pywin32 com_error from Excel itself bubbles up.
    """
    import pythoncom
    import win32com.client

    src_p = Path(src)
    if not src_p.exists():
        raise FileNotFoundError(f"source workbook not found: {src_p}")
    dst_p = Path(dst) if dst is not None else src_p

    edits_list = list(edits)
    if not edits_list:
        log.debug("com_edit.edit_xlsm: no edits — short-circuiting")
        return dst_p

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        wb = excel.Workbooks.Open(
            str(src_p),
            ReadOnly=False,
            UpdateLinks=update_links,
        )

        # Cache sheet handles so a multi-edit batch on the same sheet
        # doesn't pay N COM lookups.
        sheet_cache: dict[str, object] = {}
        for sheet_name, address, value in edits_list:
            ws = sheet_cache.get(sheet_name)
            if ws is None:
                try:
                    ws = wb.Sheets(sheet_name)
                except Exception as exc:
                    raise RuntimeError(
                        f"sheet not found in workbook: {sheet_name!r}"
                    ) from exc
                sheet_cache[sheet_name] = ws
            ws.Range(address).Value = value

        # SaveAs with explicit FileFormat=52 is the load-bearing step —
        # it's what fully re-serializes every package part and clears
        # any prior corruption. Plain wb.Save() is NOT a substitute.
        wb.SaveAs(str(dst_p), FileFormat=XL_OPEN_XML_MACRO_ENABLED)
        log.debug(
            "com_edit.edit_xlsm: wrote %d cell(s) → %s",
            len(edits_list),
            dst_p.name,
        )
        return dst_p
    finally:
        if wb is not None:
            with contextlib.suppress(Exception):
                wb.Close(SaveChanges=False)
        if excel is not None:
            with contextlib.suppress(Exception):
                excel.Quit()
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()
