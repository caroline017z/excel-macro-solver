"""dn38_solver.com.direct_runner — Direct COM execution (no subprocess).

Runs the VBA macro in-process via COM. Since SolveHeadless handles all
GoalSeek logic in VBA, there's no need for subprocess isolation.

Performance gains over subprocess approach:
  - No process spawn / JSON marshal overhead
  - Single COM connection reused for macro + result reads
  - Warm-up calculate primes the formula dependency graph
"""
from __future__ import annotations

import contextlib
import json
import logging
import shutil
import tempfile
import time
from pathlib import Path
from typing import NamedTuple

from dn38_solver.convert import safe_float, safe_str_or_float, safe_value
from dn38_solver.shadow.validation import scan_workbook_errors
from dn38_solver.types import CellAddress, SolveTask

log = logging.getLogger(__name__)

STATUS_FILE = Path(__file__).resolve().parent.parent.parent / "solver_status.json"
SOLVER_RESULTS_SHEET = "__SolverResults"

# Upper bound on rows to read from __SolverResults in one bulk Range.Value call.
# 60 active projects + slack for header growth; one COM round-trip beats N+1.
_RESULTS_BULK_ROWS = 200


class _NormCell(NamedTuple):
    sheet: str
    address: str  # may contain "{col}" template placeholder


class _NormTask(NamedTuple):
    col: str
    offset: int
    name: str
    idx_cell: CellAddress | dict
    read_cells: tuple[_NormCell, ...]


def _norm_cell(c: object) -> _NormCell:
    if hasattr(c, "address"):
        return _NormCell(sheet=c.sheet, address=c.address)
    return _NormCell(sheet=c["sheet"], address=c["address"])


def _norm_task(t: object) -> _NormTask:
    if hasattr(t, "project_col_letter"):
        cells = tuple(_norm_cell(c) for c in t.read_cells)
        return _NormTask(
            col=t.project_col_letter,
            offset=int(t.project_offset),
            name=t.project_name,
            idx_cell=t.project_index_cell,
            read_cells=cells,
        )
    cells = tuple(_norm_cell(c) for c in t.get("read_cells", []))
    return _NormTask(
        col=t["project_col_letter"],
        offset=int(t["project_offset"]),
        name=t["project_name"],
        idx_cell=t["project_index_cell"],
        read_cells=cells,
    )


class _StatusWriter:
    """Buffered solver-status writer.

    Holds the immutable run-level fields (workbook, project list, start time)
    so each phase update only constructs a small delta dict before writing.
    Output is still JSON to STATUS_FILE for the Streamlit tracker.
    """

    __slots__ = ("_base", "_start", "_path")

    def __init__(
        self,
        *,
        workbook_path: str,
        proj_names: list[str],
        start: float,
        path: Path = STATUS_FILE,
    ) -> None:
        self._base = {
            "workbook": workbook_path,
            "total_projects": len(proj_names),
            "_proj_names": proj_names,
        }
        self._start = start
        self._path = path

    def update(
        self,
        phase: str,
        *,
        per_project_status: str | None = None,
        projects: list[dict] | None = None,
        **extras: object,
    ) -> None:
        if projects is None and per_project_status is not None:
            projects = [
                {"name": n, "status": per_project_status}
                for n in self._base["_proj_names"]
            ]
        payload = {
            "phase": phase,
            "workbook": self._base["workbook"],
            "total_projects": self._base["total_projects"],
            "projects": projects or [],
            "elapsed_sec": time.time() - self._start,
        }
        payload.update(extras)
        try:
            self._path.write_text(json.dumps(payload, default=str))
        except OSError:
            pass


def run_direct(
    workbook_path: str,
    tasks: list[SolveTask],
    *,
    original_f2: int = 1,
    timeout_sec: int = 600,
) -> dict:
    """Open Excel, run SolveHeadless, read results, close. All in-process."""
    import pythoncom
    import win32com.client

    start = time.time()
    tmp_dir = Path(tempfile.mkdtemp(prefix="38dn_com_"))
    temp_path = tmp_dir / Path(workbook_path).name
    shutil.copy2(workbook_path, str(temp_path))

    excel = None
    wb = None

    result: dict = {
        "status": "error",
        "project_results": [],
        "duration_sec": 0.0,
        "saved_to": None,
        "error": None,
        "macro_used": None,
    }

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.Dispatch("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        # Normalize task payloads once so the per-project loop is branch-free.
        norm_tasks = [_norm_task(t) for t in tasks]
        proj_names = [nt.name for nt in norm_tasks]

        status = _StatusWriter(
            workbook_path=workbook_path,
            proj_names=proj_names,
            start=start,
        )
        status.update("opening", per_project_status="pending")

        log.info("  Opening workbook via COM...")
        wb = excel.Workbooks.Open(
            str(temp_path),
            ReadOnly=False,
            UpdateLinks=0,
        )
        open_time = time.time() - start
        log.info("  Opened in %.1fs", open_time)

        warmup_time = 0.0

        # --- Run the VBA macro ---
        macro_names = [
            "SolveHeadless",
            "SolveMinEquityWithHoldCo",
        ]

        macro_used = None
        macro_error = None

        log.info("  Running macro...")
        status.update(
            "solving",
            per_project_status="solving",
            macro_used="SolveHeadless",
        )

        t0 = time.time()
        for macro_name in macro_names:
            try:
                excel.Application.Run(f"'{wb.Name}'!{macro_name}")
                macro_used = macro_name
                break
            except Exception as e:
                err_str = str(e).lower()
                if "macro may not be available" in err_str or "cannot run" in err_str:
                    continue
                macro_error = str(e)
                macro_used = macro_name
                break

        solve_time = time.time() - t0
        log.info("  Macro '%s' completed in %.1fs", macro_used, solve_time)
        if timeout_sec > 0 and solve_time > timeout_sec:
            result["error"] = (
                f"Macro execution exceeded timeout_sec={timeout_sec} "
                f"(actual={solve_time:.1f}s)"
            )
            return result

        if macro_used is None:
            result["error"] = f"No macro found. Tried: {macro_names}"
            return result
        if macro_error:
            result["error"] = f"Macro {macro_used} failed: {macro_error}"
            return result

        # --- Read results per project ---
        # SolveHeadless leaves calc in MANUAL mode.
        # Use SwitchProjectAndRecalc for targeted 13-sheet recalc.
        has_switch = macro_used == "SolveHeadless"

        log.info("  Reading results for %d project(s)...", len(tasks))
        status.update(
            "reading",
            per_project_status="reading",
            macro_used=macro_used,
            macro_time_sec=solve_time,
        )
        t0 = time.time()
        project_results = []
        solver_results, heartbeat = _read_solver_results_map(wb)

        for nt in norm_tasks:
            # Switch F2 with targeted recalc
            if has_switch:
                try:
                    excel.Application.Run(
                        f"'{wb.Name}'!SwitchProjectAndRecalc",
                        nt.offset,
                    )
                except Exception:
                    _set_f2(wb, nt.idx_cell, nt.offset)
            else:
                _set_f2(wb, nt.idx_cell, nt.offset)

            # Read cells
            solved: dict[str, float | str | None] = {}
            npp = dev_fee = fmv = None
            for cell in nt.read_cells:
                addr = cell.address.replace("{col}", nt.col)
                key = f"{cell.sheet}!{addr}"
                val = _read_cell(wb, cell.sheet, addr)
                solved[key] = val
                # Capture summary scalars at read time so we don't re-scan
                # solved_values later for the tracker payload.
                if addr.endswith("38"):
                    npp = safe_float(val)
                elif addr.endswith("32"):
                    dev_fee = safe_float(val)
                elif addr.endswith("33"):
                    fmv = safe_float(val)

            # Prefer per-project DSCR captured during solve loop to avoid
            # last-project F129 bleed in multi-project runs.
            dscr_key = "PT Returns!F129"
            meta = solver_results.get(nt.offset)
            if meta and "dscr" in meta:
                solved[dscr_key] = meta["dscr"]

            project_results.append({
                "project_name": nt.name,
                "status": "converged",
                "solved_values": solved,
                "iterations_used": 0,
                "duration_sec": 0,
                "meta": meta or {},
                "_summary": {"npp": npp, "dev_fee": dev_fee, "fmv": fmv},
            })

        read_time = time.time() - t0
        log.info("  Read %d project(s) in %.1fs", len(tasks), read_time)

        # Restore original F2
        if norm_tasks and has_switch:
            with contextlib.suppress(Exception):
                excel.Application.Run(
                    f"'{wb.Name}'!SwitchProjectAndRecalc",
                    int(original_f2),
                )

        # Save solved workbook
        saved_to = None
        wb_path = Path(workbook_path)
        solved_name = wb_path.stem + "_SOLVED" + wb_path.suffix
        solved_path = wb_path.parent / solved_name
        with contextlib.suppress(Exception):
            wb.SaveAs(str(solved_path))
            saved_to = str(solved_path)

        # Post-export formula-error gate: scan the just-saved file for
        # cached Excel error tokens (#REF! / #DIV/0! / #VALUE! / etc.).
        # Pure-Python via openpyxl — no LibreOffice dependency.
        validation = None
        if saved_to is not None:
            with contextlib.suppress(Exception):
                validation = scan_workbook_errors(saved_to)

        total = time.time() - start

        # Build tracker payload from pre-computed per-project summaries.
        tracker_projects = [
            {
                "name": pr["project_name"],
                "status": "converged",
                **pr["_summary"],
            }
            for pr in project_results
        ]
        # _summary was a transport-only field for the tracker — drop it from
        # the returned project_results so downstream consumers see a clean shape.
        for pr in project_results:
            pr.pop("_summary", None)

        result = {
            "status": "converged",
            "project_results": project_results,
            "duration_sec": round(total, 2),
            "saved_to": saved_to,
            "error": None,
            "macro_used": macro_used,
            "open_time_sec": round(open_time, 2),
            "warmup_time_sec": round(warmup_time, 2),
            "solve_time_sec": round(solve_time, 2),
            "read_time_sec": round(read_time, 2),
            "solver_heartbeat": heartbeat,
            "validation": validation,
        }

        status.update(
            "complete",
            projects=tracker_projects,
            total_time_sec=total,
            macro_used=macro_used,
            open_time_sec=round(open_time, 2),
            macro_time_sec=round(solve_time, 2),
            read_time_sec=round(read_time, 2),
            solver_heartbeat=heartbeat,
            error=None,
        )

    except Exception as exc:
        result = {
            "status": "error",
            "project_results": [],
            "duration_sec": round(time.time() - start, 2),
            "saved_to": None,
            "error": f"{type(exc).__name__}: {exc}",
        }

    finally:
        if excel is not None:
            with contextlib.suppress(Exception):
                excel.Calculation = -4105  # xlCalculationAutomatic
        if wb is not None:
            with contextlib.suppress(Exception):
                wb.Close(SaveChanges=False)
        if excel is not None:
            with contextlib.suppress(Exception):
                excel.ScreenUpdating = True
            with contextlib.suppress(Exception):
                excel.EnableEvents = True
            with contextlib.suppress(Exception):
                excel.Quit()
        with contextlib.suppress(Exception):
            shutil.rmtree(tmp_dir, ignore_errors=True)
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()

    return result


def _read_cell(wb: object, sheet: str, address: str) -> float | str | None:
    return safe_value(wb.Sheets(sheet).Range(address).Value)


def _set_f2(wb: object, idx_cell: object, offset: int) -> None:
    """Set F2 directly (fallback when SwitchProjectAndRecalc unavailable)."""
    if hasattr(idx_cell, "sheet"):
        wb.Sheets(idx_cell.sheet).Range(idx_cell.address).Value = offset
    else:
        wb.Sheets(idx_cell["sheet"]).Range(idx_cell["address"]).Value = offset


def _read_solver_results_map(
    wb: object,
) -> tuple[dict[int, dict[str, float | str | None]], str | None]:
    """Read per-project solve telemetry captured by SolveHeadless VBA.

    One bulk Range.Value read covering A2:N{2 + _RESULTS_BULK_ROWS - 1} replaces
    the prior per-cell while loop (~14 COM round-trips per project, ~840 for a
    60-project portfolio). The block is sparse-tolerant: rows whose A column is
    blank are treated as end-of-data.
    """
    out: dict[int, dict[str, float | str | None]] = {}
    try:
        ws = wb.Sheets(SOLVER_RESULTS_SHEET)
    except Exception:
        return out, None

    heartbeat = safe_str_or_float(ws.Range("N1").Value)
    if not isinstance(heartbeat, str):
        heartbeat = None

    last_row = 1 + _RESULTS_BULK_ROWS
    try:
        block = ws.Range(f"A2:N{last_row}").Value
    except Exception:
        return out, heartbeat
    if block is None:
        return out, heartbeat

    # Single-row Range.Value comes back as a flat tuple; multi-row as
    # tuple-of-tuples. Normalize to the latter.
    if block and not isinstance(block[0], tuple):
        block = (block,)

    for row_vals in block:
        offset_raw = row_vals[0]
        if offset_raw is None or offset_raw == "":
            break
        try:
            offset = int(offset_raw)
        except (ValueError, TypeError):
            continue
        out[offset] = {
            "project_name": safe_str_or_float(row_vals[1]),
            "dscr": safe_float(row_vals[2]),
            "npp": safe_float(row_vals[3]),
            "dev_fee": safe_float(row_vals[4]),
            "equity_pct": safe_float(row_vals[5]),
            "irr_gap": safe_float(row_vals[6]),
            "appr_gap": safe_float(row_vals[7]),
            "converged_flag": safe_str_or_float(row_vals[8]),
            "calc_tier": safe_str_or_float(row_vals[9]),
            "gs_retry_limit": safe_str_or_float(row_vals[10]),
            "mode": safe_str_or_float(row_vals[11]),
            "solve_seconds": safe_float(row_vals[12]),
            "heartbeat": safe_str_or_float(row_vals[13]),
        }
    return out, heartbeat
