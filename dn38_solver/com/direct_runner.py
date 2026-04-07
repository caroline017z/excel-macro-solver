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

from dn38_solver.types import CellAddress, SolveTask

log = logging.getLogger(__name__)

STATUS_FILE = Path(__file__).resolve().parent.parent.parent / "solver_status.json"
SOLVER_RESULTS_SHEET = "__SolverResults"


def _write_status(data: dict) -> None:
    """Write solver status to JSON for the Streamlit tracker."""
    try:
        STATUS_FILE.write_text(json.dumps(data, default=str))
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

        # Build project list for status tracking
        proj_names = [
            (t.project_name if hasattr(t, "project_name") else t.get("project_name", "?"))
            for t in tasks
        ]
        _write_status({
            "phase": "opening",
            "workbook": workbook_path,
            "total_projects": len(tasks),
            "projects": [{"name": n, "status": "pending"} for n in proj_names],
            "elapsed_sec": time.time() - start,
        })

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
        _write_status({
            "phase": "solving",
            "workbook": workbook_path,
            "total_projects": len(tasks),
            "projects": [{"name": n, "status": "solving"} for n in proj_names],
            "elapsed_sec": time.time() - start,
            "macro_used": "SolveHeadless",
        })

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
        _write_status({
            "phase": "reading",
            "workbook": workbook_path,
            "total_projects": len(tasks),
            "projects": [{"name": n, "status": "reading"} for n in proj_names],
            "elapsed_sec": time.time() - start,
            "macro_used": macro_used,
            "macro_time_sec": solve_time,
        })
        t0 = time.time()
        project_results = []
        solver_results, heartbeat = _read_solver_results_map(wb)

        for task_dict in tasks:
            # Handle both SolveTask objects and dicts
            if hasattr(task_dict, "project_col_letter"):
                col = task_dict.project_col_letter
                offset = task_dict.project_offset
                name = task_dict.project_name
                read_cells = task_dict.read_cells
                idx_cell = task_dict.project_index_cell
            else:
                col = task_dict["project_col_letter"]
                offset = task_dict["project_offset"]
                name = task_dict["project_name"]
                read_cells = task_dict.get("read_cells", [])
                idx_cell = task_dict["project_index_cell"]

            # Switch F2 with targeted recalc
            if has_switch:
                try:
                    excel.Application.Run(
                        f"'{wb.Name}'!SwitchProjectAndRecalc",
                        int(offset),
                    )
                except Exception:
                    _set_f2(wb, idx_cell, offset)
            else:
                _set_f2(wb, idx_cell, offset)

            # Read cells
            solved: dict[str, float | str | None] = {}
            for cell in read_cells:
                if hasattr(cell, "address"):
                    addr = cell.address.replace("{col}", col)
                    sheet = cell.sheet
                else:
                    addr = cell["address"].replace("{col}", col)
                    sheet = cell["sheet"]
                key = f"{sheet}!{addr}"
                solved[key] = _read_cell(wb, sheet, addr)

            # Prefer per-project DSCR captured during solve loop to avoid
            # last-project F129 bleed in multi-project runs.
            dscr_key = "PT Returns!F129"
            meta = solver_results.get(int(offset))
            if meta and "dscr" in meta:
                solved[dscr_key] = meta["dscr"]

            project_results.append({
                "project_name": name,
                "status": "converged",
                "solved_values": solved,
                "iterations_used": 0,
                "duration_sec": 0,
                "meta": meta or {},
            })

        read_time = time.time() - t0
        log.info("  Read %d project(s) in %.1fs", len(tasks), read_time)

        # Restore original F2
        if tasks and has_switch:
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

        total = time.time() - start
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
        }

        # Write completion status for Streamlit tracker
        from dn38_solver.convert import safe_float
        tracker_projects = []
        for pr in project_results:
            sv = pr.get("solved_values", {})
            tracker_projects.append({
                "name": pr.get("project_name", "?"),
                "status": "converged",
                "npp": safe_float(sv.get(next((k for k in sv if k.endswith("38")), ""), None)),
                "dev_fee": safe_float(sv.get(next((k for k in sv if k.endswith("32")), ""), None)),
                "fmv": safe_float(sv.get(next((k for k in sv if k.endswith("33")), ""), None)),
            })
        _write_status({
            "phase": "complete",
            "workbook": workbook_path,
            "total_projects": len(tasks),
            "projects": tracker_projects,
            "elapsed_sec": total,
            "total_time_sec": total,
            "macro_used": macro_used,
            "open_time_sec": round(open_time, 2),
            "macro_time_sec": round(solve_time, 2),
            "read_time_sec": round(read_time, 2),
            "solver_heartbeat": heartbeat,
            "error": None,
        })

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
    v = wb.Sheets(sheet).Range(address).Value
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return str(v)


def _set_f2(wb: object, idx_cell: object, offset: int) -> None:
    """Set F2 directly (fallback when SwitchProjectAndRecalc unavailable)."""
    if hasattr(idx_cell, "sheet"):
        wb.Sheets(idx_cell.sheet).Range(idx_cell.address).Value = offset
    else:
        wb.Sheets(idx_cell["sheet"]).Range(idx_cell["address"]).Value = offset


def _read_solver_results_map(
    wb: object,
) -> tuple[dict[int, dict[str, float | str | None]], str | None]:
    """Read per-project solve telemetry captured by SolveHeadless VBA."""
    out: dict[int, dict[str, float | str | None]] = {}
    try:
        ws = wb.Sheets(SOLVER_RESULTS_SHEET)
    except Exception:
        return out, None

    heartbeat = _to_safe(ws.Range("N1").Value)
    if not isinstance(heartbeat, str):
        heartbeat = None

    row = 2
    while True:
        offset_val = ws.Range(f"A{row}").Value
        if offset_val in (None, ""):
            break
        try:
            offset = int(offset_val)
        except (ValueError, TypeError):
            row += 1
            continue
        out[offset] = {
            "project_name": _to_safe(ws.Range(f"B{row}").Value),
            "dscr": _to_float(ws.Range(f"C{row}").Value),
            "npp": _to_float(ws.Range(f"D{row}").Value),
            "dev_fee": _to_float(ws.Range(f"E{row}").Value),
            "equity_pct": _to_float(ws.Range(f"F{row}").Value),
            "irr_gap": _to_float(ws.Range(f"G{row}").Value),
            "appr_gap": _to_float(ws.Range(f"H{row}").Value),
            "converged_flag": _to_safe(ws.Range(f"I{row}").Value),
            "calc_tier": _to_safe(ws.Range(f"J{row}").Value),
            "gs_retry_limit": _to_safe(ws.Range(f"K{row}").Value),
            "mode": _to_safe(ws.Range(f"L{row}").Value),
            "solve_seconds": _to_float(ws.Range(f"M{row}").Value),
            "heartbeat": _to_safe(ws.Range(f"N{row}").Value),
        }
        row += 1
    return out, heartbeat


def _to_float(v: object) -> float | None:
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def _to_safe(v: object) -> str | float | None:
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return str(v)
