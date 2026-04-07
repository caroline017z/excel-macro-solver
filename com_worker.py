"""38DN COM Worker — Runs the VBA macro in an isolated subprocess.

Architecture (v2):
  1. Open workbook copy via DispatchEx
  2. Run SolveHeadless macro — it solves all projects, leaves calc MANUAL
  3. Read results per-project using targeted VBA recalc (no full-workbook recalc)
  4. Restore calc to Automatic only at close

Why this is fast:
  - SolveHeadless uses CalcModelCore (13 sheets, ~650K formulas) not
    Application.Calculate (~885K + output sheets)
  - Post-solve reads use SwitchProjectAndRecalc (same 13-sheet targeted recalc)
  - Calc stays manual throughout — no accidental full-workbook recalc bombs
"""
from __future__ import annotations

import contextlib
import json
import shutil
import sys
import tempfile
import time
from pathlib import Path


def _read_cell(wb: object, sheet: str, address: str) -> float | str | None:
    v = wb.Sheets(sheet).Range(address).Value
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return str(v)


def worker_main() -> None:
    """Entry point. Reads JSON from stdin, writes JSON to stdout."""
    try:
        _worker_inner()
    except Exception as exc:
        sys.stdout.write(json.dumps({
            "status": "error",
            "project_results": [],
            "duration_sec": 0.0,
            "saved_to": None,
            "error": f"{type(exc).__name__}: {exc}",
        }))
        sys.stdout.flush()


def _worker_inner() -> None:
    import pythoncom
    import win32com.client

    batch = json.loads(sys.stdin.read())
    workbook_path = Path(batch["workbook_path"])
    tasks = batch["tasks"]
    save_suffix = batch.get("saved_workbook_suffix", "_SOLVED")
    macro_names = batch.get("macro_names", [
        "SolveHeadless",                # Headless wrapper (no MsgBox, leaves calc manual)
        "SolveMinEquityWithHoldCo",     # Original macro (fallback)
    ])

    # Copy workbook to temp
    tmp_dir = Path(tempfile.mkdtemp(prefix="38dn_com_"))
    temp_path = tmp_dir / workbook_path.name
    shutil.copy2(str(workbook_path), str(temp_path))

    excel = None
    wb = None
    start = time.time()

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

        # Prefer Dispatch (reuses/creates single Excel process — faster
        # because VBA runs in-process without COM marshaling overhead).
        # Fall back to DispatchEx (separate process) if Dispatch fails.
        try:
            excel = win32com.client.Dispatch("Excel.Application")
        except Exception:
            excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        wb = excel.Workbooks.Open(
            str(temp_path),
            ReadOnly=False,
            UpdateLinks=0,
        )

        open_time = time.time() - start

        # --- Run the VBA macro ---
        macro_used = None
        macro_error = None

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

        solve_time = time.time() - start

        if macro_used is None:
            result["error"] = f"No macro found. Tried: {macro_names}"
        elif macro_error:
            result["error"] = f"Macro {macro_used} failed: {macro_error}"
        else:
            # --- Read results per project ---
            # SolveHeadless leaves calc in MANUAL mode, so we can read
            # without triggering full-workbook recalcs.
            # Use SwitchProjectAndRecalc for targeted 13-sheet recalc
            # when we need OFFSET-driven cells (F37, F31, etc.)

            has_switch_sub = macro_used == "SolveHeadless"

            project_results = []
            for task in tasks:
                col = task["project_col_letter"]
                read_cells = task.get("read_cells", [])

                # Switch F2 to this project with targeted recalc
                offset = task["project_offset"]
                if has_switch_sub:
                    try:
                        excel.Application.Run(
                            f"'{wb.Name}'!SwitchProjectAndRecalc",
                            offset,
                        )
                    except Exception:
                        # Fallback: set F2 directly, no recalc
                        idx_cell = task["project_index_cell"]
                        wb.Sheets(idx_cell["sheet"]).Range(idx_cell["address"]).Value = offset
                else:
                    idx_cell = task["project_index_cell"]
                    wb.Sheets(idx_cell["sheet"]).Range(idx_cell["address"]).Value = offset

                # Read cells
                solved: dict[str, float | str | None] = {}
                for cell in read_cells:
                    addr = cell["address"].replace("{col}", col)
                    key = f"{cell['sheet']}!{addr}"
                    solved[key] = _read_cell(wb, cell["sheet"], addr)

                project_results.append({
                    "project_name": task["project_name"],
                    "status": "converged",
                    "solved_values": solved,
                    "iterations_used": 0,
                    "duration_sec": 0,
                })

            read_time = time.time() - start

            # Restore original F2
            original_f2 = batch.get("original_f2", 1)
            if tasks and has_switch_sub:
                try:
                    excel.Application.Run(
                        f"'{wb.Name}'!SwitchProjectAndRecalc",
                        int(original_f2),
                    )
                except Exception:
                    pass
            elif tasks:
                idx_cell = tasks[0]["project_index_cell"]
                wb.Sheets(idx_cell["sheet"]).Range(idx_cell["address"]).Value = original_f2

            # Save solved workbook
            saved_to = None
            solved_name = workbook_path.stem + save_suffix + workbook_path.suffix
            solved_path = workbook_path.parent / solved_name
            with contextlib.suppress(Exception):
                wb.SaveAs(str(solved_path))
                saved_to = str(solved_path)

            all_ok = macro_error is None
            result = {
                "status": "converged" if all_ok else "not_converged",
                "project_results": project_results,
                "duration_sec": round(time.time() - start, 2),
                "saved_to": saved_to,
                "error": macro_error,
                "macro_used": macro_used,
                "open_time_sec": round(open_time, 2),
                "solve_time_sec": round(solve_time, 2),
                "read_time_sec": round(read_time - solve_time, 2),
            }

    except Exception as exc:
        result = {
            "status": "error",
            "project_results": [],
            "duration_sec": round(time.time() - start, 2),
            "saved_to": None,
            "error": f"{type(exc).__name__}: {exc}",
        }

    finally:
        # Restore calc mode to Automatic before closing
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

    sys.stdout.write(json.dumps(result, default=str))
    sys.stdout.flush()


if __name__ == "__main__":
    worker_main()
