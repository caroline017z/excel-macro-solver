"""
38DN Excel Macro Runner — Process-Isolated COM Executor
Runs Excel COM automation in a completely separate subprocess so it
CANNOT interfere with your active Excel session.

Key safety guarantees:
1. DispatchEx creates a NEW Excel.exe process (not your open session)
2. Runs on a TEMP COPY of the workbook (no file locks on original)
3. Subprocess isolation: if the COM code crashes, your Python session is unaffected
4. Aggressive cleanup: taskkill fallback if Excel process hangs
5. Timeout protection: auto-kills after max_seconds
"""
import json
import os
import shutil
import subprocess
import sys
import tempfile
import time
from pathlib import Path

from config import (
    DEFAULT_WORKBOOK, MACRO_VARIANTS, OUTPUT_ROWS,
    PROJECT_COL_END, PROJECT_COL_START, PROJECT_NAME_ROW,
    PROJECT_TOGGLE_ROW, SNAPSHOT_SHEETS,
)


# This script can be run directly as a subprocess OR imported
# When run as subprocess, it receives args via JSON on stdin

def _run_in_subprocess(workbook_path: str, macro_variants: list,
                       output_rows: dict, dry_run: bool = False,
                       max_seconds: int = 300) -> dict:
    """
    Launch the COM automation in a completely separate Python process.
    Returns results dict. The parent process is never exposed to COM.
    """
    # Serialize the task as JSON
    task = {
        "workbook_path": str(workbook_path),
        "macro_variants": macro_variants,
        "output_rows": {str(k): v for k, v in output_rows.items()},
        "dry_run": dry_run,
        "project_col_start": PROJECT_COL_START,
        "project_col_end": PROJECT_COL_END,
        "project_name_row": PROJECT_NAME_ROW,
        "project_toggle_row": PROJECT_TOGGLE_ROW,
        "snapshot_sheets": SNAPSHOT_SHEETS,
    }

    # Run THIS file as a subprocess with the task JSON piped in
    proc = subprocess.Popen(
        [sys.executable, __file__, "--worker"],
        stdin=subprocess.PIPE,
        stdout=subprocess.PIPE,
        stderr=subprocess.PIPE,
        text=True,
        creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,  # Full process isolation on Windows
    )

    try:
        stdout, stderr = proc.communicate(
            input=json.dumps(task),
            timeout=max_seconds,
        )
    except subprocess.TimeoutExpired:
        proc.kill()
        # Also kill any orphaned Excel processes from this subprocess
        _kill_orphan_excel()
        return {"status": "error", "error": f"Timed out after {max_seconds}s",
                "projects": []}

    if proc.returncode != 0:
        return {"status": "error", "error": stderr.strip() or f"Exit code {proc.returncode}",
                "projects": []}

    try:
        return json.loads(stdout)
    except json.JSONDecodeError:
        return {"status": "error", "error": f"Invalid JSON output: {stdout[:200]}",
                "projects": []}


def _kill_orphan_excel():
    """Kill Excel processes that were started by this runner (hidden, no window)."""
    try:
        # Only kill Excel processes that are hidden (no main window)
        # This is a safety measure — we never kill the user's visible Excel
        result = subprocess.run(
            ["powershell", "-Command",
             "Get-Process excel -ErrorAction SilentlyContinue | "
             "Where-Object { $_.MainWindowHandle -eq 0 } | "
             "Stop-Process -Force"],
            capture_output=True, timeout=10,
        )
    except Exception:
        pass


def _worker_main():
    """
    This runs INSIDE the subprocess. It's the only code that touches COM.
    Reads task JSON from stdin, runs Excel COM, writes results JSON to stdout.
    """
    import pythoncom
    import win32com.client

    task = json.loads(sys.stdin.read())
    workbook_path = Path(task["workbook_path"])
    macro_variants = task["macro_variants"]
    output_rows = {int(k): v for k, v in task["output_rows"].items()}
    dry_run = task["dry_run"]
    col_start = task["project_col_start"]
    col_end = task["project_col_end"]
    name_row = task["project_name_row"]
    toggle_row = task["project_toggle_row"]
    snapshot_sheets = task["snapshot_sheets"]

    # Copy workbook to temp
    tmp_dir = Path(tempfile.mkdtemp(prefix="38dn_iso_"))
    temp_path = tmp_dir / workbook_path.name
    shutil.copy2(str(workbook_path), str(temp_path))

    excel = None
    wb = None
    result = {"status": "error", "projects": [], "macro_used": None,
              "duration": 0, "workbook": workbook_path.name}
    start = time.time()

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        wb = excel.Workbooks.Open(str(temp_path), ReadOnly=False, UpdateLinks=0)

        if dry_run:
            projects = _extract_projects(wb, output_rows, col_start, col_end, name_row, toggle_row)
            result = {"status": "dry_run", "projects": projects,
                      "sheets": [wb.Sheets(i+1).Name for i in range(wb.Sheets.Count)],
                      "macro_variants": macro_variants,
                      "workbook": workbook_path.name,
                      "duration": time.time() - start}
        else:
            # Run macro
            macro_used = None
            for macro_name in macro_variants:
                try:
                    excel.Application.Run(f"'{wb.Name}'!{macro_name}")
                    macro_used = macro_name
                    break
                except Exception as e:
                    err = str(e)
                    if "macro may not be available" in err.lower() or "cannot run" in err.lower():
                        continue
                    raise

            if macro_used is None:
                result = {"status": "error", "error": "No matching macro name found",
                          "tried": macro_variants, "workbook": workbook_path.name,
                          "duration": time.time() - start, "projects": []}
            else:
                # Wait for calculation
                while excel.CalculationState != 0:
                    time.sleep(0.5)
                excel.Calculate()
                time.sleep(1)

                projects = _extract_projects(wb, output_rows, col_start, col_end, name_row, toggle_row)

                # Capture snapshots
                snapshots = {}
                for sheet_name in snapshot_sheets:
                    try:
                        ws = wb.Sheets(sheet_name)
                        snap = {}
                        for r in range(1, 51):
                            for c in range(1, 15):
                                v = ws.Cells(r, c).Value
                                if v is not None:
                                    snap[f"R{r}C{c}"] = v if isinstance(v, (int, float)) else str(v)
                        if snap:
                            snapshots[sheet_name] = snap
                    except Exception:
                        pass

                # Save solved workbook
                solved_name = workbook_path.stem + "_SOLVED" + workbook_path.suffix
                solved_path = workbook_path.parent / solved_name
                try:
                    wb.SaveAs(str(solved_path))
                    saved_to = str(solved_path)
                except Exception:
                    saved_to = None

                result = {
                    "status": "success",
                    "macro_used": macro_used,
                    "projects": projects,
                    "snapshots": snapshots,
                    "saved_to": saved_to,
                    "workbook": workbook_path.name,
                    "duration": time.time() - start,
                }

    except Exception as e:
        result = {"status": "error", "error": str(e),
                  "workbook": workbook_path.name,
                  "duration": time.time() - start, "projects": []}
    finally:
        if wb is not None:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel is not None:
            try: excel.Quit()
            except: pass
        try: shutil.rmtree(tmp_dir, ignore_errors=True)
        except: pass
        try: pythoncom.CoUninitialize()
        except: pass

    # Write result as JSON to stdout
    print(json.dumps(result, default=str))


def _extract_projects(wb, output_rows, col_start, col_end, name_row, toggle_row):
    """Extract active project outputs from workbook via COM."""
    ws = wb.Sheets("Project Inputs")
    projects = []
    for col in range(col_start, col_end + 1):
        name = ws.Cells(name_row, col).Value
        if not name or not str(name).strip():
            continue
        toggle = ws.Cells(toggle_row, col).Value
        is_on = str(toggle).strip().lower() in ("1", "on", "true") if toggle else False
        if not is_on:
            continue

        clean_name = " | ".join(l.strip() for l in str(name).strip().splitlines() if l.strip())
        outputs = {}
        for row, label in output_rows.items():
            v = ws.Cells(row, col).Value
            try:
                outputs[label] = float(v) if v is not None else None
            except (ValueError, TypeError):
                outputs[label] = None

        projects.append({"name": clean_name, "col": col, "outputs": outputs})
    return projects


def run_isolated(workbook_path: Path = None, dry_run: bool = False,
                 max_seconds: int = 300, custom_macro: str = None) -> dict:
    """
    Public API: Run the macro in a fully isolated subprocess.
    Your Python session and your open Excel are completely safe.

    Args:
        workbook_path: Path to .xlsm workbook (default: from config)
        dry_run: If True, inspect without running macro
        max_seconds: Timeout before killing the subprocess
        custom_macro: Override macro name to try first

    Returns:
        dict with status, projects, duration, etc.
    """
    if workbook_path is None:
        workbook_path = DEFAULT_WORKBOOK

    if not workbook_path.exists():
        return {"status": "error", "error": f"Workbook not found: {workbook_path}",
                "projects": []}

    variants = list(MACRO_VARIANTS)
    if custom_macro:
        variants.insert(0, custom_macro)

    print(f"\n{'='*60}")
    print(f"  38DN Isolated Macro Runner")
    print(f"  Workbook: {workbook_path.name}")
    print(f"  Mode: {'DRY RUN' if dry_run else 'LIVE (subprocess-isolated)'}")
    print(f"  Timeout: {max_seconds}s")
    print(f"  Your Excel session: SAFE (completely separate process)")
    print(f"{'='*60}\n")

    result = _run_in_subprocess(
        workbook_path=str(workbook_path),
        macro_variants=variants,
        output_rows=OUTPUT_ROWS,
        dry_run=dry_run,
        max_seconds=max_seconds,
    )

    # Print results
    if result["status"] == "success":
        projects = result.get("projects", [])
        print(f"  Macro: {result.get('macro_used', '?')}")
        print(f"  Duration: {result.get('duration', 0):.1f}s")
        print(f"\n  Results ({len(projects)} active projects):")
        print(f"  {'Project':<35} {'NPP $/W':>10} {'FMV $/W':>10} {'Dev Fee':>10}")
        print(f"  {'-'*35} {'-'*10} {'-'*10} {'-'*10}")
        for p in projects:
            o = p["outputs"]
            npp = o.get("NPP ($/W)")
            fmv = o.get("FMV Calculated ($/W)")
            dev = o.get("Dev Fee ($/W)")
            print(f"  {p['name']:<35} {f'${npp:.3f}' if npp else '—':>10} "
                  f"{f'${fmv:.3f}' if fmv else '—':>10} "
                  f"{f'${dev:.3f}' if dev else '—':>10}")
        if result.get("saved_to"):
            print(f"\n  Solved workbook: {result['saved_to']}")
    elif result["status"] == "dry_run":
        projects = result.get("projects", [])
        print(f"  [DRY RUN] Active projects: {len(projects)}")
        for p in projects:
            print(f"    - {p['name']} (col {p['col']})")
        print(f"\n  Sheets: {', '.join(result.get('sheets', []))}")
    else:
        print(f"  ERROR: {result.get('error', 'Unknown')}")

    return result


if __name__ == "__main__":
    if "--worker" in sys.argv:
        # Running as subprocess worker — handle COM
        _worker_main()
    else:
        # Running as CLI — use isolated execution
        import argparse
        parser = argparse.ArgumentParser(description="38DN Isolated Macro Runner")
        parser.add_argument("workbook", nargs="?", default=None)
        parser.add_argument("--dry-run", action="store_true")
        parser.add_argument("--timeout", type=int, default=300)
        parser.add_argument("--macro", default=None)
        args = parser.parse_args()

        wb_path = Path(args.workbook) if args.workbook else None
        result = run_isolated(wb_path, dry_run=args.dry_run,
                              max_seconds=args.timeout, custom_macro=args.macro)

        # Save to SQLite
        if result["status"] == "success":
            from db import get_connection, save_run
            conn = get_connection()
            for p in result.get("projects", []):
                o = p["outputs"]
                save_run(
                    conn,
                    workbook_name=result["workbook"],
                    macro_name=result.get("macro_used", "unknown"),
                    project_name=p["name"],
                    project_col=p["col"],
                    npp_per_w=o.get("NPP ($/W)"),
                    npp_total=o.get("NPP ($)"),
                    fmv_per_w=o.get("FMV Calculated ($/W)"),
                    dev_fee_per_w=o.get("Dev Fee ($/W)"),
                    target_irr=o.get("Target IRR"),
                    live_irr=o.get("Live Levered Pre-Tax IRR"),
                    status="success",
                    duration_sec=result.get("duration"),
                    raw_outputs={"project_outputs": o,
                                 "snapshots": result.get("snapshots", {})},
                )
            conn.close()
            print(f"\n  Results saved to SQLite.")
