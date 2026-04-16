"""
38DN Macro-Driven Solver with Per-Project Progress

Drives the SolveHeadless VBA module from Python at a per-project granularity,
so we get live progress reporting without paying the COM round-trip cost of
doing the inner recalc/GoalSeek loop in Python.

Flow:
    1. Copy workbook to temp, open in hidden Excel.
    2. Import SolveHeadless.bas so edits to the .bas take effect every run.
    3. Python scans active projects (row 7 toggles on "Project Inputs").
    4. Python saves original F2, calls Application.Run("InitSolveEnvHL").
    5. For each active project: Application.Run("SolveOneProjectByColHL", col, name, row)
       and read the corresponding __SolverResults row for status.
    6. Python calls Application.Run("FinalizeSolveEnvHL", originalF2).
    7. Save <workbook>_SOLVED.xlsm alongside the original.

Usage:
    python solve_via_macro.py "path/to/workbook.xlsm"
    python solve_via_macro.py --dry-run "path/to/workbook.xlsm"

Prerequisite: Excel Trust Center must have
  "Trust access to the VBA project object model" enabled.
"""
from __future__ import annotations

import argparse
import shutil
import sys
import tempfile
import time
from pathlib import Path

import pythoncom
import win32com.client

BAS_FILE = Path(__file__).parent / "SolveHeadless.bas"
MODULE_NAME = "modSolveHeadless"

COL_SCAN_LIMIT = 60
PI_FIRST_PROJ_COL = 8
PI_ROW_TOGGLE = 7
PI_ROW_NAME = 4

# __SolverResults column layout (matches VBA ResetSolverResultsHL)
RES_COL_DSCR = 3
RES_COL_NPP = 4
RES_COL_DEV_FEE = 5
RES_COL_EQUITY_PCT = 6
RES_COL_IRR_GAP = 7
RES_COL_APPR_GAP = 8
RES_COL_CONVERGED = 9
RES_COL_MODE = 12
RES_COL_SECS = 13


def safe_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def scan_active_projects(wsPI):
    projects = []
    for c in range(PI_FIRST_PROJ_COL, PI_FIRST_PROJ_COL + COL_SCAN_LIMIT):
        name = wsPI.Cells(PI_ROW_NAME, c).Value
        if not name or not str(name).strip():
            break
        toggle = wsPI.Cells(PI_ROW_TOGGLE, c).Value
        if toggle == 1 or str(toggle).strip() == "1":
            clean_name = " | ".join(
                line.strip() for line in str(name).strip().splitlines() if line.strip()
            )
            projects.append((c, clean_name))
    return projects


def import_solve_module(wb):
    """Replace any existing modSolveHeadless with the current .bas file."""
    if not BAS_FILE.exists():
        raise FileNotFoundError(f".bas file not found: {BAS_FILE}")

    vb_project = wb.VBProject
    for i in range(1, vb_project.VBComponents.Count + 1):
        comp = vb_project.VBComponents.Item(i)
        if comp.Name == MODULE_NAME:
            vb_project.VBComponents.Remove(comp)
            break
    vb_project.VBComponents.Import(str(BAS_FILE))


def run_macro(excel, wb, macro_name, *args):
    """Application.Run with workbook-qualified macro name."""
    qualified = f"'{wb.Name}'!{MODULE_NAME}.{macro_name}"
    return excel.Application.Run(qualified, *args)


def main():
    parser = argparse.ArgumentParser(description="38DN Macro-Driven Solver")
    parser.add_argument("workbook", help="Path to .xlsm workbook")
    parser.add_argument("--dry-run", action="store_true",
                        help="List projects without solving")
    parser.add_argument("--timeout", type=int, default=5400,
                        help="Max seconds total (default: 5400)")
    parser.add_argument("--skip-import", action="store_true",
                        help="Skip VBA module re-import (if already imported)")
    args = parser.parse_args()

    workbook_path = Path(args.workbook)
    if not workbook_path.exists():
        print(f"ERROR: Workbook not found: {workbook_path}")
        sys.exit(1)

    tmp_dir = Path(tempfile.mkdtemp(prefix="38dn_macro_"))
    temp_path = tmp_dir / workbook_path.name
    shutil.copy2(str(workbook_path), str(temp_path))

    print(f"\n{'='*70}")
    print(f"  38DN Macro-Driven Solver")
    print(f"  Workbook: {workbook_path.name}")
    print(f"  Mode:     {'DRY RUN' if args.dry_run else 'LIVE'}")
    print(f"{'='*70}")

    excel = None
    wb = None
    total_start = time.time()

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = True
        excel.WindowState = -4140  # xlMinimized -- visible but out of the way
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        print(f"\n  [INIT] Excel instance started ({time.time() - total_start:.1f}s)")

        wb = excel.Workbooks.Open(str(temp_path), ReadOnly=False, UpdateLinks=0)
        wsPI = wb.Sheets("Project Inputs")
        print(f"  [INIT] Workbook opened ({time.time() - total_start:.1f}s)")

        if not args.skip_import:
            try:
                import_solve_module(wb)
                print(f"  [INIT] VBA module '{MODULE_NAME}' imported from {BAS_FILE.name}")
            except Exception as exc:
                msg = str(exc).lower()
                if "programmatic access" in msg or "1004" in msg:
                    print(
                        "\n  ERROR: VBA project access is blocked.\n"
                        "  Enable in Excel: File > Options > Trust Center > Trust Center Settings\n"
                        "  > Macro Settings > 'Trust access to the VBA project object model'\n"
                    )
                    sys.exit(1)
                raise

        projects = scan_active_projects(wsPI)
        original_f2 = int(wsPI.Range("F2").Value or 1)

        print(f"  [SCAN] Found {len(projects)} active projects:\n")
        for i, (col, name) in enumerate(projects, 1):
            print(f"    {i:2d}. {name}  (col {col})")

        if args.dry_run:
            print(f"\n  [DRY RUN] No macros executed.")
            return

        if not projects:
            print(f"\n  No active projects to solve.")
            return

        # Init solve environment (sets manual calc, disables non-core sheets, etc.)
        run_macro(excel, wb, "InitSolveEnvHL")
        wsRes = wb.Sheets("__SolverResults")

        print(f"\n  {'#':>3}  {'Status':<10} {'Mode':<5} {'Time':>7}  "
              f"{'NPP $/W':>10} {'Dev Fee':>10} {'DSCR':>6}  Project")
        print(f"  {'---':>3}  {'------':<10} {'----':<5} {'-----':>7}  "
              f"{'-------':>10} {'-------':>10} {'----':>6}  -------")

        results = []
        try:
            for i, (col, name) in enumerate(projects, 1):
                elapsed = time.time() - total_start
                if elapsed > args.timeout:
                    print(f"\n  TIMEOUT after {elapsed:.0f}s -- stopping at project {i}/{len(projects)}")
                    break

                results_row = i + 1  # row 1 is header
                proj_start = time.time()
                converged_int = run_macro(
                    excel, wb, "SolveOneProjectByColHL", col, name, results_row
                )
                proj_time = time.time() - proj_start

                dscr = safe_float(wsRes.Cells(results_row, RES_COL_DSCR).Value)
                npp = safe_float(wsRes.Cells(results_row, RES_COL_NPP).Value)
                dev = safe_float(wsRes.Cells(results_row, RES_COL_DEV_FEE).Value)
                mode = str(wsRes.Cells(results_row, RES_COL_MODE).Value or "")

                status = "CONVERGED" if converged_int == 1 else "DONE"
                npp_str = f"${npp:.3f}" if npp is not None else "---"
                dev_str = f"${dev:.3f}" if dev is not None else "---"
                dscr_str = f"{dscr:.2f}" if dscr is not None else "---"
                short = name[:40] if len(name) <= 40 else name[:37] + "..."
                print(f"  {i:3d}  {status:<10} {mode:<5} {proj_time:6.1f}s  "
                      f"{npp_str:>10} {dev_str:>10} {dscr_str:>6}  {short}")

                results.append({
                    "name": name, "col": col, "converged": converged_int == 1,
                    "time": proj_time, "npp": npp, "dev_fee": dev, "dscr": dscr,
                })

                solved_name = workbook_path.stem + "_SOLVED" + workbook_path.suffix
                solved_path = workbook_path.parent / solved_name
                try:
                    wb.SaveCopyAs(str(solved_path))
                except Exception as save_exc:
                    print(f"    (incremental save failed: {save_exc})")
        finally:
            # Always try to finalize even on error so workbook state is restored
            try:
                run_macro(excel, wb, "FinalizeSolveEnvHL", original_f2)
            except Exception as fin_exc:
                print(f"\n  WARNING: Finalize failed: {fin_exc}")

        total_time = time.time() - total_start
        converged_count = sum(1 for r in results if r["converged"])

        solved_name = workbook_path.stem + "_SOLVED" + workbook_path.suffix
        solved_path = workbook_path.parent / solved_name
        try:
            wb.SaveAs(str(solved_path))
            print(f"\n  Saved: {solved_path.name}")
        except Exception as e:
            print(f"\n  Save failed: {e}")

        print(f"\n{'='*70}")
        print(f"  DONE  {len(results)} projects | {converged_count} converged | {total_time:.1f}s total")
        print(f"{'='*70}\n")

    except Exception as e:
        print(f"\n  FATAL ERROR: {e}")
        raise

    finally:
        if wb is not None:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
        if excel is not None:
            try:
                excel.Quit()
            except Exception:
                pass
        try:
            shutil.rmtree(tmp_dir, ignore_errors=True)
        except Exception:
            pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
