"""
38DN Python-Driven Solver with Per-Project Progress
Replicates the SolveHeadless VBA logic but drives the project loop from Python,
so we can report progress after each project solves.

Usage:
    python solve_with_progress.py "path/to/workbook.xlsm"
    python solve_with_progress.py --dry-run "path/to/workbook.xlsm"
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

# --- Constants matching modSolveHeadless VBA ---
MAX_ITER = 8
MAX_GS_RETRY = 6
EQUITY_FINAL_TOL = 0.005
IRR_TOLERANCE = 0.0003
APPR_TOLERANCE = 0.0003
DSCR_MIN = 0.5
DSCR_MAX = 5.0
COL_SCAN_LIMIT = 60

# Sanity bounds — GoalSeek is unconstrained Newton-style and can diverge to
# absurd values when the local slope is small. If a solve lands outside these
# bounds, snap back to the seed so the next retry starts from a sane point.
NPP_MIN, NPP_MAX, NPP_SEED = -0.20, 0.80, 0.20
DEV_FEE_MIN, DEV_FEE_MAX, DEV_FEE_SEED = 0.05, 0.50, 0.20

PI_FIRST_PROJ_COL = 8
PI_BASE_COL = 7
PI_ROW_TOGGLE = 7
PI_ROW_NAME = 4
PI_ROW_NPP = 38
PI_ROW_DEV_FEE = 32

# Output rows for result extraction
OUTPUT_ROWS = {
    32: "Dev Fee ($/W)",
    33: "FMV Calculated ($/W)",
    38: "NPP ($/W)",
    39: "NPP ($)",
    30: "FMV WACC (Target)",
    31: "Live Appraisal IRR",
    36: "Target IRR",
    37: "Live Levered Pre-Tax IRR",
}


def calc_model_core(wb):
    """Recalculate core pricing waterfall sheets in dependency order."""
    for name in [
        "Project Inputs", "Rate Curves", "Ops Sandbox", "Global",
        "Operations", "Capex", "Safe Harbor", "CL", "Perm Debt",
        "Tax Equity", "Appraisal", "NPP Calc", "PT Returns",
    ]:
        wb.Sheets(name).Calculate()
    pythoncom.PumpWaitingMessages()


def calc_output_sheets(wb):
    """Recalculate output/reporting sheets."""
    for name in [
        "Portfolio", "AT Returns_WIP", "Corp Model Output",
        "Cust Prop", "Dashboard", "Table", "Waterfall Sensitivity",
    ]:
        try:
            ws = wb.Sheets(name)
            ws.EnableCalculation = True
            ws.Calculate()
        except Exception:
            pass


def disable_non_core_sheets(wb):
    """Disable calculation on non-core sheets for performance."""
    core = {
        "Project Inputs", "Rate Curves", "Ops Sandbox", "Global",
        "Operations", "Capex", "Safe Harbor", "CL", "Perm Debt",
        "Tax Equity", "Appraisal", "NPP Calc", "PT Returns",
    }
    for i in range(1, wb.Sheets.Count + 1):
        ws = wb.Sheets(i)
        if ws.Name not in core:
            try:
                ws.EnableCalculation = False
            except Exception:
                pass


def enable_all_sheets(wb):
    """Re-enable calculation on all sheets."""
    for i in range(1, wb.Sheets.Count + 1):
        try:
            wb.Sheets(i).EnableCalculation = True
        except Exception:
            pass


def safe_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def scan_active_projects(wsPI):
    """Scan row 7 for toggled-on projects. Returns list of (col, name)."""
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


def solve_one_project(excel, wb, wsPI, wsPT, col_idx, per_project_timeout=None):
    """Solve a single project -- replicates SolveHeadless per-project logic.
    Returns (converged: bool, iterations: int, equity_pct: float, timed_out: bool).
    per_project_timeout: soft timeout in seconds; checked between iterations and
    GoalSeek retries. Cannot interrupt a GoalSeek mid-call.
    """
    proj_start = time.time()
    def _over_budget():
        return per_project_timeout is not None and (time.time() - proj_start) > per_project_timeout

    proj_offset = col_idx - PI_BASE_COL

    # Route OFFSET formulas to this project
    wsPI.Range("F2").Value = proj_offset
    calc_model_core(wb)

    # PT Returns ranges
    rHoldCo = wsPT.Range("C134")
    rEquity = wsPT.Range("C128")
    rMinEqTgt = wsPT.Range("F128")
    rDSCR = wsPT.Range("F129")
    rTotalUses = wsPT.Range("C130")

    # Project Inputs ranges
    rIRRLive = wsPI.Range("F37")
    rIRRTgt = wsPI.Range("F36")
    rNPP = wsPI.Cells(PI_ROW_NPP, col_idx)
    rApprLive = wsPI.Range("F31")
    rWACCTgt = wsPI.Range("F30")
    rDevFee = wsPI.Cells(PI_ROW_DEV_FEE, col_idx)

    prev_eq_pct = -999.0
    converged = False
    iters_used = 0
    timed_out = False
    equity_pct = 0.0

    # Pre-seed if prior project left wild values in this column
    npp_start = safe_float(rNPP.Value)
    if npp_start is None or npp_start < NPP_MIN or npp_start > NPP_MAX:
        rNPP.Value = NPP_SEED
    dev_start = safe_float(rDevFee.Value)
    if dev_start is None or dev_start < DEV_FEE_MIN or dev_start > DEV_FEE_MAX:
        rDevFee.Value = DEV_FEE_SEED

    for iIter in range(1, MAX_ITER + 1):
        iters_used = iIter
        if _over_budget():
            timed_out = True
            break

        # Step 1: HoldCo OFF + recalc
        rHoldCo.Value = 0
        calc_model_core(wb)

        # Step 2: GoalSeek Min Equity (changes DSCR Multiple)
        rEquity.GoalSeek(Goal=rMinEqTgt.Value, ChangingCell=rDSCR)
        dscr_val = safe_float(rDSCR.Value) or 1.0
        if dscr_val < DSCR_MIN:
            rDSCR.Value = DSCR_MIN
        if dscr_val > DSCR_MAX:
            rDSCR.Value = DSCR_MAX
        calc_model_core(wb)

        # Step 3: HoldCo ON + recalc
        rHoldCo.Value = 1
        calc_model_core(wb)

        # Steps 4-5: Sequential IRR + Appraisal GoalSeek with retries
        for _ in range(MAX_GS_RETRY):
            if _over_budget():
                timed_out = True
                break
            rIRRLive.GoalSeek(Goal=rIRRTgt.Value, ChangingCell=rNPP)
            npp_val = safe_float(rNPP.Value)
            if npp_val is None or npp_val < NPP_MIN or npp_val > NPP_MAX:
                rNPP.Value = NPP_SEED
            calc_model_core(wb)

            rApprLive.GoalSeek(Goal=rWACCTgt.Value, ChangingCell=rDevFee)
            dev_val = safe_float(rDevFee.Value)
            if dev_val is None or dev_val < DEV_FEE_MIN or dev_val > DEV_FEE_MAX:
                rDevFee.Value = DEV_FEE_SEED
            calc_model_core(wb)

            irr_gap = abs((safe_float(rIRRLive.Value) or 0) - (safe_float(rIRRTgt.Value) or 0))
            appr_gap = abs((safe_float(rApprLive.Value) or 0) - (safe_float(rWACCTgt.Value) or 0))
            if irr_gap <= IRR_TOLERANCE and appr_gap <= APPR_TOLERANCE:
                break
        if timed_out:
            break

        # Step 6: Convergence check
        total_uses = safe_float(rTotalUses.Value) or 0
        equity_val = safe_float(rEquity.Value) or 0
        equity_pct = equity_val / total_uses if total_uses != 0 else 0

        if abs(equity_pct - 0.1) <= EQUITY_FINAL_TOL:
            converged = True
            break
        if abs(equity_pct - prev_eq_pct) < 0.000005 and iIter > 1:
            break
        prev_eq_pct = equity_pct

    return converged, iters_used, equity_pct, timed_out


def extract_project_results(wsPI, col_idx):
    """Read output values for a solved project."""
    outputs = {}
    for row, label in OUTPUT_ROWS.items():
        outputs[label] = safe_float(wsPI.Cells(row, col_idx).Value)
    return outputs


def main():
    parser = argparse.ArgumentParser(description="38DN Solver with Progress")
    parser.add_argument("workbook", help="Path to .xlsm workbook")
    parser.add_argument("--dry-run", action="store_true",
                        help="List projects without solving")
    parser.add_argument("--timeout", type=int, default=5400,
                        help="Max seconds total (default: 5400)")
    parser.add_argument("--per-project-timeout", type=int, default=300,
                        help="Soft timeout per project in seconds; skip to next on exceed (default: 300)")
    args = parser.parse_args()

    workbook_path = Path(args.workbook)
    if not workbook_path.exists():
        print(f"ERROR: Workbook not found: {workbook_path}")
        sys.exit(1)

    # Copy to temp
    tmp_dir = Path(tempfile.mkdtemp(prefix="38dn_progress_"))
    temp_path = tmp_dir / workbook_path.name
    shutil.copy2(str(workbook_path), str(temp_path))

    print(f"\n{'='*70}")
    print(f"  38DN Solver with Progress")
    print(f"  Workbook: {workbook_path.name}")
    print(f"  Mode:     {'DRY RUN' if args.dry_run else 'LIVE'}")
    print(f"{'='*70}")

    excel = None
    wb = None
    total_start = time.time()

    try:
        pythoncom.CoInitialize()
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False
        print(f"\n  [INIT] Excel instance started ({time.time() - total_start:.1f}s)")

        wb = excel.Workbooks.Open(str(temp_path), ReadOnly=False, UpdateLinks=0)
        wsPI = wb.Sheets("Project Inputs")
        wsPT = wb.Sheets("PT Returns")
        print(f"  [INIT] Workbook opened ({time.time() - total_start:.1f}s)")

        # Scan for active projects
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

        # Performance setup
        excel.Calculation = -4135  # xlCalculationManual
        excel.MaxChange = 0.00001
        excel.MaxIterations = 1000
        disable_non_core_sheets(wb)

        try:
            excel.MultiThreadedCalculation.Enabled = True
        except Exception:
            pass

        print(f"\n  {'#':>3}  {'Status':<12} {'Iters':>5} {'Time':>7}  {'NPP $/W':>10} {'Dev Fee':>10}  Project")
        print(f"  {'---':>3}  {'------':<12} {'-----':>5} {'-----':>7}  {'-------':>10} {'-------':>10}  -------")

        results = []
        for i, (col, name) in enumerate(projects, 1):
            proj_start = time.time()

            # Check timeout
            elapsed = time.time() - total_start
            if elapsed > args.timeout:
                print(f"\n  TIMEOUT after {elapsed:.0f}s -- stopping at project {i}/{len(projects)}")
                break

            converged, iters, eq_pct, timed_out = solve_one_project(
                excel, wb, wsPI, wsPT, col,
                per_project_timeout=args.per_project_timeout,
            )
            outputs = extract_project_results(wsPI, col)
            proj_time = time.time() - proj_start

            if timed_out:
                status = "TIMEOUT"
            elif converged:
                status = "CONVERGED"
            else:
                status = "DONE"
            npp = outputs.get("NPP ($/W)")
            dev = outputs.get("Dev Fee ($/W)")
            npp_str = f"${npp:.3f}" if npp is not None else "---"
            dev_str = f"${dev:.3f}" if dev is not None else "---"

            short_name = name[:40] if len(name) <= 40 else name[:37] + "..."
            print(f"  {i:3d}  {status:<12} {iters:5d} {proj_time:6.1f}s  {npp_str:>10} {dev_str:>10}  {short_name}")

            results.append({
                "name": name, "col": col, "converged": converged,
                "iterations": iters, "time": proj_time, "outputs": outputs,
            })

        # Restore original F2 and finalize
        wsPI.Range("F2").Value = original_f2
        calc_model_core(wb)
        enable_all_sheets(wb)
        calc_output_sheets(wb)

        # Restore defaults
        excel.MaxChange = 0.001
        excel.MaxIterations = 100

        total_time = time.time() - total_start
        converged_count = sum(1 for r in results if r["converged"])

        # Save solved workbook
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
                excel.Calculation = -4105  # xlCalculationAutomatic
            except Exception:
                pass
            try:
                excel.ScreenUpdating = True
                excel.EnableEvents = True
            except Exception:
                pass
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
