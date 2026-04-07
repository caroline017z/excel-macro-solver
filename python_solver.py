"""
38DN Pricing Model — Pure Python Solver (No Excel COM)

Replicates the VBA SolveMinEquityWithHoldCo macro entirely in Python
using the `formulas` library to evaluate Excel formulas natively.

Architecture:
  1. formulas.ExcelModel parses all 885K Excel formulas into a Python
     dependency graph
  2. scipy.optimize.brentq replaces Excel GoalSeek for each solve step
  3. The solver follows the exact VBA sequence:
       Set F2 → HoldCo OFF → DSCR GoalSeek → HoldCo ON →
       NPP GoalSeek → Dev Fee GoalSeek → convergence check → repeat
  4. Results saved to SQLite (same schema as COM runner)

Zero Excel COM dependency. Your open Excel session is completely unaffected.

Usage:
    python python_solver.py                          # Solve with defaults
    python python_solver.py path/to/workbook.xlsm    # Specific workbook
    python python_solver.py --dry-run                 # Inspect only
"""
import argparse
import sys
import time
from pathlib import Path

import numpy as np
from scipy.optimize import brentq

from config import DEFAULT_WORKBOOK, OUTPUT_ROWS, PROJECT_COL_START, PROJECT_NAME_ROW, PROJECT_TOGGLE_ROW
from db import get_connection, save_run

# ---------------------------------------------------------------------------
# VBA Constants (exact match from SolveMinEquityWithHoldCo)
# ---------------------------------------------------------------------------
MAX_ITER = 8            # Outer iterations per project
MAX_GS_RETRY = 6        # Alternating NPP/Appraisal passes per iteration
EQUITY_FINAL_TOL = 0.005  # +/- 0.5pp = 5% of 10% target
IRR_TOLERANCE = 0.0003   # 0.03% — NPP goal seek precision
APPR_TOLERANCE = 0.0003  # 0.03% — Appraisal goal seek precision
DSCR_MIN = 0.5           # Safety floor for DSCR goal seek
DSCR_MAX = 5.0           # Safety ceiling for DSCR goal seek
COL_SCAN_LIMIT = 60      # Max columns to scan for projects

# Cell addresses (exact match from VBA constants)
PT_HOLDCO_ONOFF = "'PT Returns'!C134"
PT_EQUITY = "'PT Returns'!C128"
PT_MIN_EQ_TARGET = "'PT Returns'!F128"
PT_DSCR_MULTIPLE = "'PT Returns'!F129"
PT_TOTAL_USES = "'PT Returns'!C130"
PI_PROJ_INDEX = "'Project Inputs'!F2"
PI_IRR_LIVE = "'Project Inputs'!F37"
PI_IRR_TARGET = "'Project Inputs'!F36"
PI_APPR_LIVE = "'Project Inputs'!F31"
PI_WACC_TARGET = "'Project Inputs'!F30"
PI_FIRST_PROJ_COL = 8  # Column H
PI_BASE_COL = 7         # Column G (OFFSET base)

# CalcModelCore sheet order (exact match from VBA)
CALC_CORE_SHEETS = [
    "Project Inputs", "Rate Curves", "Ops Sandbox", "Global",
    "Operations", "Capex", "Safe Harbor", "CL",
    "Perm Debt", "Tax Equity", "Appraisal", "NPP Calc", "PT Returns",
]
CALC_OUTPUT_SHEETS = [
    "Portfolio", "AT Returns_WIP", "Corp Model Output",
    "Cust Prop", "Dashboard", "Table", "Waterfall Sensitivity",
]


class PythonSolver:
    """
    Pure-Python replacement for the VBA SolveMinEquityWithHoldCo macro.
    Uses the `formulas` library to evaluate Excel formulas natively.
    """

    def __init__(self, workbook_path: Path):
        self.workbook_path = workbook_path
        self.wb_name = workbook_path.name
        self.model = None
        self._cell_cache = {}

    def load_model(self):
        """Parse all Excel formulas into a Python calculation graph."""
        import formulas
        print(f"  Loading formulas from {self.wb_name}...")
        t0 = time.time()
        self.model = formulas.ExcelModel().loads(str(self.workbook_path)).finish()
        print(f"  Loaded in {time.time() - t0:.1f}s")
        # Initial calculation to populate all values
        print(f"  Running initial calculation...")
        t0 = time.time()
        self._cell_cache = self.model.calculate()
        print(f"  Calculated in {time.time() - t0:.1f}s ({len(self._cell_cache):,} cells)")

    def _cell_key(self, sheet, cell_ref):
        """Build the formulas library cell key."""
        return f"'[{self.wb_name}]{sheet}'!{cell_ref}"

    def get_cell(self, sheet, cell_ref):
        """Read a cell value from the calculated model."""
        key = self._cell_key(sheet, cell_ref)
        val = self._cell_cache.get(key)
        if val is not None:
            try:
                return float(val)
            except (ValueError, TypeError):
                return val
        return None

    def set_cell(self, sheet, cell_ref, value):
        """Set a cell value and recalculate dependent formulas."""
        key = self._cell_key(sheet, cell_ref)
        # Update the input value in the model
        self.model.inputs[key] = value
        self._cell_cache[key] = value

    def recalc_core(self):
        """Recalculate core model sheets (equivalent to CalcModelCore)."""
        # With the formulas library, we recalculate the entire dependency graph
        # The library handles dependency ordering internally
        self._cell_cache = self.model.calculate(inputs=self.model.inputs)

    def goal_seek(self, target_sheet, target_cell, goal_value,
                  changing_sheet, changing_cell,
                  tol=1e-5, max_iter=1000, bounds=None):
        """
        Python equivalent of Excel GoalSeek.
        Uses scipy.optimize.brentq (bisection) for robust convergence.

        Args:
            target_sheet/cell: The cell whose value should match goal_value
            goal_value: The target value
            changing_sheet/cell: The cell to adjust
            tol: Convergence tolerance
            bounds: (low, high) search range for changing cell
        """
        current_val = self.get_cell(changing_sheet, changing_cell) or 0.0

        if bounds is None:
            # Auto-detect reasonable bounds based on current value
            if abs(current_val) > 0:
                bounds = (current_val * 0.01, current_val * 10)
            else:
                bounds = (-10, 10)

        def objective(x):
            self.set_cell(changing_sheet, changing_cell, x)
            self.recalc_core()
            result = self.get_cell(target_sheet, target_cell)
            if result is None:
                return float('inf')
            return float(result) - float(goal_value)

        try:
            # Try brentq first (requires sign change in interval)
            fa = objective(bounds[0])
            fb = objective(bounds[1])

            if fa * fb < 0:
                result = brentq(objective, bounds[0], bounds[1],
                                xtol=tol, maxiter=max_iter)
            else:
                # Fallback: use secant method starting from current value
                from scipy.optimize import minimize_scalar
                result_obj = minimize_scalar(
                    lambda x: abs(objective(x)),
                    bounds=bounds, method='bounded',
                    options={'xatol': tol, 'maxiter': max_iter}
                )
                result = result_obj.x

            # Set final value
            self.set_cell(changing_sheet, changing_cell, result)
            self.recalc_core()
            return True, result

        except Exception as e:
            print(f"    GoalSeek failed: {e}")
            return False, current_val

    def get_active_projects(self):
        """Scan Project Inputs row 7 for toggled-on projects."""
        projects = []
        for col in range(PI_FIRST_PROJ_COL, PI_FIRST_PROJ_COL + COL_SCAN_LIMIT):
            from openpyxl.utils import get_column_letter
            col_letter = get_column_letter(col)
            name = self.get_cell("Project Inputs", f"{col_letter}{PROJECT_NAME_ROW}")
            if not name or not str(name).strip():
                break
            toggle = self.get_cell("Project Inputs", f"{col_letter}{PROJECT_TOGGLE_ROW}")
            if toggle == 1 or str(toggle).strip().lower() in ("1", "on", "true"):
                projects.append({
                    "name": str(name).strip(),
                    "col": col,
                    "col_letter": col_letter,
                    "offset": col - PI_BASE_COL,
                })
        return projects

    def solve_project(self, project):
        """
        Solve a single project — exact replication of VBA per-project loop.

        Returns dict with solve results.
        """
        col = project["col"]
        col_letter = project["col_letter"]
        offset = project["offset"]
        name = project["name"]

        print(f"\n  Solving: {name} (col {col_letter}, offset {offset})")

        # Route OFFSET formulas to this project
        self.set_cell("Project Inputs", "F2", offset)
        self.recalc_core()

        # Read targets
        irr_target = self.get_cell("Project Inputs", "F36") or 0.18
        wacc_target = self.get_cell("Project Inputs", "F30") or 0.0725

        prev_eq_pct = -999
        converged = False
        final_iter = 0

        for iteration in range(1, MAX_ITER + 1):
            print(f"    Iteration {iteration}/{MAX_ITER}")

            # Step 1: HoldCo OFF → recalc
            self.set_cell("PT Returns", "C134", 0)
            self.recalc_core()

            # Step 2: GoalSeek Min Equity = 10% (changes DSCR Multiple)
            min_eq_target = self.get_cell("PT Returns", "F128")
            if min_eq_target is not None:
                ok, dscr_val = self.goal_seek(
                    "PT Returns", "C128", min_eq_target,
                    "PT Returns", "F129",
                    bounds=(DSCR_MIN, DSCR_MAX),
                    tol=1e-5,
                )
                # Clamp DSCR
                dscr_val = max(DSCR_MIN, min(DSCR_MAX, dscr_val))
                self.set_cell("PT Returns", "F129", dscr_val)
                self.recalc_core()

            # Step 3: HoldCo ON → recalc
            self.set_cell("PT Returns", "C134", 1)
            self.recalc_core()

            # Steps 4-5: Alternating NPP / Appraisal solve
            for inner in range(1, MAX_GS_RETRY + 1):
                # Solve Levered IRR → Target (changes NPP $/W)
                self.goal_seek(
                    "Project Inputs", "F37", irr_target,
                    "Project Inputs", f"{col_letter}38",
                    bounds=(-2.0, 5.0),
                    tol=1e-5,
                )

                # Solve Appraisal IRR → WACC Target (changes Dev Fee $/W)
                self.goal_seek(
                    "Project Inputs", "F31", wacc_target,
                    "Project Inputs", f"{col_letter}32",
                    bounds=(0.0, 10.0),
                    tol=1e-5,
                )

                # Check both within tolerance
                irr_live = self.get_cell("Project Inputs", "F37") or 0
                appr_live = self.get_cell("Project Inputs", "F31") or 0
                irr_gap = abs(irr_live - irr_target)
                appr_gap = abs(appr_live - wacc_target)

                if irr_gap <= IRR_TOLERANCE and appr_gap <= APPR_TOLERANCE:
                    break

            # Step 6: Convergence check — equity % within tolerance of 10%
            equity = self.get_cell("PT Returns", "C128") or 0
            total_uses = self.get_cell("PT Returns", "C130") or 1
            eq_pct = equity / total_uses if total_uses != 0 else 0

            if abs(eq_pct - 0.1) <= EQUITY_FINAL_TOL:
                converged = True
                final_iter = iteration
                break

            if abs(eq_pct - prev_eq_pct) < 0.000005 and iteration > 1:
                final_iter = iteration
                break

            prev_eq_pct = eq_pct
            final_iter = iteration

        # Read final outputs
        npp = self.get_cell("Project Inputs", f"{col_letter}38")
        dev_fee = self.get_cell("Project Inputs", f"{col_letter}32")
        irr_live = self.get_cell("Project Inputs", "F37")
        appr_live = self.get_cell("Project Inputs", "F31")
        dscr = self.get_cell("PT Returns", "F129")
        fmv = self.get_cell("Project Inputs", f"{col_letter}33")

        result = {
            "name": name,
            "col": col,
            "converged": converged,
            "iterations": final_iter,
            "npp_per_w": npp,
            "dev_fee_per_w": dev_fee,
            "fmv_per_w": fmv,
            "irr_live": irr_live,
            "irr_target": irr_target,
            "appr_live": appr_live,
            "wacc_target": wacc_target,
            "dscr_multiple": dscr,
            "equity_pct": eq_pct,
        }

        status = "CONVERGED" if converged else "NOT CONVERGED"
        irr_flag = "OK" if abs((irr_live or 0) - irr_target) <= IRR_TOLERANCE else "CHECK"
        appr_flag = "OK" if abs((appr_live or 0) - wacc_target) <= APPR_TOLERANCE else "CHECK"
        eq_flag = "OK" if abs(eq_pct - 0.1) <= EQUITY_FINAL_TOL else "CHECK"

        print(f"    Status:          {status} ({final_iter} iter)")
        print(f"    Equity % [{eq_flag}]:  {eq_pct:.2%}  (target 10.00%)")
        print(f"    Levered IRR [{irr_flag}]: {(irr_live or 0):.4%}  (target {irr_target:.4%})")
        print(f"    Appraisal [{appr_flag}]: {(appr_live or 0):.4%}  (target {wacc_target:.4%})")
        print(f"    DSCR Multiple:   {(dscr or 0):.4f}x")
        print(f"    NPP ($/W):       ${(npp or 0):.4f}")
        print(f"    Dev Fee ($/W):   ${(dev_fee or 0):.4f}")

        return result

    def solve_all(self, dry_run=False):
        """
        Main entry point — solve all active projects.
        Exact replication of VBA SolveMinEquityWithHoldCo.
        """
        print(f"\n{'='*60}")
        print(f"  38DN Python Solver (No Excel COM)")
        print(f"  Workbook: {self.wb_name}")
        print(f"  Mode: {'DRY RUN' if dry_run else 'LIVE'}")
        print(f"  Engine: formulas library (pure Python)")
        print(f"{'='*60}")

        # Load model
        self.load_model()

        # Save original F2
        original_f2 = self.get_cell("Project Inputs", "F2")

        # Scan for active projects
        projects = self.get_active_projects()
        if not projects:
            print("\n  No projects toggled On (row 7 = 1).")
            return []

        print(f"\n  Active projects: {len(projects)}")
        for p in projects:
            print(f"    - {p['name']} (col {p['col_letter']})")

        if dry_run:
            # Show current values without solving
            for p in projects:
                col_letter = p["col_letter"]
                self.set_cell("Project Inputs", "F2", p["offset"])
                self.recalc_core()
                npp = self.get_cell("Project Inputs", f"{col_letter}38")
                dev = self.get_cell("Project Inputs", f"{col_letter}32")
                fmv = self.get_cell("Project Inputs", f"{col_letter}33")
                print(f"\n    {p['name']}:")
                print(f"      NPP: ${(npp or 0):.4f}/W  |  Dev Fee: ${(dev or 0):.4f}/W  |  FMV: ${(fmv or 0):.4f}/W")
            # Restore F2
            if original_f2 is not None:
                self.set_cell("Project Inputs", "F2", original_f2)
            return []

        # Solve each project
        results = []
        start_time = time.time()

        for i, project in enumerate(projects, 1):
            print(f"\n  [{i}/{len(projects)}]")
            result = self.solve_project(project)
            results.append(result)

        # Restore F2
        if original_f2 is not None:
            self.set_cell("Project Inputs", "F2", original_f2)
            self.recalc_core()

        total_time = time.time() - start_time

        # Summary
        print(f"\n{'='*60}")
        print(f"  SOLVE COMPLETE — {len(results)} project(s) in {total_time:.1f}s")
        print(f"{'='*60}")
        print(f"  {'Project':<30} {'NPP $/W':>10} {'Dev Fee':>10} {'FMV':>10} {'Status':>12}")
        print(f"  {'-'*30} {'-'*10} {'-'*10} {'-'*10} {'-'*12}")
        for r in results:
            npp = f"${r['npp_per_w']:.3f}" if r['npp_per_w'] else "—"
            dev = f"${r['dev_fee_per_w']:.3f}" if r['dev_fee_per_w'] else "—"
            fmv = f"${r['fmv_per_w']:.3f}" if r['fmv_per_w'] else "—"
            status = "OK" if r['converged'] else "CHECK"
            print(f"  {r['name']:<30} {npp:>10} {dev:>10} {fmv:>10} {status:>12}")

        return results


def main():
    parser = argparse.ArgumentParser(
        description="38DN Pure Python Solver (No Excel COM)")
    parser.add_argument("workbook", nargs="?", default=None,
                        help="Path to .xlsm workbook")
    parser.add_argument("--dry-run", action="store_true",
                        help="Inspect without solving")
    args = parser.parse_args()

    workbook_path = Path(args.workbook) if args.workbook else DEFAULT_WORKBOOK
    if not workbook_path.exists():
        print(f"ERROR: Workbook not found: {workbook_path}")
        sys.exit(1)

    solver = PythonSolver(workbook_path)
    results = solver.solve_all(dry_run=args.dry_run)

    # Save to SQLite
    if results:
        conn = get_connection()
        for r in results:
            save_run(
                conn,
                workbook_name=workbook_path.name,
                macro_name="python_solver",
                project_name=r["name"],
                project_col=r["col"],
                npp_per_w=r["npp_per_w"],
                fmv_per_w=r["fmv_per_w"],
                dev_fee_per_w=r["dev_fee_per_w"],
                target_irr=r["irr_target"],
                live_irr=r["irr_live"],
                status="success" if r["converged"] else "check",
                duration_sec=None,
                raw_outputs={
                    "solver": "python_formulas",
                    "iterations": r["iterations"],
                    "equity_pct": r["equity_pct"],
                    "dscr_multiple": r["dscr_multiple"],
                    "appr_live": r["appr_live"],
                    "wacc_target": r["wacc_target"],
                },
            )
        conn.close()
        print(f"\n  Results saved to SQLite.")


if __name__ == "__main__":
    main()
