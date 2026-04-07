"""Full portfolio test with per-project timing, convergence validation, and bug tracking.

Tests the solver against a new workbook that has NOT been previously solved
(seed values: NPP=0.2, DevFee=1 across all projects).
"""
from __future__ import annotations

import logging
import sys
import time
from pathlib import Path

# Target workbook
WB_PATH = Path(
    r"C:\Users\CarolineZepecki\Desktop"
    r"\38DN-NY_IL_Lightstar Project Delta_PricingModel_2026.04.07.xlsm"
)

# Convergence targets from the pricing model
TARGET_EQUITY_PCT = 0.10       # 10%
EQUITY_TOL = 0.005             # +/- 0.5pp
TARGET_IRR_TOL = 0.0003        # 0.03%
TARGET_APPR_TOL = 0.0003       # 0.03%


def main() -> None:
    if not WB_PATH.exists():
        print(f"ERROR: Workbook not found: {WB_PATH}")
        sys.exit(1)

    logging.basicConfig(level=logging.INFO, format="%(message)s", stream=sys.stdout)
    log = logging.getLogger(__name__)

    # First: import VBA module into this workbook
    log.info("=" * 70)
    log.info("  PORTFOLIO TEST: %s", WB_PATH.name)
    log.info("=" * 70)

    log.info("\n[Step 1] Importing SolveHeadless VBA module...")
    _import_vba(WB_PATH)

    # Run dry-run to capture pre-solve state
    log.info("\n[Step 2] Pre-solve snapshot (dry-run)...")
    from dn38_solver.shadow.reader import WorkbookReader
    from dn38_solver.config import OUTPUT_ROWS

    with WorkbookReader(WB_PATH) as reader:
        projects = reader.extract_active_projects()
        original_f2 = reader.cell_value("Project Inputs", 2, 6)

    log.info("  %d active projects, F2=%s", len(projects), original_f2)
    for p in projects:
        log.info("  - %s (col %s, offset %d)", p.name, p.col_letter, p.offset)

    # Run the solve with per-project timing
    log.info("\n[Step 3] Running solver...")
    from dn38_solver.com.direct_runner import run_direct
    from dn38_solver.solver.sequence import build_solve_task

    tasks = [build_solve_task(p, str(WB_PATH)) for p in projects]

    t_start = time.time()
    result = run_direct(
        workbook_path=str(WB_PATH),
        tasks=tasks,
        original_f2=int(original_f2) if original_f2 else 1,
        timeout_sec=600,
    )
    t_total = time.time() - t_start

    # Analyze results
    log.info("\n[Step 4] Results Analysis")
    log.info("=" * 70)

    bugs: list[str] = []

    log.info("  Timing:")
    log.info("    Open:       %6.1fs", result.get("open_time_sec", 0))
    log.info("    Warmup:     %6.1fs", result.get("warmup_time_sec", 0))
    log.info("    Macro:      %6.1fs", result.get("solve_time_sec", 0))
    log.info("    Read:       %6.1fs", result.get("read_time_sec", 0))
    log.info("    Total:      %6.1fs", t_total)
    log.info("    Macro used: %s", result.get("macro_used"))
    log.info("    Status:     %s", result.get("status"))

    if result.get("error"):
        log.error("    ERROR: %s", result["error"])
        bugs.append(f"Batch error: {result['error']}")

    raw_results = result.get("project_results", [])
    log.info("\n  Per-Project Results (%d):", len(raw_results))
    log.info("  %-32s %10s %10s %10s %8s %8s %8s",
             "Project", "NPP $/W", "Dev Fee", "FMV", "DSCR", "Eq%", "Status")
    log.info("  %s %s %s %s %s %s %s",
             "-" * 32, "-" * 10, "-" * 10, "-" * 10, "-" * 8, "-" * 8, "-" * 8)

    from dn38_solver.convert import safe_float

    for i, (proj, raw) in enumerate(zip(projects, raw_results)):
        sv = raw.get("solved_values", {})
        col = proj.col_letter

        npp = safe_float(sv.get(f"Project Inputs!{col}38"))
        dev = safe_float(sv.get(f"Project Inputs!{col}32"))
        fmv = safe_float(sv.get(f"Project Inputs!{col}33"))
        dscr = safe_float(sv.get("PT Returns!F129"))
        eq_val = safe_float(sv.get("PT Returns!C128"))
        uses_val = safe_float(sv.get("PT Returns!C130"))
        irr_live = safe_float(sv.get("Project Inputs!F37"))
        irr_tgt = safe_float(sv.get("Project Inputs!F36"))
        appr_live = safe_float(sv.get("Project Inputs!F31"))
        wacc_tgt = safe_float(sv.get("Project Inputs!F30"))

        eq_pct = eq_val / uses_val if eq_val and uses_val and uses_val != 0 else None

        npp_s = f"${npp:.4f}" if npp is not None else "—"
        dev_s = f"${dev:.4f}" if dev is not None else "—"
        fmv_s = f"${fmv:.4f}" if fmv is not None else "—"
        dscr_s = f"{dscr:.4f}x" if dscr is not None else "—"
        eq_s = f"{eq_pct:.2%}" if eq_pct is not None else "—"

        log.info("  %-32s %10s %10s %10s %8s %8s %8s",
                 proj.name[:32], npp_s, dev_s, fmv_s, dscr_s, eq_s, raw.get("status", "?"))

        # --- Convergence validation ---
        project_bugs: list[str] = []

        # Check NPP changed from seed (0.2)
        if npp is not None and abs(npp - 0.2) < 0.0001:
            project_bugs.append("NPP unchanged from seed (0.2)")

        # Check DevFee changed from seed (1.0)
        if dev is not None and abs(dev - 1.0) < 0.0001:
            project_bugs.append("DevFee unchanged from seed (1.0)")

        # Check IRR convergence
        if irr_live is not None and irr_tgt is not None:
            irr_gap = abs(irr_live - irr_tgt)
            if irr_gap > TARGET_IRR_TOL:
                project_bugs.append(f"IRR gap {irr_gap:.6f} > tol {TARGET_IRR_TOL}")

        # Check Appraisal convergence
        if appr_live is not None and wacc_tgt is not None:
            appr_gap = abs(appr_live - wacc_tgt)
            if appr_gap > TARGET_APPR_TOL:
                project_bugs.append(f"Appraisal gap {appr_gap:.6f} > tol {TARGET_APPR_TOL}")

        # Check equity near 10%
        if eq_pct is not None and eq_pct < 1.0:  # Skip 100% equity projects
            eq_gap = abs(eq_pct - TARGET_EQUITY_PCT)
            if eq_gap > EQUITY_TOL:
                project_bugs.append(f"Equity {eq_pct:.2%} gap {eq_gap:.4f} > tol {EQUITY_TOL}")

        if project_bugs:
            for b in project_bugs:
                log.warning("    BUG [%s]: %s", proj.name, b)
                bugs.append(f"{proj.name}: {b}")

    # Check result count
    if len(raw_results) != len(projects):
        msg = f"Result count mismatch: {len(projects)} projects, {len(raw_results)} results"
        log.error("  BUG: %s", msg)
        bugs.append(msg)

    # Summary
    macro_time = result.get("solve_time_sec", 0)
    per_project = macro_time / len(projects) if projects else 0

    log.info("\n" + "=" * 70)
    log.info("  SUMMARY")
    log.info("=" * 70)
    log.info("  Projects:     %d", len(projects))
    log.info("  Total time:   %.1fs (%.1f min)", t_total, t_total / 60)
    log.info("  Macro time:   %.1fs (%.1fs per project)", macro_time, per_project)
    log.info("  Read time:    %.1fs", result.get("read_time_sec", 0))
    log.info("  Saved to:     %s", result.get("saved_to", "—"))
    log.info("  Bugs found:   %d", len(bugs))
    for b in bugs:
        log.info("    - %s", b)
    if not bugs:
        log.info("    None — all convergence targets met")
    log.info("=" * 70)


def _import_vba(wb_path: Path) -> None:
    """Import SolveHeadless.bas into the workbook if not already present."""
    import pythoncom
    import win32com.client

    bas_file = Path(__file__).parent / "SolveHeadless.bas"
    module_name = "modSolveHeadless"

    pythoncom.CoInitialize()
    excel = None
    wb = None
    try:
        try:
            excel = win32com.client.Dispatch("Excel.Application")
        except Exception:
            excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        wb = excel.Workbooks.Open(str(wb_path), ReadOnly=False, UpdateLinks=0)

        vb_project = wb.VBProject
        # Remove existing if present
        for i in range(1, vb_project.VBComponents.Count + 1):
            comp = vb_project.VBComponents.Item(i)
            if comp.Name == module_name:
                vb_project.VBComponents.Remove(comp)
                break

        vb_project.VBComponents.Import(str(bas_file))
        wb.Save()
        print(f"  Imported {module_name} into {wb_path.name}")
    finally:
        if wb:
            try: wb.Close(SaveChanges=False)
            except: pass
        if excel:
            try: excel.Quit()
            except: pass
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
