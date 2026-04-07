"""
38DN Excel Macro Runner
Runs VBA macros from pricing model workbooks in a hidden Excel instance,
extracts outputs, and stores results in SQLite.

Usage:
    python macro_runner.py                          # Run with defaults
    python macro_runner.py path/to/workbook.xlsm    # Specify workbook
    python macro_runner.py --dry-run                 # Inspect without running
    python macro_runner.py --list-history            # Show past runs
    python macro_runner.py --batch                   # Batch: default dir
    python macro_runner.py --batch C:\\path\\to\\dir   # Batch: custom dir
    python macro_runner.py --batch --pattern "38DN*.xlsm"  # Batch: filter
    python macro_runner.py --batch --dry-run          # Batch: preview only
"""
import argparse
import fnmatch
import json
import os
import shutil
import sys
import tempfile
import time
import uuid
from pathlib import Path

import pythoncom
import win32com.client

from config import (
    DEFAULT_BATCH_DIR, DEFAULT_WORKBOOK, MACRO_VARIANTS, OUTPUT_ROWS,
    PROJECT_COL_END, PROJECT_COL_START, PROJECT_NAME_ROW,
    PROJECT_TOGGLE_ROW, SNAPSHOT_SHEETS,
)
from db import get_batch_runs, get_connection, get_latest_run, get_runs, save_run
from diff_report import capture_state, compute_diff, print_diff_report, save_diff_to_json


def safe_float(v):
    if v is None:
        return None
    try:
        return float(v)
    except (ValueError, TypeError):
        return None


def copy_workbook_to_temp(src: Path) -> Path:
    """Copy workbook to temp dir so the original stays untouched."""
    tmp_dir = Path(tempfile.mkdtemp(prefix="38dn_macro_"))
    dest = tmp_dir / src.name
    shutil.copy2(str(src), str(dest))
    print(f"  Copied to temp: {dest}")
    return dest


def create_excel_instance():
    """Create a hidden Excel COM instance (fresh process)."""
    pythoncom.CoInitialize()
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = False
    excel.DisplayAlerts = False
    excel.ScreenUpdating = False
    excel.EnableEvents = False
    print("  Hidden Excel instance started (PID separate from your session)")
    return excel


def try_run_macro(excel, wb, macro_variants):
    """Try each macro name variant until one works. Returns the name that worked."""
    for macro_name in macro_variants:
        try:
            print(f"  Trying: Application.Run(\"{macro_name}\")...")
            excel.Application.Run(f"'{wb.Name}'!{macro_name}")
            print(f"  SUCCESS: {macro_name}")
            return macro_name
        except Exception as e:
            err = str(e)
            if "macro may not be available" in err.lower() or "cannot run" in err.lower():
                print(f"    Not found, trying next...")
                continue
            # If it's a different error (runtime error in the macro), it ran but failed
            print(f"    Macro ran but errored: {err}")
            raise
    return None


def extract_project_outputs(wb):
    """Extract key output values from Project Inputs sheet for active projects."""
    ws = wb.Sheets("Project Inputs")
    projects = []

    for col in range(PROJECT_COL_START, PROJECT_COL_END + 1):
        name = ws.Cells(PROJECT_NAME_ROW, col).Value
        if not name or not str(name).strip():
            continue
        toggle = ws.Cells(PROJECT_TOGGLE_ROW, col).Value
        is_on = str(toggle).strip().lower() in ("1", "on", "true") if toggle else False
        if not is_on:
            continue

        clean_name = " | ".join(
            line.strip() for line in str(name).strip().splitlines() if line.strip()
        )
        outputs = {}
        for row, label in OUTPUT_ROWS.items():
            outputs[label] = safe_float(ws.Cells(row, col).Value)

        projects.append({
            "name": clean_name,
            "col": col,
            "outputs": outputs,
        })

    return projects


def extract_sheet_snapshot(wb, sheet_name, max_rows=100, max_cols=30):
    """Capture a snapshot of key data from a sheet."""
    try:
        ws = wb.Sheets(sheet_name)
    except Exception:
        return None

    data = {}
    for r in range(1, max_rows + 1):
        for c in range(1, max_cols + 1):
            v = ws.Cells(r, c).Value
            if v is not None:
                data[f"R{r}C{c}"] = str(v) if not isinstance(v, (int, float)) else v
    return data


def run_macro(workbook_path: Path, dry_run: bool = False, batch_id: str = None):
    """Main workflow: copy workbook, open in hidden Excel, run macro, extract results."""
    if not workbook_path.exists():
        print(f"ERROR: Workbook not found: {workbook_path}")
        if batch_id:
            return {"workbook": workbook_path.name, "status": "error",
                    "error": "File not found"}
        sys.exit(1)

    print(f"\n{'='*60}")
    print(f"  38DN Macro Runner")
    print(f"  Workbook: {workbook_path.name}")
    print(f"  Mode: {'DRY RUN' if dry_run else 'LIVE'}")
    print(f"{'='*60}\n")

    # Step 1: Copy to temp
    print("[1/5] Copying workbook to temp location...")
    temp_path = copy_workbook_to_temp(workbook_path)

    excel = None
    wb = None
    start_time = time.time()
    conn = get_connection()

    try:
        # Step 2: Start hidden Excel
        print("\n[2/5] Starting hidden Excel instance...")
        excel = create_excel_instance()

        # Step 3: Open workbook
        print("\n[3/5] Opening workbook...")
        wb = excel.Workbooks.Open(str(temp_path), ReadOnly=False, UpdateLinks=0)
        print(f"  Opened: {wb.Name}")
        print(f"  Sheets: {', '.join(wb.Sheets(i+1).Name for i in range(wb.Sheets.Count))}")

        # Capture pre-macro state for diff reporting
        print("  Capturing pre-macro cell state for diff report...")
        state_before = capture_state(wb)
        print(f"  Captured {len(state_before)} cells")

        if dry_run:
            print("\n[DRY RUN] Inspecting workbook (not running macro)...")
            projects = extract_project_outputs(wb)
            print(f"\n  Active projects found: {len(projects)}")
            for p in projects:
                print(f"    - {p['name']} (col {p['col']})")
                for label, val in p["outputs"].items():
                    v_str = f"{val:.4f}" if val is not None else "—"
                    print(f"        {label}: {v_str}")
            print(f"\n  Macro variants that would be tried:")
            for m in MACRO_VARIANTS:
                print(f"    - {m}")
            wb.Close(SaveChanges=False)
            return

        # Step 4: Run macro
        print("\n[4/5] Running macro...")
        macro_name = try_run_macro(excel, wb, MACRO_VARIANTS)
        if macro_name is None:
            print("\n  ERROR: No macro variant found. Tried:")
            for m in MACRO_VARIANTS:
                print(f"    - {m}")
            print("\n  TIP: Open the workbook in Excel, press Alt+F11,")
            print("  and check Module names / Sub names in the VBA editor.")
            save_run(conn, workbook_name=workbook_path.name, macro_name="NONE_FOUND",
                     status="error", error_message="No matching macro name found",
                     duration_sec=time.time() - start_time, batch_id=batch_id)
            return {"workbook": workbook_path.name, "status": "error",
                    "error": "No matching macro name found"}

        # Wait briefly for Excel to finish calculating
        while excel.CalculationState != 0:  # xlDone = 0
            time.sleep(0.5)
        excel.Calculate()
        time.sleep(1)

        # Capture post-macro state and compute diff
        print("  Capturing post-macro cell state...")
        state_after = capture_state(wb)
        diff_changes = compute_diff(state_before, state_after)
        diff_summary = {
            "total_changes": len(diff_changes),
            "changes": diff_changes,
        }

        # Step 5: Extract results
        print("\n[5/5] Extracting results...")
        duration = time.time() - start_time
        projects = extract_project_outputs(wb)

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

            # Capture sheet snapshots for this run
            snapshots = {}
            for sheet in SNAPSHOT_SHEETS:
                snap = extract_sheet_snapshot(wb, sheet)
                if snap:
                    snapshots[sheet] = snap

            save_run(
                conn,
                workbook_name=workbook_path.name,
                macro_name=macro_name,
                project_name=p["name"],
                project_col=p["col"],
                npp_per_w=npp, npp_total=o.get("NPP ($)"),
                fmv_per_w=fmv, dev_fee_per_w=dev,
                target_irr=o.get("Target IRR"),
                live_irr=o.get("Live Levered Pre-Tax IRR"),
                status="success", duration_sec=duration,
                raw_outputs={"project_outputs": o, "snapshots": snapshots,
                             "diff_summary": diff_summary},
                batch_id=batch_id,
            )

        result = {"workbook": workbook_path.name, "status": "success",
                  "projects": len(projects), "duration": duration}

        print(f"\n  Completed in {duration:.1f}s")
        print(f"  Results saved to: {conn.execute('SELECT COUNT(*) FROM macro_runs').fetchone()[0]} total records")

        # Save the solved workbook alongside the original
        solved_name = workbook_path.stem + "_SOLVED" + workbook_path.suffix
        solved_path = workbook_path.parent / solved_name
        wb.SaveAs(str(solved_path))
        print(f"  Solved workbook saved: {solved_path.name}")

        # Print diff report to console and save JSON
        print_diff_report(diff_changes, workbook_path.name)
        diff_json_path = workbook_path.parent / (workbook_path.stem + "_DIFF.json")
        save_diff_to_json(diff_changes, diff_json_path)

    except Exception as e:
        duration = time.time() - start_time
        result = {"workbook": workbook_path.name, "status": "error",
                  "error": str(e), "duration": duration}
        print(f"\n  ERROR: {e}")
        save_run(conn, workbook_name=workbook_path.name,
                 macro_name=MACRO_VARIANTS[0], status="error",
                 error_message=str(e), duration_sec=duration, batch_id=batch_id)
        if batch_id:
            return result
        raise

    finally:
        # CRITICAL: Always close Excel to prevent ghost processes
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
        # Clean up temp directory
        try:
            shutil.rmtree(temp_path.parent, ignore_errors=True)
        except Exception:
            pass
        pythoncom.CoUninitialize()
        print("  Excel instance closed.")

    conn.close()
    return result


def run_batch(directory: Path, pattern: str = "*.xlsm", dry_run: bool = False):
    """Run macro on all matching workbooks in a directory sequentially."""
    if not directory.exists():
        print(f"ERROR: Directory not found: {directory}")
        sys.exit(1)
    if not directory.is_dir():
        print(f"ERROR: Not a directory: {directory}")
        sys.exit(1)

    # Find matching files
    all_files = sorted(directory.iterdir())
    matched = [f for f in all_files if f.is_file() and fnmatch.fnmatch(f.name, pattern)]

    if not matched:
        print(f"No files matching '{pattern}' found in {directory}")
        return

    batch_id = str(uuid.uuid4())[:8]

    print(f"\n{'='*60}")
    print(f"  38DN Batch Macro Runner")
    print(f"  Directory: {directory}")
    print(f"  Pattern:   {pattern}")
    print(f"  Files:     {len(matched)}")
    print(f"  Batch ID:  {batch_id}")
    print(f"  Mode:      {'DRY RUN' if dry_run else 'LIVE'}")
    print(f"{'='*60}\n")

    if dry_run:
        print("Files that would be processed:")
        for i, f in enumerate(matched, 1):
            print(f"  {i:3d}. {f.name}")
        print()

    results = []
    batch_start = time.time()

    for i, workbook_path in enumerate(matched, 1):
        print(f"\n[Batch {i}/{len(matched)}] {workbook_path.name}")
        print(f"{'-'*60}")
        try:
            result = run_macro(workbook_path, dry_run=dry_run, batch_id=batch_id)
            if result is None:
                result = {"workbook": workbook_path.name, "status": "dry_run"
                          if dry_run else "unknown"}
            results.append(result)
        except Exception as e:
            print(f"  BATCH ERROR (caught): {e}")
            results.append({
                "workbook": workbook_path.name,
                "status": "error",
                "error": str(e),
            })

    batch_duration = time.time() - batch_start

    # Print summary table
    print(f"\n{'='*60}")
    print(f"  Batch Summary  (ID: {batch_id})")
    print(f"{'='*60}")
    print(f"  {'#':<4} {'Workbook':<40} {'Status':<10} {'Time':>6}")
    print(f"  {'-'*4} {'-'*40} {'-'*10} {'-'*6}")

    success_count = 0
    error_count = 0
    for i, r in enumerate(results, 1):
        status = r.get("status", "unknown")
        dur = r.get("duration")
        dur_str = f"{dur:.1f}s" if dur else "--"
        name = r.get("workbook", "?")
        if len(name) > 38:
            name = name[:35] + "..."
        print(f"  {i:<4} {name:<40} {status:<10} {dur_str:>6}")
        if status == "success":
            success_count += 1
        elif status == "error":
            error_count += 1

    print(f"\n  Total: {len(results)} workbooks | "
          f"{success_count} success | {error_count} errors | "
          f"{batch_duration:.1f}s elapsed")
    print(f"  Batch ID for lookup: {batch_id}\n")


def show_history():
    """Display recent macro run history."""
    conn = get_connection()
    runs = get_runs(conn, limit=20)
    if not runs:
        print("No runs recorded yet.")
        return

    print(f"\n{'='*80}")
    print(f"  Recent Macro Runs")
    print(f"{'='*80}")
    print(f"  {'Timestamp':<22} {'Workbook':<30} {'Project':<20} {'NPP':>8} {'Status':>8}")
    print(f"  {'-'*22} {'-'*30} {'-'*20} {'-'*8} {'-'*8}")

    for r in runs:
        ts = r["run_timestamp"][:19]
        wb_name = (r["workbook_name"] or "")[:28]
        proj = (r["project_name"] or "—")[:18]
        npp = f"${r['npp_per_w']:.3f}" if r["npp_per_w"] else "—"
        print(f"  {ts:<22} {wb_name:<30} {proj:<20} {npp:>8} {r['status']:>8}")

    conn.close()


def main():
    parser = argparse.ArgumentParser(description="38DN Excel Macro Runner")
    parser.add_argument("workbook", nargs="?", default=None,
                        help="Path to .xlsm workbook (default: from config)")
    parser.add_argument("--dry-run", action="store_true",
                        help="Open workbook and inspect without running macro")
    parser.add_argument("--list-history", action="store_true",
                        help="Show recent run history from SQLite")
    parser.add_argument("--macro", default=None,
                        help="Override macro name (e.g. 'Module2.MyMacro')")
    parser.add_argument("--batch", nargs="?", const="__default__", default=None,
                        help="Run in batch mode on a directory (default: DEFAULT_BATCH_DIR)")
    parser.add_argument("--pattern", default="*.xlsm",
                        help="File pattern for batch mode (default: *.xlsm)")
    args = parser.parse_args()

    if args.list_history:
        show_history()
        return

    if args.batch is not None:
        batch_dir = (DEFAULT_BATCH_DIR if args.batch == "__default__"
                     else Path(args.batch))
        run_batch(batch_dir, pattern=args.pattern, dry_run=args.dry_run)
        return

    workbook_path = Path(args.workbook) if args.workbook else DEFAULT_WORKBOOK

    if args.macro:
        # Prepend custom macro name to variants list
        from config import MACRO_VARIANTS as variants
        variants.insert(0, args.macro)

    run_macro(workbook_path, dry_run=args.dry_run)


if __name__ == "__main__":
    main()
