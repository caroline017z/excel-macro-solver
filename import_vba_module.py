"""Import SolveHeadless.bas into the workbook's VBA project.

Prerequisite: Excel Trust Center must have
  "Trust access to the VBA project object model" enabled.
  (File > Options > Trust Center > Trust Center Settings > Macro Settings)

Usage:
    python import_vba_module.py
"""
from __future__ import annotations

import sys
import time
from pathlib import Path

BAS_FILE = Path(__file__).parent / "SolveHeadless.bas"
MODULE_NAME = "modSolveHeadless"


def main() -> None:
    # Resolve workbook path from config
    from dn38_solver.config import DEFAULT_WORKBOOK

    wb_path = DEFAULT_WORKBOOK
    if not wb_path.exists():
        print(f"ERROR: Workbook not found: {wb_path}")
        sys.exit(1)
    if not BAS_FILE.exists():
        print(f"ERROR: .bas file not found: {BAS_FILE}")
        sys.exit(1)

    print(f"Workbook: {wb_path.name}")
    print(f"Module:   {BAS_FILE.name}")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    wb = None

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(wb_path), ReadOnly=False, UpdateLinks=0)
        vb_project = wb.VBProject

        # Remove existing module if present (so re-runs are idempotent)
        for i in range(1, vb_project.VBComponents.Count + 1):
            comp = vb_project.VBComponents.Item(i)
            if comp.Name == MODULE_NAME:
                print(f"  Removing existing '{MODULE_NAME}' module...")
                vb_project.VBComponents.Remove(comp)
                break

        # Import the .bas file
        print(f"  Importing '{BAS_FILE.name}'...")
        vb_project.VBComponents.Import(str(BAS_FILE))

        # Verify it landed
        found = False
        for i in range(1, vb_project.VBComponents.Count + 1):
            comp = vb_project.VBComponents.Item(i)
            if comp.Name == MODULE_NAME:
                found = True
                break

        if found:
            print(f"  SUCCESS: '{MODULE_NAME}' is now in the VBA project.")
            wb.Save()
            print(f"  Workbook saved.")
        else:
            print(f"  WARNING: Import ran but module '{MODULE_NAME}' not found.")

    except Exception as exc:
        err = str(exc)
        if "programmatic access" in err.lower() or "1004" in err:
            print(
                "\nERROR: VBA project access is blocked.\n"
                "Enable it in Excel:\n"
                "  File > Options > Trust Center > Trust Center Settings\n"
                "  > Macro Settings > check 'Trust access to the VBA project object model'\n"
            )
        else:
            print(f"\nERROR: {type(exc).__name__}: {exc}")
        sys.exit(1)

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
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
