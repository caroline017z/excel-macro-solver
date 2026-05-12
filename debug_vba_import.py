"""Diagnostic: figure out whether VBA Import loads code or just the module shell.

Reports:
  1. .bas file line count (sanity: should be ~1000)
  2. Module line count immediately after Import (in-memory)
  3. Module line count after Save + Close + reopen in fresh Excel instance
  4. First 5 and last 5 lines of the module's code, post-reopen

Run after import_vba_module.py has been attempted, against the SAME workbook:
    python debug_vba_import.py "C:\Temp\38DN-SMP_PricingModel_2026.05.12_WalkTEST.xlsm"
"""
from __future__ import annotations

import sys
from pathlib import Path

BAS_FILE = Path(__file__).parent / "SolveHeadless.bas"
MODULE_NAME = "modSolveHeadless"


def main() -> None:
    if len(sys.argv) < 2:
        print("Usage: python debug_vba_import.py <workbook.xlsm>")
        sys.exit(1)
    wb_path = Path(sys.argv[1])
    if not wb_path.exists():
        print(f"ERROR: not found: {wb_path}")
        sys.exit(1)

    # --- 1. Source .bas file metrics ---
    bas_lines = BAS_FILE.read_text(encoding="utf-8", errors="replace").splitlines()
    print(f"[1] {BAS_FILE.name}: {len(bas_lines)} lines, {BAS_FILE.stat().st_size} bytes")

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        # --- 2. Fresh Excel session, open the workbook AS-IS (no import) ---
        # We're inspecting what import_vba_module.py left behind on disk.
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False

        wb = xl.Workbooks.Open(str(wb_path), ReadOnly=True, UpdateLinks=0)
        try:
            vbp = wb.VBProject
        except Exception as e:
            print(f"[2] ERROR accessing VBProject: {e}")
            wb.Close(SaveChanges=False)
            xl.Quit()
            sys.exit(1)

        print(f"[2] VBProject opened. {vbp.VBComponents.Count} component(s):")
        target = None
        for i in range(1, vbp.VBComponents.Count + 1):
            c = vbp.VBComponents.Item(i)
            try:
                lines = c.CodeModule.CountOfLines
            except Exception:
                lines = "?"
            print(f"    - {c.Name!r}  type={c.Type}  lines={lines}")
            if c.Name == MODULE_NAME:
                target = c

        if target is None:
            print(f"[3] '{MODULE_NAME}' NOT present in saved file.")
        else:
            cm = target.CodeModule
            n = cm.CountOfLines
            print(f"[3] '{MODULE_NAME}' has {n} lines in saved file.")
            if n == 0:
                print(f"    DIAGNOSIS: module shell saved, but code body is EMPTY.")
                print(f"    Likely cause: Trust Center policy is blocking VBA Import")
                print(f"    body content, OR the .bas file isn't being read.")
            else:
                # Print first/last 5 lines so we can compare against SolveHeadless.bas
                first = cm.Lines(1, min(5, n))
                last = cm.Lines(max(1, n - 4), min(5, n))
                print(f"[4] First 5 lines:")
                for ln in first.splitlines():
                    print(f"    {ln}")
                print(f"[4] Last 5 lines:")
                for ln in last.splitlines():
                    print(f"    {ln}")

        wb.Close(SaveChanges=False)
        xl.Quit()
    finally:
        pythoncom.CoUninitialize()


if __name__ == "__main__":
    main()
