"""Import SolveHeadless.bas into the workbook's VBA project.

Prerequisite: Excel Trust Center must have
  "Trust access to the VBA project object model" enabled.
  (File > Options > Trust Center > Trust Center Settings > Macro Settings)

Usage:
    python import_vba_module.py                       # uses DEFAULT_WORKBOOK
    python import_vba_module.py path/to/workbook.xlsm # specific workbook
"""
from __future__ import annotations

import hashlib
import sys
import time
from pathlib import Path

BAS_FILE = Path(__file__).parent / "SolveHeadless.bas"
MODULE_NAME = "modSolveHeadless"

# Custom document property name used to stamp the imported .bas hash.
# Read by dn38_solver.shadow.preflight.check_macro_hash to detect drift
# between the repo's current SolveHeadless.bas and what's actually
# embedded in this workbook. Stable name = stable diagnostic code D17.
BAS_HASH_PROP = "DN38_BAS_SHA256"


def _bas_sha256() -> str:
    return hashlib.sha256(BAS_FILE.read_bytes()).hexdigest()


def _stamp_bas_hash(wb: object, hash_value: str) -> None:
    """Write the .bas SHA256 into the workbook's custom doc properties.

    Replaces any prior stamp. Must be called BEFORE wb.SaveAs so the
    property lands in the persisted file. Done via Excel COM (not
    openpyxl) so we don't violate the openpyxl-xlsm save rule.
    """
    cdp = wb.CustomDocumentProperties
    # Delete prior stamp if present — CDPs have no upsert.
    try:
        cdp.Item(BAS_HASH_PROP).Delete()
    except Exception:
        pass
    # Type=4 = msoPropertyTypeString. Value passed as keyword for clarity.
    cdp.Add(
        Name=BAS_HASH_PROP,
        LinkToContent=False,
        Type=4,
        Value=hash_value,
    )


def _verify_macro_via_com(wb_path: Path) -> tuple[bool, int]:
    """Reopen the saved workbook in a fresh Excel COM session and confirm
    modSolveHeadless has executable code (CountOfLines > 0).

    Returns (ok, line_count). ok is True only if the module exists AND has
    at least one line of code. A zip-level check would be cheaper but
    xl/vbaProject.bin is a CFB container with compressed module streams;
    raw substring search produces false negatives.
    """
    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    try:
        xl = win32com.client.DispatchEx("Excel.Application")
        xl.Visible = False
        xl.DisplayAlerts = False
        wb = xl.Workbooks.Open(str(wb_path), ReadOnly=True, UpdateLinks=0)
        try:
            vbp = wb.VBProject
            for i in range(1, vbp.VBComponents.Count + 1):
                c = vbp.VBComponents.Item(i)
                if c.Name == MODULE_NAME:
                    n = c.CodeModule.CountOfLines
                    return (n > 0, n)
            return (False, 0)
        finally:
            try:
                wb.Close(SaveChanges=False)
            except Exception:
                pass
            try:
                xl.Quit()
            except Exception:
                pass
    finally:
        pythoncom.CoUninitialize()


def main() -> None:
    if len(sys.argv) > 1:
        wb_path = Path(sys.argv[1])
    else:
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
    import_landed = False

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

        # Verify it landed in memory
        found = any(
            vb_project.VBComponents.Item(i).Name == MODULE_NAME
            for i in range(1, vb_project.VBComponents.Count + 1)
        )

        if found:
            print(f"  In-memory import OK: '{MODULE_NAME}' present in VBProject.")
            # Stamp the .bas hash into a custom doc property so preflight
            # can detect drift later (D17). Done here so the stamp lands
            # in the same SaveAs that persists the macro import.
            try:
                hash_value = _bas_sha256()
                _stamp_bas_hash(wb, hash_value)
                print(f"  Stamped {BAS_HASH_PROP} = {hash_value[:12]}...")
            except Exception as hash_exc:
                # Non-fatal: import still works without the stamp. Preflight
                # will surface a "no stamp" warning instead of a drift error.
                print(f"  WARNING: Could not stamp {BAS_HASH_PROP}: {hash_exc}")
            # Force the workbook dirty before Save. Without this, Excel can
            # decide our VBProject.Import didn't change "workbook content"
            # and skip re-serializing the VBA stream entirely — Save returns
            # cleanly but the saved file keeps the old VBA blob.
            wb.Saved = False
            try:
                # SaveAs with explicit FileFormat=52 (xlOpenXMLWorkbookMacroEnabled)
                # is more reliable than wb.Save() for VBA-modifying flows
                # because it always re-serializes every part of the package.
                wb.SaveAs(str(wb_path), FileFormat=52)
                print(f"  wb.SaveAs(.xlsm) returned without exception.")
            except Exception as sa_exc:
                # Fall back to plain Save() if SaveAs is blocked for some
                # reason (e.g., file path resolves differently in COM).
                print(f"  SaveAs failed ({sa_exc}); falling back to Save().")
                wb.Save()
                print(f"  wb.Save() returned without exception.")
            import_landed = True
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
        # Close + Quit BEFORE verifying the saved file. Excel can hold the
        # VBA stream in an unflushed buffer until the workbook closes; a
        # zip-read before Close sees a stale vbaProject.bin even though the
        # subsequent Excel session can see the module fine.
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

    # Verify the macro persisted by reopening the workbook in a fresh
    # Excel COM session and inspecting the VBProject. Runs AFTER the prior
    # Excel quit so we read the post-flush state. A small retry loop
    # tolerates OS-level file handle release lag.
    if not import_landed:
        sys.exit(1)

    last_lines = 0
    for _ in range(5):
        ok, n = _verify_macro_via_com(wb_path)
        last_lines = n
        if ok:
            print(f"  SUCCESS: '{MODULE_NAME}' verified in saved xlsm ({n} lines).")
            return
        time.sleep(0.5)

    if last_lines == 0:
        print(
            f"\n  ERROR: '{MODULE_NAME}' exists in the saved file but has 0 "
            f"lines of code.\n"
            "  Likely cause: Trust Center / corporate policy is permitting "
            "the module shell to be created but blocking the code body. "
            "Check 'Disable all macros without notification' or Group Policy.\n"
        )
    else:
        print(
            f"\n  ERROR: '{MODULE_NAME}' not found in saved file. wb.Save() "
            f"returned cleanly but the module was not persisted.\n"
            "  Possible causes:\n"
            "    1. OneDrive/Box AutoSave reverted the file post-save — pause syncing.\n"
            "    2. Antivirus / Defender blocking VBA writes — check exclusions.\n"
            "    3. Corporate policy preventing VBA project modification.\n"
        )
    sys.exit(1)


if __name__ == "__main__":
    main()
