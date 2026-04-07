"""Extract all VBA module source code from the workbook for inspection."""
from __future__ import annotations

import sys
from pathlib import Path


def main() -> None:
    from dn38_solver.config import DEFAULT_WORKBOOK

    wb_path = DEFAULT_WORKBOOK
    if not wb_path.exists():
        print(f"ERROR: Workbook not found: {wb_path}")
        sys.exit(1)

    import pythoncom
    import win32com.client

    pythoncom.CoInitialize()
    excel = None
    wb = None

    out_dir = Path(__file__).parent / "vba_source"
    out_dir.mkdir(exist_ok=True)

    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False

        wb = excel.Workbooks.Open(str(wb_path), ReadOnly=True, UpdateLinks=0)
        vb_project = wb.VBProject

        for i in range(1, vb_project.VBComponents.Count + 1):
            comp = vb_project.VBComponents.Item(i)
            code_module = comp.CodeModule
            if code_module.CountOfLines > 0:
                code = code_module.Lines(1, code_module.CountOfLines)
                ext = {1: ".bas", 2: ".cls", 3: ".frm", 100: ".cls"}.get(comp.Type, ".txt")
                fname = f"{comp.Name}{ext}"
                (out_dir / fname).write_text(code, encoding="utf-8")
                print(f"  {fname} ({code_module.CountOfLines} lines)")

        print(f"\nExported to: {out_dir}")

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
