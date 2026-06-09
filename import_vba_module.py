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
import re
import sys
import time
import zipfile
from pathlib import Path

BAS_FILE = Path(__file__).parent / "SolveHeadless.bas"
MODULE_NAME = "modSolveHeadless"

# Custom document property name used to stamp the imported .bas hash.
# Read by dn38_solver.shadow.preflight.check_macro_hash to detect drift
# between the repo's current SolveHeadless.bas and what's actually
# embedded in this workbook. Stable name = stable diagnostic code D17.
BAS_HASH_PROP = "DN38_BAS_SHA256"

_CUSTOM_XML_PATH = "docProps/custom.xml"
_PROP_RE = re.compile(
    r'(<property[^>]+name="' + BAS_HASH_PROP +
    r'"[^>]*>\s*<vt:[^>]+>)([^<]*)(</vt:[^>]+>\s*</property>)',
    re.IGNORECASE,
)
_PROPERTIES_CLOSE_RE = re.compile(r"</Properties>\s*$")
_FMT_ID = "{D5CDD505-2E9C-101B-9397-08002B2CF9AE}"


def _bas_sha256() -> str:
    return hashlib.sha256(BAS_FILE.read_bytes()).hexdigest()


def _warn_if_bas_dirty() -> None:
    """Warn (non-blocking) when SolveHeadless.bas differs from the committed
    tree before we stamp its hash into a workbook.

    The D17 drift check (preflight.check_macro_hash) judges every workbook
    against the hash of whatever SolveHeadless.bas is on disk right now. If
    that .bas carries uncommitted edits, the hash we stamp here is not
    reproducible from any commit — a `git stash` or a fresh clone silently
    changes what preflight considers "correct", and a workbook stamped now
    becomes un-reverifiable later. Committing the .bas first makes the stamp
    a stable, shareable contract.

    Kept as a warning rather than a hard block on purpose: the --auto-fix
    recovery path invokes this module as a subprocess and must not be gated
    on a clean tree. Best-effort — if git is unavailable or this isn't a
    repo, stay silent.
    """
    import subprocess
    try:
        out = subprocess.run(
            ["git", "status", "--porcelain", "--", str(BAS_FILE)],
            cwd=str(BAS_FILE.parent),
            capture_output=True, text=True, timeout=5, check=False,
        )
    except (OSError, subprocess.SubprocessError):
        return  # no git / not a repo — nothing to check
    if out.returncode != 0:
        return
    if out.stdout.strip():
        print(
            "  WARNING: SolveHeadless.bas has uncommitted changes. The hash "
            "stamped into this workbook will NOT be reproducible from any "
            "commit — commit the .bas first so the D17 drift check stays "
            "verifiable across checkouts."
        )


def _stamp_bas_hash_via_zip(wb_path: Path, hash_value: str) -> None:
    """Write the .bas SHA256 into docProps/custom.xml via zip-level edit.

    Runs AFTER Excel has saved + closed the file so we operate on the
    serialized package directly. Replaces the prior COM-based stamp,
    which had two failure modes: (1) the COM Add() positional-vs-keyword
    arg mismatch, (2) Add() raising E_FAIL after a Delete() that left
    the COM object in a transient state.

    Three cases handled:
      - Property already exists → update <vt:lpwstr> value in place
      - custom.xml exists, no DN38_BAS_SHA256 property → append it
      - custom.xml missing entirely → no-op (rare; would require
        editing [Content_Types].xml + _rels too. The next preflight
        will surface a "no stamp" warning, which is non-blocking.)
    """
    # Read existing custom.xml if present.
    existing: str | None = None
    with zipfile.ZipFile(wb_path, "r") as zin:
        if _CUSTOM_XML_PATH in zin.namelist():
            existing = zin.read(_CUSTOM_XML_PATH).decode("utf-8")

    if existing is None:
        # Missing custom.xml is a rare case for our workbooks (Excel
        # adds it for many features). Skipping stamp here keeps this
        # helper simple; preflight surfaces "no stamp" as a warning.
        print(
            f"  WARNING: docProps/custom.xml not present in {wb_path.name}; "
            f"skipping stamp. Preflight will surface this as a D17 warning."
        )
        return

    if _PROP_RE.search(existing):
        new_xml = _PROP_RE.sub(
            lambda m: m.group(1) + hash_value + m.group(3),
            existing,
            count=1,
        )
    else:
        # Append a new property. pid must be unique; pick max(existing) + 1.
        pids = [int(p) for p in re.findall(r'pid="(\d+)"', existing)]
        next_pid = (max(pids) + 1) if pids else 2
        new_prop = (
            f'<property fmtid="{_FMT_ID}" pid="{next_pid}" '
            f'name="{BAS_HASH_PROP}">'
            f'<vt:lpwstr>{hash_value}</vt:lpwstr>'
            f'</property>'
        )
        new_xml = _PROPERTIES_CLOSE_RE.sub(
            new_prop + "</Properties>", existing, count=1
        )
        if new_xml == existing:
            # No closing tag found — malformed XML; bail rather than corrupt.
            print(
                f"  WARNING: could not locate </Properties> in custom.xml; "
                f"stamp skipped."
            )
            return

    # Rewrite the zip with the updated custom.xml.
    tmp = wb_path.with_suffix(wb_path.suffix + ".stamp.tmp")
    with zipfile.ZipFile(wb_path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                if item.filename == _CUSTOM_XML_PATH:
                    zout.writestr(item, new_xml.encode("utf-8"))
                else:
                    zout.writestr(item, zin.read(item.filename))
    tmp.replace(wb_path)


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
            # The .bas hash stamp now happens AFTER Excel quits (zip-level
            # rewrite of docProps/custom.xml) — see end of main(). The
            # earlier COM-based stamp had two failure modes that left
            # workbooks with stale stamps and tripped D17 errors on the
            # next preflight.
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
            # Stamp the .bas hash into the saved file via zip-level edit.
            # Runs after Excel has flushed and closed so we operate on the
            # final serialized package — bypasses the COM Add() race that
            # left workbooks with stale stamps tripping D17 errors.
            try:
                _warn_if_bas_dirty()
                hash_value = _bas_sha256()
                _stamp_bas_hash_via_zip(wb_path, hash_value)
                print(f"  Stamped {BAS_HASH_PROP} = {hash_value[:12]}...")
            except Exception as hash_exc:
                # Non-fatal: import still works without the stamp.
                print(f"  WARNING: Could not stamp {BAS_HASH_PROP}: {hash_exc}")
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
