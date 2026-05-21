"""scripts/patch_pi_row1_toggle.py — Standalone one-off patcher.

Wraps `dn38_solver.patches.pi_row1_toggle.patch_pi_row1_toggle` as a
command-line tool for cleaning up workbooks where the Project Inputs
row 1 toggle multipliers are missing or stray (template ships row 1
populated only through col Q; portfolios extending past that produce
#VALUE! propagation on the Table tab). Used for one-off file cleanup
and bulk patching without paying the cost of a full re-solve.

Usage:
    # Single file
    python -m scripts.patch_pi_row1_toggle path/to/workbook.xlsm

    # Multiple files (mixed source + _SOLVED is fine; the patch is
    # idempotent and only touches PI row 1 formulas, not stamped values)
    python -m scripts.patch_pi_row1_toggle \\
        "C:/path/to/a.xlsm" "C:/path/to/b_SOLVED.xlsm"

Exit code is 0 if every file was patched cleanly (or was already
clean); non-zero if any file failed to patch.
"""
from __future__ import annotations

import logging
import sys
from pathlib import Path

from dn38_solver.patches.pi_row1_toggle import (
    format_patch_report,
    patch_pi_row1_toggle,
)

logging.basicConfig(
    level=logging.INFO,
    format="%(levelname)s %(message)s",
)
log = logging.getLogger(__name__)


def main(argv: list[str]) -> int:
    if len(argv) < 2:
        print(__doc__)
        return 1

    paths = [Path(arg) for arg in argv[1:]]
    failures: list[Path] = []

    for p in paths:
        print(f"\n=== {p.name} ===")
        if not p.exists():
            print(f"  SKIP: file not found ({p})")
            failures.append(p)
            continue

        result = patch_pi_row1_toggle(p)
        for line in format_patch_report(result).splitlines():
            print(f"  {line}")

        if result.status == "patch_failed":
            failures.append(p)

    print()
    if failures:
        print(f"FAILED on {len(failures)}/{len(paths)} file(s):")
        for f in failures:
            print(f"  - {f}")
        return 2

    print(f"OK — {len(paths)} file(s) processed.")
    return 0


if __name__ == "__main__":
    sys.exit(main(sys.argv))
