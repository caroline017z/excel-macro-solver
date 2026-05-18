"""dn38_solver.com.auto_recovery — Re-import macro + retry once on first
COM exception from a per-project solve.

Motivated by the 2026-05-13 RP Puma incident: a stale workbook state
(left by an earlier openpyxl save) caused SolveOneProjectByColHL to
throw a generic COM error. The fix was a second `import_vba_module.py`
pass — Excel COM's SaveAs fully rewrote the file and cleared the
corruption. The solver itself was fine.

This module wraps that recovery as an automatic single-shot retry: when
SolveOneProjectByColHL throws a decode-as-auto-recoverable error on
the first project of a worker, the worker closes the workbook, hands
off to a fresh Excel session to re-import the .bas, reopens, and
retries the failed project once. If the retry also fails, the worker
gives up — no infinite recovery loops.

The retry is per-WORKBOOK, not per-project. Once a workbook has been
re-imported during a run, subsequent failures don't trigger another
re-import (the recovery is already in effect).
"""
from __future__ import annotations

import logging
import subprocess
import sys
import time
from pathlib import Path
from typing import Callable

log = logging.getLogger(__name__)

# Path to import_vba_module.py — colocated with SolveHeadless.bas
# at the repo root. Resolve relative to this module so the helper works
# regardless of cwd.
_REPO_ROOT = Path(__file__).parent.parent.parent
_IMPORT_SCRIPT = _REPO_ROOT / "import_vba_module.py"


class AutoRecoveryUnavailable(RuntimeError):
    """Raised when the recovery flow cannot run (missing script, etc).

    Caller should fall back to surfacing the original COM error rather
    than masking with a different error.
    """


def reimport_macro_subprocess(workbook_path: Path | str, timeout_sec: int = 120) -> None:
    """Spawn a subprocess to run import_vba_module.py against the given
    workbook. Returns cleanly on success; raises CalledProcessError on
    non-zero exit. Uses a separate Excel instance so it cannot collide
    with the orchestrator's open workbook in the parent process.

    The parent should close its handle to the workbook before calling
    this and re-open after — Excel COM file locks would otherwise block
    the SaveAs inside import_vba_module.
    """
    if not _IMPORT_SCRIPT.exists():
        raise AutoRecoveryUnavailable(
            f"import_vba_module.py not found at {_IMPORT_SCRIPT}"
        )
    wb = Path(workbook_path)
    log.info(
        "  Auto-recovery: re-importing macro into %s via subprocess...", wb.name
    )
    result = subprocess.run(
        [sys.executable, "-u", str(_IMPORT_SCRIPT), str(wb)],
        capture_output=True,
        text=True,
        timeout=timeout_sec,
    )
    if result.returncode != 0:
        log.warning(
            "  Auto-recovery: macro re-import failed (exit %d).\n  stdout: %s\n  stderr: %s",
            result.returncode,
            result.stdout[-500:],
            result.stderr[-500:],
        )
        raise AutoRecoveryUnavailable(
            f"re-import subprocess failed exit={result.returncode}"
        )
    log.info("  Auto-recovery: macro re-import OK.")


def with_recovery(
    *,
    workbook_path: Path | str,
    close_open_handle: Callable[[], None],
    reopen_handle: Callable[[], None],
    retry_callable: Callable[[], None],
    already_recovered: bool,
) -> bool:
    """Drive the close → re-import → reopen → retry flow.

    Returns:
        True  — recovery+retry succeeded; caller should proceed as if
                first attempt had worked.
        False — recovery cycle ran but retry still failed (caller should
                surface the recovery-pass error to the user).

    Raises:
        AutoRecoveryUnavailable — if the prerequisite tooling is missing
            or `already_recovered` is True (one attempt only).
    """
    if already_recovered:
        raise AutoRecoveryUnavailable(
            "already attempted auto-recovery once for this workbook"
        )

    close_open_handle()
    # Brief pause so OS-level file handles drop before the subprocess
    # opens the file. Excel sometimes lags 100-300ms releasing locks.
    time.sleep(0.5)

    reimport_macro_subprocess(workbook_path)

    # Brief pause so the imported workbook's filesystem state stabilizes
    # before we reopen. Mirrors the verify-retry loop in import_vba_module.
    time.sleep(0.5)
    reopen_handle()

    try:
        retry_callable()
    except RuntimeError as e:
        # Narrow to RuntimeError because that's the type _retry raises in
        # direct_runner.py. A broad `except Exception` would swallow real
        # COM exceptions (e.g., new HRESULT during retry), which deserve
        # to propagate up the stack with their actual error info rather
        # than be flattened into "retry also failed" (Agent 1 P2-8 from
        # the 2026-05-15 audit).
        log.warning("  Auto-recovery: retry also failed: %s", e)
        return False

    log.info("  Auto-recovery: SUCCESS -- retry converged after re-import.")
    return True
