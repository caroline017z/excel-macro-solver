"""dn38_solver.com.cleanup — Excel process cleanup utilities.

Two patterns supported, in order of preference:

1. **PID-scoped (preferred)**: when a worker process spawned an Excel
   COM instance, call `kill_excel_children(worker_pid)` to terminate
   just that worker's Excel children. Surgical — never touches the
   user's interactive Excel or unrelated automation.

2. **Heuristic last-resort**: `kill_hidden_excel_orphans()` kills any
   Excel process with no visible main window. Use only when no PID is
   tracked (e.g., post-crash recovery where the worker died before
   reporting its child PIDs).
"""
from __future__ import annotations

import contextlib
import logging
import subprocess

log = logging.getLogger(__name__)


def kill_excel_children(parent_pid: int, *, timeout_sec: float = 5.0) -> int:
    """Terminate any EXCEL.EXE children of `parent_pid`. Returns count killed.

    Walks the process tree under `parent_pid` (recursive) and signals
    only descendants whose name matches `excel.exe` (case-insensitive).
    Calls `terminate()` first for a clean shutdown, then `kill()` for
    any process still alive after `timeout_sec`.

    The parallel runner (planned in Issue #8) calls this on worker crash
    or post-run cleanup so a zombie EXCEL.EXE doesn't survive — without
    risking the user's interactive Excel or another worker's Excel.

    Returns 0 (no-op) if psutil is unavailable; logs a warning. The
    caller can fall back to `kill_hidden_excel_orphans` if a guaranteed-
    coverage sweep is required.
    """
    try:
        import psutil
    except ImportError:
        log.warning(
            "psutil not installed — kill_excel_children is a no-op. "
            "Install with: pip install psutil"
        )
        return 0

    try:
        parent = psutil.Process(parent_pid)
        children = parent.children(recursive=True)
    except psutil.NoSuchProcess:
        return 0

    excel_kids = [
        c for c in children
        if c.name().lower() in {"excel.exe", "excel"}
    ]
    if not excel_kids:
        return 0

    for c in excel_kids:
        with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied):
            c.terminate()
    gone, alive = psutil.wait_procs(excel_kids, timeout=timeout_sec)
    for c in alive:
        with contextlib.suppress(psutil.NoSuchProcess, psutil.AccessDenied):
            c.kill()

    killed = len(gone) + len(alive)
    if killed:
        log.info(
            "Killed %d Excel child process(es) under PID %d",
            killed, parent_pid,
        )
    return killed


def kill_hidden_excel_orphans() -> None:
    """Last-resort: kill all Excel processes with no main window.

    WARNING: this also kills hidden Excel from unrelated automation
    (other Python scripts, scheduled tasks, COM clients). Prefer
    `kill_excel_children(worker_pid)` when a PID is tracked.
    """
    with contextlib.suppress(Exception):
        result = subprocess.run(
            [
                "powershell", "-Command",
                "Get-Process excel -ErrorAction SilentlyContinue | "
                "Where-Object { $_.MainWindowHandle -eq 0 } | "
                "Stop-Process -Force",
            ],
            capture_output=True,
            timeout=10,
        )
        if result.returncode == 0:
            log.info("Cleaned up orphaned hidden Excel processes (heuristic)")


# Backwards-compat alias. No callers in-tree today, but preserved in case
# downstream tooling imports the old name.
kill_orphan_excel = kill_hidden_excel_orphans
