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


def _kill_excel_kids_for_process(parent, *, timeout_sec: float) -> int:
    """Inner helper: walk descendants of an already-validated psutil.Process
    and reap Excel children. Caller is responsible for handling NoSuchProcess
    when constructing or refreshing `parent`.
    """
    import psutil  # already imported by caller; re-import is a no-op
    try:
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
        try:
            pid_label = parent.pid
        except Exception:
            pid_label = "?"
        log.info(
            "Killed %d Excel child process(es) under PID %s",
            killed, pid_label,
        )
    return killed


def kill_excel_children(parent_pid: int, *, timeout_sec: float = 5.0) -> int:
    """Terminate any EXCEL.EXE children of `parent_pid`. Returns count killed.

    Walks the process tree under `parent_pid` (recursive) and signals
    only descendants whose name matches `excel.exe` (case-insensitive).
    Calls `terminate()` first for a clean shutdown, then `kill()` for
    any process still alive after `timeout_sec`.

    The parallel runner calls this on worker crash or post-run cleanup
    so a zombie EXCEL.EXE doesn't survive — without risking the user's
    interactive Excel or another worker's Excel.

    PID reuse note: this looks up `parent_pid` fresh; on Windows after the
    worker has already exited, the PID may have been reused by the OS for
    an unrelated process. To defeat that race, the caller should prefer
    `kill_excel_children_for_handle` with a `psutil.Process` captured at
    spawn time (psutil pins create_time and raises NoSuchProcess on reuse).

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
    except psutil.NoSuchProcess:
        return 0
    return _kill_excel_kids_for_process(parent, timeout_sec=timeout_sec)


def kill_excel_children_for_handle(parent, *, timeout_sec: float = 5.0) -> int:
    """PID-reuse-safe variant: takes a `psutil.Process` captured at spawn.

    psutil.Process stores the target's create_time at construction and
    validates it on each operation; if the OS has reused the PID for a
    different process, calls raise NoSuchProcess (which we swallow into
    a 0 return). This is the right entry point any time the worker may
    already have exited before cleanup runs.
    """
    if parent is None:
        return 0
    return _kill_excel_kids_for_process(parent, timeout_sec=timeout_sec)


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
