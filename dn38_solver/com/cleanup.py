"""dn38_solver.com.cleanup — Process cleanup utilities.

Safety net for orphaned Excel processes spawned by the COM worker.
Only kills hidden (no main window) Excel processes.
"""
from __future__ import annotations

import contextlib
import logging
import subprocess

log = logging.getLogger(__name__)


def kill_orphan_excel() -> None:
    """Kill Excel processes that have no visible main window.

    This is safe: it never kills the user's visible Excel session.
    Only targets hidden instances spawned by COM automation.
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
            log.info("Cleaned up orphaned hidden Excel processes")
