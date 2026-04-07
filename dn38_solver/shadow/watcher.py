"""dn38_solver.shadow.watcher — File change monitor for .xlsm workbooks.

Watches a workbook file's mtime and triggers a callback when it changes.
Useful for re-reading the model after manual edits in Excel.
"""
from __future__ import annotations

import logging
import time
from pathlib import Path
from typing import Callable

log = logging.getLogger(__name__)


class FileWatcher:
    """Polls a file's mtime at a configurable interval."""

    __slots__ = ("_path", "_interval", "_last_mtime", "_running")

    def __init__(self, path: Path, interval_sec: float = 2.0) -> None:
        self._path = path
        self._interval = interval_sec
        self._last_mtime = self._get_mtime()
        self._running = False

    def _get_mtime(self) -> float:
        try:
            return self._path.stat().st_mtime
        except OSError:
            return 0.0

    def check_once(self) -> bool:
        """Check if file changed since last check. Returns True if changed."""
        current = self._get_mtime()
        if current != self._last_mtime:
            self._last_mtime = current
            return True
        return False

    def watch(self, on_change: Callable[[Path], None]) -> None:
        """Blocking loop that calls on_change when file mtime changes.

        Call stop() from another thread to exit.
        """
        self._running = True
        log.info("Watching %s (interval: %.1fs)", self._path.name, self._interval)
        while self._running:
            if self.check_once():
                log.info("Change detected: %s", self._path.name)
                on_change(self._path)
            time.sleep(self._interval)

    def stop(self) -> None:
        """Signal the watch loop to exit."""
        self._running = False
