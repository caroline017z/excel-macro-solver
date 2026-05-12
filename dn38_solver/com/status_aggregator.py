"""dn38_solver.com.status_aggregator — Merge per-worker status JSONs.

Each worker writes `solver_status_w{id}.json` in its own temp dir. In
parallel mode the Streamlit tracker still reads the canonical
`solver_status.json` next to the project root, so the parent runs a
background thread that polls the worker files and writes an aggregated
view to that path.

Aggregation rules:
- overall phase = least-completed phase across workers (opening >
  solving > reading > complete; error overrides)
- projects = union of all workers' project lists with worker_id added
- elapsed_sec = max across workers
- total_projects = sum across workers
"""
from __future__ import annotations

import contextlib
import json
import logging
import threading
import time
from pathlib import Path

log = logging.getLogger(__name__)

# Phase ordering for "least-completed" rollup. Lower index = earlier phase.
_PHASE_ORDER = ("opening", "solving", "reading", "complete")


def _phase_rank(phase: str) -> int:
    try:
        return _PHASE_ORDER.index(phase)
    except ValueError:
        return -1  # unknown phases (e.g., "error") sort first


def _aggregate(
    payloads: list[dict],
    workbook_path: str,
    *,
    expected_worker_count: int | None = None,
) -> dict:
    """Merge a list of worker status payloads into one aggregate.

    `expected_worker_count` is the number of workers the parent spawned.
    When provided, we refuse to roll up to "complete" until that many
    workers have actually written their terminal status — otherwise a
    slow-to-start worker (still importing pythoncom, opening Excel) gets
    dropped from the rollup and the dashboard prematurely shows the run
    as done. Without this guard, the user sees "complete" while one
    worker is still busy solving.
    """
    if not payloads:
        return {
            "phase": "opening",
            "workbook": workbook_path,
            "total_projects": 0,
            "projects": [],
            "elapsed_sec": 0.0,
        }

    # Error in any worker dominates the overall phase
    if any(p.get("phase") == "error" or p.get("error") for p in payloads):
        overall_phase = "error"
    else:
        # Least-completed phase wins
        phases = [p.get("phase", "opening") for p in payloads]
        overall_phase = min(phases, key=_phase_rank)
        # Don't roll up to "complete" until every expected worker has
        # actually reported. Missing payloads here mean a worker hasn't
        # written its first status yet (still in subprocess startup) —
        # downgrading to "opening" keeps the dashboard honest.
        if (
            overall_phase == "complete"
            and expected_worker_count is not None
            and len(payloads) < expected_worker_count
        ):
            overall_phase = "opening"

    merged_projects: list[dict] = []
    for p in payloads:
        wid = p.get("worker_id")
        for proj in p.get("projects") or []:
            # Avoid mutating caller's dict
            entry = dict(proj)
            if wid is not None and "worker_id" not in entry:
                entry["worker_id"] = wid
            merged_projects.append(entry)

    total = sum(int(p.get("total_projects", 0) or 0) for p in payloads)
    elapsed = max((float(p.get("elapsed_sec", 0) or 0) for p in payloads), default=0.0)
    errors = [p.get("error") for p in payloads if p.get("error")]

    out: dict = {
        "phase": overall_phase,
        "workbook": workbook_path,
        "total_projects": total,
        "projects": merged_projects,
        "elapsed_sec": elapsed,
        "worker_count": len(payloads),
    }
    if errors:
        out["error"] = "; ".join(str(e) for e in errors)
    return out


def _atomic_write_json(path: Path, payload: dict) -> None:
    """Write JSON via a sibling tmp + replace so readers never see a half-write.

    On Windows, `replace` can fail with PermissionError when another process
    (e.g., the Streamlit tracker mid-read) has the target open. Log the
    failure rather than swallow it silently; otherwise a tmp file is left
    behind and no one knows.
    """
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    try:
        tmp_path.write_text(json.dumps(payload, default=str), encoding="utf-8")
    except OSError as exc:
        log.warning("Status tmp-write to %s failed: %s", tmp_path, exc)
        return
    try:
        tmp_path.replace(path)
    except OSError as exc:
        log.warning(
            "Status replace %s -> %s failed (%s); leaving tmp file in place "
            "until next successful write",
            tmp_path, path, exc,
        )


class StatusAggregator(threading.Thread):
    """Background thread that polls worker status files and emits an aggregate.

    Usage:
        agg = StatusAggregator(
            worker_status_paths=[Path("w0/status.json"), Path("w1/status.json")],
            output_path=Path("solver_status.json"),
            workbook_path="…",
            poll_interval=1.0,
        )
        agg.start()
        # ... workers run ...
        agg.stop()  # signal exit, joins automatically
    """

    def __init__(
        self,
        *,
        worker_status_paths: list[Path],
        output_path: Path,
        workbook_path: str,
        poll_interval: float = 1.0,
        expected_worker_count: int | None = None,
    ) -> None:
        super().__init__(daemon=True, name="StatusAggregator")
        self._paths = list(worker_status_paths)
        self._out = output_path
        self._wb_path = workbook_path
        self._interval = poll_interval
        # Default to len(worker_status_paths): the parent passes one path
        # per worker it spawned, so this is the right baseline. Callers
        # can override (e.g., when one worker had no tasks and didn't get
        # a status path).
        self._expected_workers = (
            expected_worker_count
            if expected_worker_count is not None
            else len(worker_status_paths)
        )
        self._stop_evt = threading.Event()

    def run(self) -> None:
        while not self._stop_evt.is_set():
            self._poll_once()
            if self._stop_evt.wait(timeout=self._interval):
                break
        # One final flush so consumers see the terminal state
        self._poll_once()

    def _poll_once(self) -> None:
        payloads: list[dict] = []
        for path in self._paths:
            try:
                payloads.append(json.loads(path.read_text(encoding="utf-8")))
            except (FileNotFoundError, json.JSONDecodeError):
                # Worker hasn't written its first status yet, or is mid-write
                # (the worker's _StatusWriter uses atomic-swap so this is rare)
                continue
        agg = _aggregate(
            payloads, self._wb_path,
            expected_worker_count=self._expected_workers,
        )
        _atomic_write_json(self._out, agg)

    def stop(self, *, join_timeout: float = 5.0) -> None:
        self._stop_evt.set()
        self.join(timeout=join_timeout)
