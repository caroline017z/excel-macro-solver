"""dn38_solver.com.status_aggregator — Merge per-worker status JSONs.

Each worker writes `solver_status_w{id}.json` in its own temp dir. In
parallel mode the parent runs a background thread that polls the worker
files, writes an aggregated `solver_status.json` next to the project
root, and runs the per-worker stall detector. No live viewer ships (the
Streamlit dashboard was removed) — the aggregate file is retained for
stall detection and ad-hoc inspection of an in-flight run; live progress
goes to stdout via the orchestrator's logging.

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


# Retry schedule for the .tmp → target rename. On Windows, transient
# PermissionError fires when AV (Defender) momentarily holds an exclusive
# handle on the file just-written, when a Streamlit dashboard reader
# happens to have it open, or when the Windows file-index service is
# mid-scan. Empirically these clear within ~50-500ms; this schedule
# (~1.55s total) catches >99% in practice while bounding the worst case.
# Caller still leaves the .tmp behind on exhaustion so the next successful
# write recovers state.
_RENAME_RETRY_DELAYS_SEC: tuple[float, ...] = (0.05, 0.1, 0.2, 0.4, 0.8)


def atomic_write_json(path: Path, payload: dict) -> None:
    """Write JSON via a sibling tmp + replace so readers never see a half-write.

    On Windows, `replace` can fail with PermissionError when another process
    (Streamlit tracker reading; Defender real-time scan; file-index service)
    has the target open. We retry on a short backoff schedule
    (~50ms..800ms, ~1.55s total) before giving up. Tmp file is left in
    place on exhaustion so the NEXT successful write still recovers state
    — and so debug forensics show the failing payload.

    Without the retry, a transient `WinError 5: Access is denied` from
    AV / Streamlit reader / file-index service silently drops a status
    update even when the underlying solve succeeded.
    """
    tmp_path = path.with_suffix(path.suffix + ".tmp")
    try:
        tmp_path.write_text(json.dumps(payload, default=str), encoding="utf-8")
    except OSError as exc:
        log.warning("Status tmp-write to %s failed: %s", tmp_path, exc)
        return

    last_exc: OSError | None = None
    # First attempt is immediate; only sleep BEFORE each retry, not before
    # the first try. len(delays) + 1 total attempts.
    for attempt, delay in enumerate((0.0,) + _RENAME_RETRY_DELAYS_SEC):
        if delay > 0:
            time.sleep(delay)
        try:
            tmp_path.replace(path)
            if attempt > 0:
                # Visible-but-not-loud: the retry worked, so the run is
                # fine, but the operator should know contention happened
                # (could indicate AV scanning or a stuck dashboard).
                log.info(
                    "Status replace %s -> %s succeeded on attempt %d",
                    tmp_path, path, attempt + 1,
                )
            return
        except OSError as exc:
            last_exc = exc

    log.warning(
        "Status replace %s -> %s failed after %d attempts (%s); leaving "
        "tmp file in place until next successful write. Common cause on "
        "Windows: AV real-time scan, a Streamlit reader holding the "
        "target open, or the Windows file-index service.",
        tmp_path, path, len(_RENAME_RETRY_DELAYS_SEC) + 1, last_exc,
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
        stall_threshold_sec: float = 600.0,
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
        # Stall detection. Track each worker's last-observed elapsed_sec
        # and the wall-clock time at which it last advanced. If it sits
        # unchanged for STALL_THRESHOLD_SEC the aggregator logs a WARNING
        # — this is the operator-visible signal that a worker is hung
        # (Application.Run not returning, Excel grinding silently). The
        # COM-side watchdog in direct_runner will kill Excel at
        # DEFAULT_PER_CALL_TIMEOUT_SEC (600s), but if that mechanism is
        # also broken, the stall detector here is the backstop.
        self._last_elapsed: dict[int, tuple[float, float]] = {}  # wid -> (elapsed, wall_t)
        self._stall_warned: set[int] = set()
        self._stall_threshold_sec = float(stall_threshold_sec)

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

        # Stall check — per-worker, per-poll. Compare each worker's current
        # elapsed_sec against its last observation; if unchanged and the
        # gap since last advance exceeds the threshold, emit one WARNING
        # (deduped via _stall_warned) so the operator gets a clear signal
        # the worker is hung rather than legitimately slow.
        now_wall = time.time()
        for p in payloads:
            wid = p.get("worker_id")
            phase = p.get("phase", "")
            elapsed = float(p.get("elapsed_sec", 0) or 0)
            if wid is None or phase not in ("solving", "reading"):
                continue
            prev = self._last_elapsed.get(wid)
            if prev is None or elapsed > prev[0] + 0.001:
                # advanced (or first observation) — reset stall tracker
                self._last_elapsed[wid] = (elapsed, now_wall)
                self._stall_warned.discard(wid)
            else:
                stalled_for = now_wall - prev[1]
                if (
                    stalled_for >= self._stall_threshold_sec
                    and wid not in self._stall_warned
                ):
                    log.warning(
                        "  Worker %d STALL detected: elapsed_sec stuck at "
                        "%.1fs for %.0fs (project=%s, phase=%s). Excel may "
                        "be hung; per-call watchdog should fire at %ds.",
                        wid, elapsed, stalled_for,
                        (p.get("current_project") or "?"), phase,
                        600,  # DEFAULT_PER_CALL_TIMEOUT_SEC in direct_runner
                    )
                    self._stall_warned.add(wid)

        agg = _aggregate(
            payloads, self._wb_path,
            expected_worker_count=self._expected_workers,
        )
        atomic_write_json(self._out, agg)

    def stop(self, *, join_timeout: float = 5.0) -> None:
        self._stop_evt.set()
        self.join(timeout=join_timeout)
