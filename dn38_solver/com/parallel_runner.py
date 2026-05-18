"""dn38_solver.com.parallel_runner — Spawn N worker processes for parallel solves.

Each worker runs in its own subprocess with its own Excel COM session, its
own temp workbook copy, and a round-robin slice of the project list.
Parent collects per-worker results and merges them into a single
`<workbook>_SOLVED.xlsm` next to the source.

Why a separate parallel runner instead of an N=N>1 path inside run_direct:
- run_direct holds an Excel COM apartment; spawning subprocesses from
  inside a COM-using process is messy and risks hidden EXCEL.EXE leaks
  when the parent crashes.
- Parent stays COM-free — only worker processes touch Excel — so PID
  tracking, timeout enforcement, and cleanup all work via stdlib subprocess.
- Each worker reuses run_direct unchanged. No code duplication.

Caroline's correctness contract (Issue #8 + her clarification):
- Per-project NPP / Dev Fee / DSCR / FMV / Live IRR / Appraisal IRR must
  match single-worker output to ~1e-4.
- Portfolio-level aggregates (Portfolio sheet, cross-column sums) may be
  stale in the merged _SOLVED.xlsm; Excel will recalc them on next
  interactive open since the underlying project columns are correct.
- Critical sheets per spec (Dashboard, Table, PT Returns, NPP Calc,
  Appraisal, Perm Debt, Tax Equity, CL, Project Inputs) recalc inside
  each worker's own Excel session for its assigned projects.
"""
from __future__ import annotations

import contextlib
import json
import logging
import os
import shutil
import subprocess
import sys
import tempfile
import threading
import time
from pathlib import Path
from typing import Any, NamedTuple, Sequence

import msgspec

from dn38_solver.com.cleanup import (
    kill_excel_children,
    kill_excel_children_for_handle,
)
from dn38_solver.com.direct_runner import STATUS_FILE
from dn38_solver.com.status_aggregator import StatusAggregator
from dn38_solver.merge import merge_via_openpyxl, merge_via_vba_fallback
from dn38_solver.types import SolveTask
from dn38_solver.validation.post_merge import verify_merged_file


class WorkerProc(NamedTuple):
    """A spawned worker subprocess and the artifacts the parent needs to
    track it. Replaces the historical 7-tuple where every unpack site had
    to know the positional order of `(wid, popen, cfg, res, stdout, stderr,
    handle)` — adding a field meant updating every unpack with the right
    underscores. Named fields make the structure self-documenting and
    future field additions safe.

    `handle` is the `psutil.Process` captured at spawn time so cleanup
    after worker exit defeats Windows PID reuse races (see
    `_capture_proc_handle`). None when psutil is unavailable.
    """
    worker_id: int
    popen: subprocess.Popen
    config_path: Path
    result_path: Path
    stdout_path: Path
    stderr_path: Path
    handle: Any = None  # psutil.Process | None — string-loose to keep psutil import lazy

log = logging.getLogger(__name__)

# Hard cap so a user typo (`--workers 64`) doesn't open 64 Excel processes
# and consume all RAM. The Plan agent's analysis suggested cpu_count // 2
# as a sane upper bound; 8 is a portable conservative cap.
MAX_WORKERS = 8


class _WorkerLogTailer(threading.Thread):
    """Background thread that tails each worker's stderr.log and forwards
    new lines to the parent log so the terminal isn't silent during long
    cold solves.

    Tracks per-file byte offsets so each line is forwarded exactly once.
    Lines already carry the worker id via the `[wN]` prefix the worker
    sets in `logging.basicConfig(format=...)`, so we don't re-decorate.
    """

    def __init__(
        self,
        *,
        sources: list[tuple[int, Path]],
        poll_interval: float = 2.0,
    ) -> None:
        super().__init__(daemon=True, name="WorkerLogTailer")
        self._sources = list(sources)
        self._interval = poll_interval
        self._offsets: dict[Path, int] = {p: 0 for _, p in sources}
        self._stop_evt = threading.Event()

    def run(self) -> None:
        while not self._stop_evt.is_set():
            self._poll_once()
            if self._stop_evt.wait(timeout=self._interval):
                break
        # Final flush so we don't miss the last few lines after stop
        self._poll_once()

    def _poll_once(self) -> None:
        for _wid, path in self._sources:
            try:
                size = path.stat().st_size
            except FileNotFoundError:
                continue
            offset = self._offsets[path]
            if size <= offset:
                continue
            try:
                with open(path, "rb") as fh:
                    fh.seek(offset)
                    chunk = fh.read(size - offset)
            except OSError:
                continue
            # Decode UTF-8 then strip non-ASCII so the parent's stdout writer
            # (cp1252 on Windows) can't crash mid-line on a stray Unicode char
            # that would silence the rest of the worker's output. Replacement
            # character (U+FFFD) is the specific failure mode from 2026-05-15.
            text = (
                chunk.decode("utf-8", errors="replace")
                .encode("ascii", errors="replace")
                .decode("ascii")
            )
            advanced = False
            for line in text.splitlines():
                if line.strip():
                    try:
                        log.info("  %s", line)
                    except Exception:
                        # Belt-and-suspenders: even if a logging handler
                        # somehow still chokes, don't drop the offset advance.
                        pass
                advanced = True
            # Advance offset only AFTER the loop completes — if anything in
            # the loop raised, we want the next poll to retry the same bytes
            # rather than skip them silently.
            if advanced:
                self._offsets[path] = size

    def stop(self, *, join_timeout: float = 5.0) -> None:
        self._stop_evt.set()
        self.join(timeout=join_timeout)


def _partition_round_robin(
    tasks: Sequence[SolveTask], n_workers: int
) -> list[list[SolveTask]]:
    """Distribute tasks round-robin so warm/cold projects spread evenly.

    Block partitioning (chunks of consecutive indices) tends to put slow
    cold-mode projects in the same chunk, leaving one worker as the
    long-tail bottleneck. Round-robin avoids that.
    """
    slices: list[list[SolveTask]] = [[] for _ in range(n_workers)]
    for i, t in enumerate(tasks):
        slices[i % n_workers].append(t)
    return slices


def _kill_worker_excel_children(worker_pid: int, handle=None) -> None:
    """Terminate any EXCEL.EXE under a worker. No-op if psutil missing.

    Prefers a `psutil.Process` handle captured at spawn time (defeats PID
    reuse races on Windows after the worker has exited); falls back to
    PID-based lookup when no handle is available.
    """
    with contextlib.suppress(Exception):
        if handle is not None:
            kill_excel_children_for_handle(handle)
        else:
            kill_excel_children(worker_pid)


def _sweep_old_parent_tmps() -> None:
    """Delete `38dn_parallel_*` dirs in %TEMP% older than the retention window.

    Errored parallel runs are preserved for forensics, but Windows %TEMP%
    is rarely auto-cleaned and each preserved dir holds N copies of the
    workbook (~50–200MB each). After a few weeks of failures this silently
    consumes tens of GB. Sweep on every run-start so the disk cost is
    bounded by the retention window.

    Retention defaults to 7 days; override with DN38_TMP_RETENTION_DAYS=N.
    Failures are swallowed silently — disk hygiene must never block the
    user's actual solve.
    """
    try:
        days_raw = os.environ.get("DN38_TMP_RETENTION_DAYS", "7")
        days = max(0, int(days_raw))
    except ValueError:
        days = 7
    if days == 0:
        return  # 0 = retain forever (escape hatch)
    cutoff = time.time() - (days * 86400)
    swept = 0
    swept_bytes = 0
    try:
        tmpdir = Path(tempfile.gettempdir())
    except Exception:
        return
    inflight_window = time.time() - 3600  # files written in last hour
    for entry in tmpdir.glob("38dn_parallel_*"):
        try:
            mtime = entry.stat().st_mtime
        except OSError:
            continue
        if mtime >= cutoff:
            continue
        # Defense against stomping on a long-running solve. The dir's
        # OWN mtime can age past the retention window even while a worker
        # is still actively writing inside it (the directory entry doesn't
        # touch on every child write). Walk one level into the worker
        # subdirs and look for a recent file write — if any worker has
        # written in the last hour, the solve is still in flight and we
        # leave the dir alone.
        try:
            recently_active = any(
                f.is_file() and f.stat().st_mtime >= inflight_window
                for f in entry.rglob("*")
            )
        except OSError:
            recently_active = True  # err on the side of NOT deleting
        if recently_active:
            log.debug(
                "  Sweep: skipping %s (recent worker activity within 1h)",
                entry,
            )
            continue
        # Best-effort byte count for the log line; cheap recursive glob
        # and we already know the entry is going to be deleted.
        try:
            for f in entry.rglob("*"):
                if f.is_file():
                    swept_bytes += f.stat().st_size
        except OSError:
            pass
        with contextlib.suppress(OSError):
            shutil.rmtree(entry)
            swept += 1
    if swept:
        log.info(
            "  Swept %d old parent_tmp dir(s) (~%d MB freed, retention=%dd)",
            swept, swept_bytes // (1024 * 1024), days,
        )


def _capture_proc_handle(pid: int):
    """Best-effort psutil.Process at spawn time. Returns None if psutil
    missing or process already gone (very tight race we accept silently)."""
    try:
        import psutil
    except ImportError:
        return None
    try:
        return psutil.Process(pid)
    except (psutil.NoSuchProcess, psutil.AccessDenied):
        # NoSuchProcess: process exited between Popen and our handle
        # capture (very tight race). AccessDenied: rare on Windows for a
        # process we own, but possible under certain group policies.
        # Both are silently-acceptable — cleanup falls back to bare-PID
        # kill_excel_children which handles the missing-process case.
        return None


def run_parallel(
    workbook_path: str,
    tasks: list[SolveTask],
    *,
    workers: int,
    original_f2: int = 1,
    timeout_sec: int = 3600,
    use_chunked: bool = True,
    save_solved: bool = True,
    skip_output_recalc: bool = False,
    strip_sheets: tuple[str, ...] = (),
    excel_threads_per_worker: int | None = None,
) -> dict:
    """Spawn `workers` subprocess workers and merge their results.

    Returns the same shape as `run_direct` so the orchestrator can branch
    on workers > 1 without other changes downstream:
        {status, project_results, duration_sec, saved_to, error,
         macro_used, open_time_sec, warmup_time_sec, solve_time_sec,
         read_time_sec, solver_heartbeat, validation,
         merge_path, estimated_sequential_sec}

    `merge_path` is one of "openpyxl" | "vba_fallback" | "copy_master" |
    None and tells the orchestrator how authoritative the merged file is.
    `estimated_sequential_sec` is the sum of per-project solve seconds
    (from each worker's __SolverResults timings), so the caller can
    print real speedup instead of guessing.
    """
    n = max(1, min(workers, MAX_WORKERS))
    if n != workers:
        log.warning("Clamping workers from %d to %d (MAX_WORKERS)", workers, n)

    start = time.time()
    _sweep_old_parent_tmps()
    parent_tmp = Path(tempfile.mkdtemp(prefix="38dn_parallel_"))

    # Tracked across the try/finally so cleanup can always tear them down,
    # even if Ctrl+C lands mid-spawn or mid-wait. Without this, an interrupt
    # during a 20-minute solve leaves N hidden Excel processes running until
    # the user reboots.
    procs: list[WorkerProc] = []
    aggregator = None
    progress = None
    saved_to: str | None = None
    merge_path: str | None = None
    worker_errors: list[str] = []
    worker_results: dict[int, dict] = {}
    project_results: list[dict] = []

    # Cap Excel threads per worker so N × Excel doesn't oversubscribe CPU.
    # Validated empirically on SMP WalkTEST: cpu_count // n_workers (24/2=12)
    # produced ZERO speedup vs single-worker because the 24 calc threads
    # competed with the OS, pywin32, the aggregator, and the live-progress
    # tailer for the same 24 cores. Excel's multi-threaded recalc has
    # sharply diminishing returns past ~4 threads anyway, so the default
    # is a conservative cap that leaves headroom for coordination overhead.
    # User can override via the CLI flag if their workload differs.
    #
    # Semantics of `excel_threads_per_worker`:
    #   None    -> use the default 1/3-of-cores cap (per-worker)
    #   0       -> no cap; let Excel use its built-in default (cpu_count)
    #   1..32   -> set Excel's ThreadCount to exactly that value
    #   >32     -> clamped to 32 (sanity cap; Excel's own implementation
    #              caps at 1024 but anything past 32 is wasted on the kind
    #              of dependency-chain heavy recalc this model does)
    cpu = os.cpu_count() or 4
    SANITY_MAX_THREADS = 32
    if excel_threads_per_worker is None:
        # Default: 1/3 of available cores per worker, min 2, max 4.
        # On a 24-core box with 2 workers: 4 + 4 = 8 calc threads, leaving
        # 16 cores for OS / coordination / file I/O.
        threads_per_worker: int | None = max(2, min(4, cpu // (n * 3)))
    elif excel_threads_per_worker == 0:
        threads_per_worker = None  # signals "no cap" to direct_runner
    elif excel_threads_per_worker < 0:
        # Refuse to silently coerce a typo. Without this, `--excel-threads-per-worker -1`
        # would fall through `> 0` checks in direct_runner and apply no
        # cap with no log line, leaving the user thinking their flag took
        # effect. Fail loudly.
        raise ValueError(
            f"excel_threads_per_worker must be >= 0 (got {excel_threads_per_worker}); "
            "use 0 for 'no cap, Excel default' or a positive integer to set the cap"
        )
    else:
        if excel_threads_per_worker > SANITY_MAX_THREADS:
            log.warning(
                "  Clamping excel_threads_per_worker from %d to %d (sanity cap)",
                excel_threads_per_worker, SANITY_MAX_THREADS,
            )
            threads_per_worker = SANITY_MAX_THREADS
        else:
            threads_per_worker = int(excel_threads_per_worker)
    threads_label = (
        "Excel default (no cap)" if threads_per_worker is None
        else f"{threads_per_worker} Excel threads each"
    )
    log.info(
        "  Parallel mode: %d workers × %s (cpu_count=%d)",
        n, threads_label, cpu,
    )

    partitions = _partition_round_robin(tasks, n)
    log.info(
        "  Task split (round-robin): %s",
        ", ".join(f"w{i}={len(p)}" for i, p in enumerate(partitions)),
    )

    try:
        # Spawn workers. Each worker's stdout/stderr is redirected to a file
        # in its temp dir, NOT to a subprocess.PIPE — Windows anonymous pipes
        # have a ~65KB buffer, and a long cold solve can easily emit more than
        # that to stderr, deadlocking the worker (writes to a full pipe block
        # until the parent reads, but proc.communicate() is what reads). Log
        # files round-trip cleanly and let the parent tail them for live
        # progress without any blocking risk.
        worker_status_paths: list[Path] = []
        for worker_id, slice_tasks in enumerate(partitions):
            if not slice_tasks:
                continue
            wdir = parent_tmp / f"w{worker_id}"
            wdir.mkdir(parents=True, exist_ok=True)
            config_path = wdir / "config.json"
            result_path = wdir / "result.json"
            output_path = wdir / f"{Path(workbook_path).stem}_SOLVED_w{worker_id}.xlsm"
            status_path = wdir / f"solver_status_w{worker_id}.json"
            worker_status_paths.append(status_path)

            tasks_json = msgspec.json.encode(slice_tasks).decode("utf-8")
            config = {
                "workbook_path": workbook_path,
                "tasks_json": tasks_json,
                "output_path": str(output_path),
                "status_path": str(status_path),
                "worker_id": worker_id,
                "original_f2": int(original_f2),
                "timeout_sec": int(timeout_sec),
                "use_chunked": bool(use_chunked),
                "save_solved": bool(save_solved),
                "skip_output_recalc": bool(skip_output_recalc),
                "strip_sheets": list(strip_sheets),
                "excel_threads": threads_per_worker,
            }
            config_path.write_text(json.dumps(config), encoding="utf-8")

            cmd = [
                sys.executable, "-u", "-m", "dn38_solver.com.worker",
                str(config_path), str(result_path),
            ]
            # CREATE_NO_WINDOW = 0x08000000 — suppress console flash for each
            # worker subprocess. Windows-only flag; ignored on other OS.
            creationflags = 0
            if sys.platform == "win32":
                creationflags = 0x08000000
            log.info("  Spawning worker %d (%d projects)...", worker_id, len(slice_tasks))
            # Open log files unbuffered (binary 'wb' with no flush deferral)
            # so the parent's tailing thread sees worker writes immediately.
            # `with` ensures fh close on the error path — if Popen raises
            # (cmd missing, OSError on spawn) without a context-manager,
            # both handles would leak.
            stdout_path = wdir / "stdout.log"
            stderr_path = wdir / "stderr.log"
            with (
                open(stdout_path, "wb", buffering=0) as stdout_fh,
                open(stderr_path, "wb", buffering=0) as stderr_fh,
            ):
                # The child inherits these file descriptors; the parent's
                # copies are then closed by the `with` exit, leaving only
                # the child holding them. No inode pinning either way.
                proc = subprocess.Popen(
                    cmd,
                    stdout=stdout_fh,
                    stderr=stderr_fh,
                    creationflags=creationflags,
                )
            # Capture a psutil.Process handle NOW, while we know the PID is
            # still our worker. Used by cleanup paths after the worker has
            # exited — psutil's create_time check defeats Windows PID reuse,
            # so we don't accidentally reap a stranger's Excel children.
            proc_handle = _capture_proc_handle(proc.pid)
            procs.append(WorkerProc(
                worker_id=worker_id,
                popen=proc,
                config_path=config_path,
                result_path=result_path,
                stdout_path=stdout_path,
                stderr_path=stderr_path,
                handle=proc_handle,
            ))

        # Start the status aggregator. Each worker writes its own status JSON;
        # the aggregator merges them into the canonical `solver_status.json`
        # the Streamlit tracker reads, so dashboards work without changes.
        # Pass expected_worker_count explicitly so the aggregator's
        # "don't roll up to complete until all workers have reported"
        # guard is wired correctly. Defaulting via len(worker_status_paths)
        # works today by coincidence (skipped empty slices => same length)
        # but a future caller passing a different path set would silently
        # break the gate.
        aggregator = StatusAggregator(
            worker_status_paths=worker_status_paths,
            output_path=STATUS_FILE,
            workbook_path=workbook_path,
            poll_interval=1.0,
            expected_worker_count=len(procs),
        )
        aggregator.start()

        # Live-progress thread: tails each worker's stderr.log and forwards
        # new lines to the parent log so the terminal isn't silent during
        # a 10-20 minute cold solve. Without this the user sees nothing
        # between "Spawning worker N" and "Worker N completed".
        progress = _WorkerLogTailer(
            sources=[(wp.worker_id, wp.stderr_path) for wp in procs],
            poll_interval=2.0,
        )
        progress.start()

        # Wait for all workers. Per-worker timeout is the parent-level timeout
        # since each worker handles only its slice; the cap should be generous
        # enough that no worker hits it before its slice's COM session does.
        parent_timeout = max(timeout_sec, 60)
        deadline = time.time() + parent_timeout + 60  # +60s for cleanup grace

        for wp in procs:
            remaining = max(1, int(deadline - time.time()))
            try:
                # stdout/stderr are redirected to files, so communicate()
                # blocks only on the process exit — no PIPE buffer to drain.
                wp.popen.communicate(timeout=remaining)
            except subprocess.TimeoutExpired:
                log.warning(
                    "  Worker %d timeout after %ds — terminating and reaping "
                    "child Excel processes by PID", wp.worker_id, remaining,
                )
                _kill_worker_excel_children(wp.popen.pid, handle=wp.handle)
                wp.popen.terminate()
                with contextlib.suppress(subprocess.TimeoutExpired):
                    wp.popen.communicate(timeout=10)
                if wp.popen.poll() is None:
                    wp.popen.kill()
                worker_results[wp.worker_id] = {
                    "status": "error",
                    "error": f"Worker {wp.worker_id} timed out",
                    "project_results": [],
                }
                continue

            # Worker stderr is now in stderr_path and the live tailer has
            # already forwarded it during the run. Drain any final lines
            # for completeness.
            with contextlib.suppress(OSError):
                tail = wp.stderr_path.read_text(encoding="utf-8", errors="replace")
                # Only forward lines the tailer hasn't already printed;
                # tailer tracks per-file byte offsets so we just print the
                # recent tail as INFO if the worker errored, otherwise the
                # lines were already streamed live.
                if wp.popen.returncode != 0 and tail.strip():
                    log.info("  --- worker %d stderr tail ---", wp.worker_id)
                    for line in tail.splitlines()[-30:]:
                        log.info("  %s", line)

            if wp.popen.returncode != 0:
                log.error(
                    "  Worker %d exited with code %d",
                    wp.worker_id, wp.popen.returncode,
                )
                # Use the pinned handle, not the bare PID — the worker has
                # already exited so the OS may have reused its PID.
                _kill_worker_excel_children(wp.popen.pid, handle=wp.handle)
                try:
                    worker_results[wp.worker_id] = json.loads(
                        wp.result_path.read_text(encoding="utf-8")
                    )
                except (OSError, json.JSONDecodeError):
                    worker_results[wp.worker_id] = {
                        "status": "error",
                        "error": f"Worker {wp.worker_id} exited non-zero, no result file",
                        "project_results": [],
                    }
                continue

            try:
                worker_results[wp.worker_id] = json.loads(
                    wp.result_path.read_text(encoding="utf-8")
                )
            except (OSError, json.JSONDecodeError) as exc:
                log.error("  Could not read worker %d result: %s", wp.worker_id, exc)
                worker_results[wp.worker_id] = {
                    "status": "error",
                    "error": f"Could not read worker {wp.worker_id} result: {exc}",
                    "project_results": [],
                }

        # Aggregate per-project results in original task order.
        by_offset: dict[int, dict] = {}
        for wid, wresult in worker_results.items():
            if wresult.get("error"):
                worker_errors.append(f"w{wid}: {wresult['error']}")
            for pr in wresult.get("project_results", []):
                # Use the top-level project_offset stamped by direct_runner.
                # Falling back to name-based lookup is unsafe — portfolios
                # can have duplicate project names across LLCs.
                offset = pr.get("project_offset")
                if offset is None:
                    meta = pr.get("meta") or {}
                    offset = meta.get("project_offset") or meta.get("offset")
                if offset is None:
                    log.warning(
                        "Worker %d returned a project_result with no offset "
                        "key; skipping rather than name-matching (unsafe with "
                        "duplicate names). project_name=%r",
                        wid, pr.get("project_name"),
                    )
                    continue
                by_offset[int(offset)] = pr

        for t in tasks:
            pr = by_offset.get(t.project_offset)
            if pr is None:
                project_results.append({
                    "project_name": t.project_name,
                    "status": "not_attempted",
                    "solved_values": {},
                    "iterations_used": 0,
                    "duration_sec": 0,
                    "meta": {},
                })
            else:
                project_results.append(pr)

        # Merge the worker _SOLVED.xlsm files into one canonical output.
        # Strategy: pick a successfully-converged worker as master, copy
        # converged cells (the per-project columns) from each other worker's
        # file via openpyxl. Cross-project portfolio aggregates may be stale
        # per Caroline's spec — Excel recalcs them on next interactive open.
        # Prefer the lowest-id worker whose status is "converged" AND who
        # actually saved a file. Falling back to procs[0] regardless of
        # outcome risks merging on top of an errored worker's empty/corrupt
        # output.
        #
        # When save_solved=False (typically validation runs that just want
        # the per-project numbers in memory), no worker wrote a file so
        # candidate_masters comes back empty and the merge is skipped via
        # the existing `else:` branch below — same code path as "all
        # workers errored." Logged distinctly so the difference between
        # "no merge by request" vs "no merge because everything broke"
        # is visible.
        if not save_solved:
            log.info(
                "  save_solved=False — skipping per-worker file merge "
                "(no _SOLVED.xlsm produced)"
            )
        candidate_masters = sorted(
            (wid for wid, r in worker_results.items()
             if r.get("status") == "converged" and r.get("saved_to")
             and Path(r["saved_to"]).exists()),
        )
        if candidate_masters:
            master_worker_id = candidate_masters[0]
            master_result = worker_results.get(master_worker_id, {})
            master_src = master_result.get("saved_to")
            if master_src and Path(master_src).exists():
                wb_path = Path(workbook_path)
                final_path = wb_path.parent / f"{wb_path.stem}_SOLVED{wb_path.suffix}"
                try:
                    merge_via_openpyxl(
                        master_src=Path(master_src),
                        others=[
                            Path(worker_results[wid].get("saved_to", ""))
                            for wid in worker_results
                            if wid != master_worker_id
                            and worker_results[wid].get("saved_to")
                        ],
                        final_path=final_path,
                        partitions=partitions,
                        master_worker_id=master_worker_id,
                    )
                    saved_to = str(final_path)
                    merge_path = "openpyxl"
                    log.info("  Merged %d worker output(s) into %s",
                             len(worker_results), final_path)
                except Exception as merge_exc:
                    # openpyxl merge failed — most likely keep_vba=True did
                    # not survive the round-trip on this workbook's macro
                    # project. Fall back to a VBA-side merge that uses Excel
                    # COM (which handles .xlsm natively) to stamp converged
                    # column values from peer workers into the master.
                    log.warning(
                        "  openpyxl merge failed (%s) — trying VBA-helper "
                        "fallback via Excel COM", merge_exc,
                    )
                    others_paths = [
                        Path(worker_results[wid].get("saved_to", ""))
                        for wid in worker_results
                        if wid != master_worker_id
                        and worker_results[wid].get("saved_to")
                    ]
                    try:
                        merge_via_vba_fallback(
                            master_src=Path(master_src),
                            others=others_paths,
                            final_path=final_path,
                            partitions=partitions,
                            worker_results=worker_results,
                            master_worker_id=master_worker_id,
                        )
                        saved_to = str(final_path)
                        merge_path = "vba_fallback"
                        log.info(
                            "  VBA-helper fallback succeeded: merged %d output(s) into %s",
                            len(worker_results), final_path,
                        )
                    except Exception as vba_exc:
                        log.warning(
                            "  VBA-helper fallback also failed (%s) — copying "
                            "master worker's output as-is. Per-project columns "
                            "owned by other workers will reflect their PRE-solve "
                            "values; consult the per-worker _SOLVED.xlsm files "
                            "in %s for forensic recovery.",
                            vba_exc, parent_tmp,
                        )
                        with contextlib.suppress(Exception):
                            shutil.copy2(master_src, final_path)
                            saved_to = str(final_path)
                            merge_path = "copy_master"
                        # Force run-level error so the user knows the merged
                        # file is not authoritative and so parent_tmp is
                        # preserved for forensics (see cleanup gate below).
                        worker_errors.append(
                            "merge fell back to master-only — per-project "
                            "columns owned by non-master workers reflect "
                            "PRE-solve values, not converged values"
                        )

                # Post-merge sanity gate. Re-open the merged file and
                # confirm the hard-stamped convergence cells match what
                # each worker reported. If they don't, the merge silently
                # corrupted data and we must surface that as an error so
                # the user doesn't ship a file with wrong NPP / FMV /
                # Dev Fee in some columns. Skip the gate when we already
                # know the merge fell back to master-only — its mismatches
                # are expected and would just spam the log.
                if saved_to and merge_path in ("openpyxl", "vba_fallback"):
                    mismatches = verify_merged_file(
                        final_path=Path(saved_to),
                        worker_results=worker_results,
                        partitions=partitions,
                    )
                    if mismatches:
                        log.error(
                            "  [FAIL] POST-MERGE VERIFICATION FAILED: "
                            "%d cell mismatch(es) between merged file and "
                            "worker reports. DO NOT SHIP %s -- use the "
                            "per-worker outputs in %s instead.",
                            len(mismatches), saved_to, parent_tmp,
                        )
                        for m in mismatches[:10]:
                            log.error("    %s", m)
                        if len(mismatches) > 10:
                            log.error("    ... (%d more)", len(mismatches) - 10)
                        worker_errors.append(
                            f"merged file failed sanity gate "
                            f"({len(mismatches)} cell mismatch(es)) -- "
                            f"per-worker outputs in {parent_tmp} are authoritative"
                        )
                    else:
                        log.info("  [OK] Post-merge verification passed")
        elif save_solved:
            # No worker produced a saved, converged output — nothing to merge.
            # Surface this clearly so the orchestrator marks the run as error
            # and keeps parent_tmp for forensics. Skip this branch entirely
            # when save_solved=False — empty candidate_masters is expected
            # in that case (workers were told not to save), not a failure.
            log.error(
                "  No worker produced a converged _SOLVED.xlsm — skipping merge. "
                "Worker outcomes: %s",
                {wid: r.get("status") for wid, r in worker_results.items()},
            )
            worker_errors.append(
                "no converged worker output to merge — all workers errored or "
                "produced no saved file"
            )

    finally:
        # Always tear down workers and the helper threads, including on
        # KeyboardInterrupt or any unhandled exception during merge. Order
        # matters: kill workers first so their final stderr writes have
        # a chance to land before the tailer stops.
        for wp in procs:
            if wp.popen.poll() is None:
                log.warning(
                    "  Cleanup: terminating worker %d (pid=%d) and its "
                    "Excel children", wp.worker_id, wp.popen.pid,
                )
                _kill_worker_excel_children(wp.popen.pid, handle=wp.handle)
                with contextlib.suppress(Exception):
                    wp.popen.terminate()
                with contextlib.suppress(Exception):
                    wp.popen.wait(timeout=5)
                if wp.popen.poll() is None:
                    with contextlib.suppress(Exception):
                        wp.popen.kill()
                # Re-reap after kill — TerminateProcess may leave a brief
                # window where Excel children get re-parented (typically
                # to init / Session Manager) before they're cleaned up.
                # Second pass catches them; safe to call when first call
                # already cleaned everything (returns 0).
                _kill_worker_excel_children(wp.popen.pid, handle=wp.handle)
        # Guard against the "constructed but never started" state — if
        # an exception fired between StatusAggregator(...) and .start(),
        # the thread object exists but has no underlying OS thread, and
        # Thread.join() raises RuntimeError. `Thread.ident` is None until
        # start() runs and stays set after the thread exits, so it's the
        # right public-API check (is_alive() returns False post-exit and
        # would skip a legitimate stop+join).
        if progress is not None and progress.ident is not None:
            with contextlib.suppress(Exception):
                progress.stop(join_timeout=5.0)
        if aggregator is not None and aggregator.ident is not None:
            with contextlib.suppress(Exception):
                aggregator.stop(join_timeout=5.0)

    # Cleanup worker temp dirs (keep parent tmp_dir only if a worker errored
    # — useful for forensics; otherwise drop everything).
    # DN38_KEEP_WORKER_TMP=1 forces retention regardless (debugging hatch).
    keep_tmp = bool(worker_errors) or os.environ.get("DN38_KEEP_WORKER_TMP") == "1"
    if not keep_tmp:
        with contextlib.suppress(OSError):
            shutil.rmtree(parent_tmp)
    else:
        log.info(
            "  Worker errors detected — preserving %s for forensics",
            parent_tmp,
        )

    total = time.time() - start
    batch_status = "error" if worker_errors else "converged"
    error_msg = "; ".join(worker_errors) if worker_errors else None

    # Compose the run_direct-compatible result. Per-stage timings are
    # summed across workers (open_time, solve_time, read_time); the parent
    # didn't directly measure them so we accumulate worker reports.
    def _sum_stage(key: str) -> float:
        return sum(
            float(w.get(key, 0) or 0) for w in worker_results.values()
        )

    # Estimated sequential time = sum of every attempted project's
    # solve_seconds (from VBA's __SolverResults!M, surfaced as
    # meta["solve_seconds"]). Comparing the parallel wall time to this
    # gives the actual speedup, which is more useful than "elapsed vs.
    # previous run" since project mix differs across runs.
    # Skip not_attempted explicitly: their meta is empty so they'd
    # contribute 0, but being explicit keeps the math obvious.
    estimated_sequential = 0.0
    for wresult in worker_results.values():
        for pr in wresult.get("project_results", []):
            if pr.get("status") == "not_attempted":
                continue
            meta = pr.get("meta") or {}
            try:
                estimated_sequential += float(meta.get("solve_seconds") or 0.0)
            except (TypeError, ValueError):
                pass

    return {
        "status": batch_status,
        "project_results": project_results,
        "duration_sec": round(total, 2),
        "saved_to": saved_to,
        "error": error_msg,
        "macro_used": "SolveHeadless",
        "open_time_sec": round(_sum_stage("open_time_sec"), 2),
        "warmup_time_sec": round(_sum_stage("warmup_time_sec"), 2),
        "solve_time_sec": round(_sum_stage("solve_time_sec"), 2),
        "read_time_sec": round(_sum_stage("read_time_sec"), 2),
        "solver_heartbeat": None,
        "validation": None,
        "workers_used": n,
        "merge_path": merge_path,
        "estimated_sequential_sec": round(estimated_sequential, 2),
    }

