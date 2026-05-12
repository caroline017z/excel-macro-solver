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
import time
from pathlib import Path
from typing import Sequence

import msgspec

from dn38_solver.types import SolveTask

log = logging.getLogger(__name__)

# Hard cap so a user typo (`--workers 64`) doesn't open 64 Excel processes
# and consume all RAM. The Plan agent's analysis suggested cpu_count // 2
# as a sane upper bound; 8 is a portable conservative cap.
MAX_WORKERS = 8


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


def _kill_worker_excel_children(worker_pid: int) -> None:
    """Terminate any EXCEL.EXE under a worker PID. No-op if psutil missing."""
    from dn38_solver.com.cleanup import kill_excel_children
    with contextlib.suppress(Exception):
        kill_excel_children(worker_pid)


def run_parallel(
    workbook_path: str,
    tasks: list[SolveTask],
    *,
    workers: int,
    original_f2: int = 1,
    timeout_sec: int = 3600,
    use_chunked: bool = True,
    skip_output_recalc: bool = False,
    strip_sheets: tuple[str, ...] = (),
) -> dict:
    """Spawn `workers` subprocess workers and merge their results.

    Returns the same shape as `run_direct` so the orchestrator can branch
    on workers > 1 without other changes downstream:
        {status, project_results, duration_sec, saved_to, error,
         macro_used, open_time_sec, warmup_time_sec, solve_time_sec,
         read_time_sec, solver_heartbeat, validation}
    """
    n = max(1, min(workers, MAX_WORKERS))
    if n != workers:
        log.warning("Clamping workers from %d to %d (MAX_WORKERS)", workers, n)

    start = time.time()
    parent_tmp = Path(tempfile.mkdtemp(prefix="38dn_parallel_"))

    # Cap Excel threads per worker so N × Excel doesn't oversubscribe CPU.
    cpu = os.cpu_count() or 4
    threads_per_worker = max(1, cpu // n)
    log.info(
        "  Parallel mode: %d workers × ~%d Excel threads each (cpu_count=%d)",
        n, threads_per_worker, cpu,
    )

    partitions = _partition_round_robin(tasks, n)
    log.info(
        "  Task split (round-robin): %s",
        ", ".join(f"w{i}={len(p)}" for i, p in enumerate(partitions)),
    )

    # Spawn workers
    procs: list[tuple[int, subprocess.Popen[bytes], Path, Path]] = []
    for worker_id, slice_tasks in enumerate(partitions):
        if not slice_tasks:
            continue
        wdir = parent_tmp / f"w{worker_id}"
        wdir.mkdir(parents=True, exist_ok=True)
        config_path = wdir / "config.json"
        result_path = wdir / "result.json"
        output_path = wdir / f"{Path(workbook_path).stem}_SOLVED_w{worker_id}.xlsm"
        status_path = wdir / f"solver_status_w{worker_id}.json"

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
            "skip_output_recalc": bool(skip_output_recalc),
            "strip_sheets": list(strip_sheets),
            "excel_threads": threads_per_worker,
        }
        config_path.write_text(json.dumps(config), encoding="utf-8")

        cmd = [
            sys.executable, "-m", "dn38_solver.com.worker",
            str(config_path), str(result_path),
        ]
        # CREATE_NO_WINDOW = 0x08000000 — suppress console flash for each
        # worker subprocess. Windows-only flag; ignored on other OS.
        creationflags = 0
        if sys.platform == "win32":
            creationflags = 0x08000000
        log.info("  Spawning worker %d (%d projects)...", worker_id, len(slice_tasks))
        proc = subprocess.Popen(
            cmd,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            creationflags=creationflags,
        )
        procs.append((worker_id, proc, config_path, result_path))

    # Wait for all workers. Per-worker timeout is the parent-level timeout
    # since each worker handles only its slice; the cap should be generous
    # enough that no worker hits it before its slice's COM session does.
    worker_results: dict[int, dict] = {}
    parent_timeout = max(timeout_sec, 60)
    deadline = time.time() + parent_timeout + 60  # +60s for cleanup grace

    for worker_id, proc, config_path, result_path in procs:
        remaining = max(1, int(deadline - time.time()))
        try:
            stdout, stderr = proc.communicate(timeout=remaining)
        except subprocess.TimeoutExpired:
            log.warning(
                "  Worker %d timeout after %ds — terminating and reaping "
                "child Excel processes by PID", worker_id, remaining,
            )
            _kill_worker_excel_children(proc.pid)
            proc.terminate()
            with contextlib.suppress(subprocess.TimeoutExpired):
                proc.communicate(timeout=10)
            if proc.poll() is None:
                proc.kill()
            worker_results[worker_id] = {
                "status": "error",
                "error": f"Worker {worker_id} timed out",
                "project_results": [],
            }
            continue

        if stderr:
            # Forward worker stderr to parent log so users see worker output.
            for line in stderr.decode("utf-8", errors="replace").splitlines():
                log.info("  %s", line)

        if proc.returncode != 0:
            log.error("  Worker %d exited with code %d", worker_id, proc.returncode)
            _kill_worker_excel_children(proc.pid)
            try:
                worker_results[worker_id] = json.loads(
                    result_path.read_text(encoding="utf-8")
                )
            except Exception:
                worker_results[worker_id] = {
                    "status": "error",
                    "error": f"Worker {worker_id} exited non-zero, no result file",
                    "project_results": [],
                }
            continue

        try:
            worker_results[worker_id] = json.loads(
                result_path.read_text(encoding="utf-8")
            )
        except Exception as exc:
            log.error("  Could not read worker %d result: %s", worker_id, exc)
            worker_results[worker_id] = {
                "status": "error",
                "error": f"Could not read worker {worker_id} result: {exc}",
                "project_results": [],
            }

    # Aggregate per-project results in original task order. Build offset->result.
    by_offset: dict[int, dict] = {}
    worker_errors: list[str] = []
    for wid, wresult in worker_results.items():
        if wresult.get("error"):
            worker_errors.append(f"w{wid}: {wresult['error']}")
        for pr in wresult.get("project_results", []):
            # The project_results dicts already key into __SolverResults
            # by name; we need the offset to align back to the master
            # task ordering. Pull it from meta where the worker captured it.
            meta = pr.get("meta") or {}
            offset = meta.get("project_offset") or meta.get("offset")
            if offset is None:
                # Fall back to name lookup if offset wasn't recorded
                for t in tasks:
                    if t.project_name == pr.get("project_name"):
                        offset = t.project_offset
                        break
            if offset is not None:
                by_offset[int(offset)] = pr

    project_results: list[dict] = []
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
    # Strategy: pick worker 0's file as master, copy converged cells
    # (the per-project columns) from each other worker's file via
    # openpyxl. Cross-project portfolio aggregates may be stale per
    # Caroline's spec — Excel recalcs them on next interactive open.
    saved_to: str | None = None
    if procs:
        master_worker_id = procs[0][0]
        master_result = worker_results.get(master_worker_id, {})
        master_src = master_result.get("saved_to")
        if master_src and Path(master_src).exists():
            wb_path = Path(workbook_path)
            final_path = wb_path.parent / f"{wb_path.stem}_SOLVED{wb_path.suffix}"
            try:
                _merge_solved_workbooks(
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
                log.info("  Merged %d worker output(s) into %s",
                         len(worker_results), final_path)
            except Exception as merge_exc:
                log.warning(
                    "  Merge step failed (%s) — falling back to master "
                    "worker's output as-is. Per-project columns owned by "
                    "other workers will reflect their PRE-solve values.",
                    merge_exc,
                )
                # Fall back: just copy master_src to final_path. Caller
                # still has each worker's individual _SOLVED.xlsm in the
                # parent tmp_dir for forensic recovery.
                final_path = Path(workbook_path).parent / (
                    Path(workbook_path).stem + "_SOLVED" + Path(workbook_path).suffix
                )
                with contextlib.suppress(Exception):
                    shutil.copy2(master_src, final_path)
                    saved_to = str(final_path)

    # Cleanup worker temp dirs (keep parent tmp_dir only if a worker errored
    # — useful for forensics; otherwise drop everything).
    if not worker_errors:
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
    }


def _merge_solved_workbooks(
    *,
    master_src: Path,
    others: list[Path],
    final_path: Path,
    partitions: list[list[SolveTask]],
    master_worker_id: int,
) -> None:
    """Copy per-project converged column values from other workers' SOLVED
    workbooks into the master, then save to final_path.

    Only the project-column cells are copied (rows 31, 32, 33, 37, 38, 39
    on Project Inputs — the cached convergence outputs). The rest of the
    workbook in the master is left as-is; per Caroline's spec, portfolio
    aggregates may be stale and will refresh on next interactive open.

    keep_vba=True is critical — without it, openpyxl strips the macro
    project on save. Verified working on real xlsm round-trips during
    development of the speed-win strip-sheets feature.
    """
    import openpyxl

    # Convergence-output rows that get hard-stamped by VBA into the per-
    # project column cells (see SolveHeadless.bas lines around 678-679,
    # 968-969). Reading from VBA-written values means we don't need to
    # recalc anything in openpyxl.
    OUTPUT_ROWS = (31, 32, 33, 37, 38, 39)

    # Load master with keep_vba so the .xlsm round-trips with macros intact
    wb_master = openpyxl.load_workbook(str(master_src), keep_vba=True)
    ws_pi_master = wb_master["Project Inputs"] if "Project Inputs" in wb_master.sheetnames else None

    if ws_pi_master is None:
        # Nothing to merge into — just save master to final_path
        wb_master.save(str(final_path))
        return

    for other_path in others:
        if not other_path.exists():
            continue
        try:
            wb_other = openpyxl.load_workbook(str(other_path), keep_vba=True, data_only=True)
        except Exception:
            continue
        ws_pi_other = wb_other["Project Inputs"] if "Project Inputs" in wb_other.sheetnames else None
        if ws_pi_other is None:
            wb_other.close()
            continue

        # Figure out which worker this is and which task columns it owned
        # by matching against the file name's worker id suffix.
        wid = None
        stem = other_path.stem  # e.g., "<name>_SOLVED_w2"
        if "_w" in stem:
            try:
                wid = int(stem.rsplit("_w", 1)[1])
            except ValueError:
                wid = None
        if wid is None or wid >= len(partitions):
            wb_other.close()
            continue

        # Copy the convergence-output cells for each project this worker owned
        for task in partitions[wid]:
            # project_col_letter is the Excel column letter (e.g., "L"); convert
            # to numeric index for openpyxl.
            from openpyxl.utils import column_index_from_string
            col_idx = column_index_from_string(task.project_col_letter)
            for row in OUTPUT_ROWS:
                src_val = ws_pi_other.cell(row=row, column=col_idx).value
                ws_pi_master.cell(row=row, column=col_idx).value = src_val

        wb_other.close()

    wb_master.save(str(final_path))
    wb_master.close()
