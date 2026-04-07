"""dn38_solver.com.launcher — Subprocess launcher for COM worker.

Sends ALL projects as a single batch to one COM worker process.
The worker opens Excel once, solves all projects, then closes.
This mirrors the VBA approach and avoids the cold-start penalty per project.
"""
from __future__ import annotations

import logging
import subprocess
import sys
from pathlib import Path

import msgspec

from dn38_solver.types import SolveResult, SolveTask
from dn38_solver.com.cleanup import kill_orphan_excel

log = logging.getLogger(__name__)

_WORKER_PATH = Path(__file__).resolve().parent.parent.parent / "com_worker.py"


def run_worker_batch(
    workbook_path: str,
    tasks: list[SolveTask],
    *,
    original_f2: int = 1,
    timeout_sec: int = 600,
    gs_max_change: float = 0.00001,
    gs_max_iterations: int = 1000,
) -> dict:
    """Launch ONE COM worker subprocess for ALL projects.

    Returns the raw batch result dict with project_results list.
    """
    if not _WORKER_PATH.exists():
        return {
            "status": "error",
            "project_results": [],
            "duration_sec": 0.0,
            "saved_to": None,
            "error": f"COM worker not found: {_WORKER_PATH}",
        }

    # Build the batch payload (all tasks in one JSON)
    batch_payload = {
        "workbook_path": workbook_path,
        "tasks": [msgspec.to_builtins(t) for t in tasks],
        "saved_workbook_suffix": "_SOLVED",
        "gs_max_change": gs_max_change,
        "gs_max_iterations": gs_max_iterations,
        "original_f2": original_f2,
    }

    import json
    payload_json = json.dumps(batch_payload, default=str)

    n = len(tasks)
    log.info(
        "Launching COM worker for %d project(s) (timeout=%ds)",
        n, timeout_sec,
    )

    try:
        proc = subprocess.Popen(
            [sys.executable, str(_WORKER_PATH)],
            stdin=subprocess.PIPE,
            stdout=subprocess.PIPE,
            stderr=subprocess.PIPE,
            text=True,
            creationflags=subprocess.CREATE_NEW_PROCESS_GROUP,
        )

        stdout, stderr = proc.communicate(
            input=payload_json,
            timeout=timeout_sec,
        )

    except subprocess.TimeoutExpired:
        log.error("COM worker timed out after %ds", timeout_sec)
        proc.kill()
        kill_orphan_excel()
        return {
            "status": "timeout",
            "project_results": [],
            "duration_sec": float(timeout_sec),
            "saved_to": None,
            "error": f"Timed out after {timeout_sec}s",
        }
    except Exception as exc:
        log.error("COM worker launch failed: %s", exc)
        return {
            "status": "error",
            "project_results": [],
            "duration_sec": 0.0,
            "saved_to": None,
            "error": f"Launch failed: {exc}",
        }

    if proc.returncode != 0:
        err_msg = stderr.strip() or stdout.strip() or f"Exit code {proc.returncode}"
        log.error("COM worker exited with code %d: %s", proc.returncode, err_msg)
        return {
            "status": "error",
            "project_results": [],
            "duration_sec": 0.0,
            "saved_to": None,
            "error": err_msg,
        }

    try:
        return json.loads(stdout)
    except Exception as exc:
        log.error("Failed to parse COM worker output: %s", exc)
        return {
            "status": "error",
            "project_results": [],
            "duration_sec": 0.0,
            "saved_to": None,
            "error": f"Invalid JSON: {stdout[:200]}",
        }
