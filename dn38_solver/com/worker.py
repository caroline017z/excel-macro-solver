"""dn38_solver.com.worker — Subprocess entry point for parallel solves.

Invoked by `dn38_solver.com.parallel_runner` as one subprocess per worker:

    python -m dn38_solver.com.worker <config.json> <result.json>

The worker reads its full task configuration from `config.json`, runs the
existing `run_direct` flow (its own Excel COM session, own temp dir, own
slice of projects, own _SOLVED.xlsm output path, own status file), and
writes the result dict to `result.json` on success or exits non-zero on
failure.

This keeps the parent orchestrator out of the COM apartment — only the
worker processes hold Excel instances, so the parent can spawn / kill /
monitor them without any of its own pythoncom state to manage.

Config schema (JSON object):
    {
        "workbook_path": str,           # absolute path; worker copies this
        "tasks_json": str,              # msgspec-serialized SolveTask list
        "output_path": str,             # where worker writes its _SOLVED.xlsm
        "status_path": str,             # where worker writes its status JSON
        "worker_id": int,               # for logging only
        "original_f2": int,
        "timeout_sec": int,
        "use_chunked": bool,
        "skip_output_recalc": bool,
        "strip_sheets": list[str],
        "excel_threads": int | null,
    }

Result schema (JSON object): the full dict returned by run_direct, plus
`worker_id`.
"""
from __future__ import annotations

import json
import logging
import sys
import traceback
from pathlib import Path

import msgspec

from dn38_solver.com.direct_runner import run_direct
from dn38_solver.types import SolveTask


def main() -> int:
    if len(sys.argv) != 3:
        sys.stderr.write(
            "Usage: python -m dn38_solver.com.worker <config.json> <result.json>\n"
        )
        return 2

    config_path = Path(sys.argv[1])
    result_path = Path(sys.argv[2])

    try:
        config = json.loads(config_path.read_text(encoding="utf-8"))
    except Exception as exc:
        sys.stderr.write(f"Worker failed to read config: {exc}\n")
        return 2

    worker_id = config.get("worker_id", 0)
    logging.basicConfig(
        level=logging.INFO,
        format=f"[w{worker_id}] %(message)s",
        stream=sys.stderr,
    )
    log = logging.getLogger(__name__)
    log.info("Worker %d starting on %d task(s)", worker_id, len(config.get("tasks_json", "")))

    try:
        tasks_bytes = config["tasks_json"].encode("utf-8")
        tasks = msgspec.json.decode(tasks_bytes, type=list[SolveTask])

        result = run_direct(
            workbook_path=config["workbook_path"],
            tasks=tasks,
            original_f2=int(config.get("original_f2", 1)),
            timeout_sec=int(config.get("timeout_sec", 3600)),
            use_chunked=bool(config.get("use_chunked", True)),
            checkpoint_callback=None,  # parent does checkpoint aggregation
            save_solved=True,
            skip_output_recalc=bool(config.get("skip_output_recalc", False)),
            strip_sheets=tuple(config.get("strip_sheets", [])),
            output_path=config["output_path"],
            status_path=Path(config["status_path"]),
            excel_threads=config.get("excel_threads"),
        )
        result["worker_id"] = worker_id

        result_path.write_text(json.dumps(result, default=str), encoding="utf-8")
        log.info("Worker %d wrote result (%s)", worker_id, result.get("status"))
        return 0

    except Exception as exc:
        # Surface the failure to the parent via both the result file (so
        # logs survive) and a non-zero exit code (so Popen.poll() sees it).
        err_payload = {
            "worker_id": worker_id,
            "status": "error",
            "error": f"{type(exc).__name__}: {exc}",
            "traceback": traceback.format_exc(),
            "project_results": [],
        }
        try:
            result_path.write_text(json.dumps(err_payload), encoding="utf-8")
        except Exception:
            pass
        log.exception("Worker %d crashed", worker_id)
        return 1


if __name__ == "__main__":
    sys.exit(main())
