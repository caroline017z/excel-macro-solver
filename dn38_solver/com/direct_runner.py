"""dn38_solver.com.direct_runner — Direct COM execution (no subprocess).

Runs the VBA macro in-process via COM. Since SolveHeadless handles all
GoalSeek logic in VBA, there's no need for subprocess isolation.

Performance gains over subprocess approach:
  - No process spawn / JSON marshal overhead
  - Single COM connection reused for macro + result reads
  - Warm-up calculate primes the formula dependency graph
"""
from __future__ import annotations

import contextlib
import logging
import os as _os
import shutil
import tempfile
import time
from pathlib import Path
from typing import Callable, NamedTuple

from dn38_solver.com.auto_recovery import (
    AutoRecoveryUnavailable,
    with_recovery,
)
from dn38_solver.com.hresult import decode_com_error, format_decoded
from dn38_solver.com.status_aggregator import atomic_write_json
from dn38_solver.com.vba_contract import (
    FINALIZE_SOLVE_ENV,
    INIT_SOLVE_ENV,
    SET_SKIP_OUTPUT_RECALC,
    SOLVE_ONE_PROJECT_BY_COL,
    STAMP_ACTIVE_PROJECT_COLUMN,
    SWITCH_PROJECT_AND_RECALC,
    vba_call_str,
)
from dn38_solver.config import BASE_COL
from dn38_solver.convert import safe_float, safe_str_or_float, safe_value
from dn38_solver.shadow.validation import scan_workbook_errors
from dn38_solver.types import CellAddress, SolveTask

log = logging.getLogger(__name__)

STATUS_FILE = Path(__file__).resolve().parent.parent.parent / "solver_status.json"
SOLVER_RESULTS_SHEET = "__SolverResults"

# Upper bound on rows to read from __SolverResults in one bulk Range.Value call.
# 60 active projects + slack for header growth; one COM round-trip beats N+1.
_RESULTS_BULK_ROWS = 200

# Wall-clock cap on a single COM Application.Run call. Typical solve is
# 2-3 min per project; cold-start solves can reach 5-6 min. 600s gives
# 2x headroom for the slowest legitimate solve while still surfacing
# stuck-recalc hangs. Configurable via env override.
DEFAULT_PER_CALL_TIMEOUT_SEC = int(_os.environ.get("DN38_PER_CALL_TIMEOUT_SEC", "600"))


def _capture_excel_proc(excel=None):
    """Find the EXCEL.EXE process backing this COM session.

    Two-tier lookup:

    1. **Hwnd path (preferred)** — query ``excel.Application.Hwnd`` and
       resolve the owning PID via ``GetWindowThreadProcessId``. COM
       activation typically spawns ``EXCEL.EXE /automation -Embedding``
       detached from the caller (parent is svchost / DCOMLAUNCH), so the
       child-process scan below returns ``None`` on most configurations.
       The window handle is the only reliable bridge from the Dispatch
       handle to a PID we can kill.

    2. **Child-process scan (fallback)** — preserved for builds where
       Hwnd is 0 (Excel still warming up) or psutil can't see the PID.

    Returns a ``psutil.Process`` or ``None``. The watchdog falls back to
    a no-op kill if this returns ``None`` — the timeout still raises,
    just without forcing Excel to disconnect.
    """
    try:
        import psutil  # type: ignore
    except ImportError:
        return None

    # Tier 1: Hwnd → PID via Win32. Allow up to 3s for the Excel window
    # handle to appear after Dispatch (some configurations create it
    # lazily on the first property access).
    if excel is not None:
        try:
            import ctypes
            from ctypes import wintypes
            user32 = ctypes.windll.user32
            for _ in range(30):
                try:
                    hwnd = int(excel.Application.Hwnd)
                except Exception:
                    hwnd = 0
                if hwnd:
                    pid = wintypes.DWORD(0)
                    user32.GetWindowThreadProcessId(
                        wintypes.HWND(hwnd), ctypes.byref(pid)
                    )
                    if pid.value > 0:
                        try:
                            proc = psutil.Process(pid.value)
                            if proc.name().lower() == "excel.exe":
                                return proc
                        except (psutil.NoSuchProcess, psutil.AccessDenied):
                            pass
                    break  # got Hwnd but PID lookup failed — don't keep polling
                time.sleep(0.1)
        except Exception:
            pass  # fall through to legacy child-scan

    # Tier 2: legacy child-process scan. DispatchEx returns before
    # Excel's child PID is necessarily registered as a descendant of
    # the caller; allow up to 3s for the relationship to materialize.
    # Worker startup is already 5-10s so this is in-noise.
    try:
        me = psutil.Process(_os.getpid())
    except Exception:
        return None
    for _ in range(30):
        try:
            for child in me.children(recursive=False):
                try:
                    if child.name().lower() == "excel.exe":
                        return child
                except (psutil.NoSuchProcess, psutil.AccessDenied):
                    continue
        except Exception:
            pass
        time.sleep(0.1)
    return None


def _kill_excel_after_timeout(done, *, timeout_sec, excel_proc, label):
    """Watchdog body shared by `_call_with_timeout` and `_run_macro_with_timeout`.

    Waits up to `timeout_sec` for `done` to signal. On timeout, escalates:
    psutil.kill → 3s poll → taskkill /F via subprocess. The escalation is
    necessary because psutil.kill silently fails on Excel instances deep
    in a COM RPC call — taskkill's Win32 TerminateProcess bypasses the
    COM busy state. Without an `excel_proc` handle (Hwnd lookup failed),
    the timeout still fires but the kill is a no-op and the COM call may
    block until the worker is restarted.
    """
    if done.wait(timeout=timeout_sec):
        return
    log.error(
        "  %s exceeded %ds wall-clock -- killing Excel to break the hang",
        label, timeout_sec,
    )
    if excel_proc is None:
        log.error(
            "  %s: no Excel proc handle captured -- watchdog cannot force "
            "COM to disconnect. Restart the worker to recover.", label,
        )
        return
    try:
        excel_proc.kill()
    except Exception as kill_exc:
        log.warning("  Excel psutil.kill failed: %s", kill_exc)
    for _ in range(30):
        try:
            if not excel_proc.is_running():
                return
        except Exception:
            break  # poll can fail mid-RPC; fall through to taskkill anyway
        time.sleep(0.1)
    log.warning(
        "  Excel pid=%s still alive (or unreadable) after psutil.kill "
        "-- escalating to taskkill /F", getattr(excel_proc, "pid", "?"),
    )
    try:
        import subprocess
        subprocess.run(
            ["taskkill", "/F", "/PID", str(excel_proc.pid)],
            capture_output=True, timeout=5, check=False,
        )
    except Exception as tk_exc:
        log.error("  taskkill escalation failed: %s", tk_exc)


def _with_watchdog(fn, *, timeout_sec, excel_proc, label):
    """Run `fn()` on the calling thread with a daemon watchdog that
    hard-kills Excel if the call exceeds `timeout_sec`. Returns `fn()`'s
    return value. The COM apartment stays on the calling thread.

    Used by both the macro-call wrapper (`_run_macro_with_timeout`) and
    the generic COM-call wrapper (`_call_with_timeout`). Without this
    guard, a VBA macro that enters an infinite recalc loop blocks
    Python indefinitely with no way to surface the failure.
    """
    import threading
    done = threading.Event()
    t = threading.Thread(
        target=_kill_excel_after_timeout,
        kwargs=dict(done=done, timeout_sec=timeout_sec,
                    excel_proc=excel_proc, label=label),
        daemon=True,
        name=f"COM-WD-{label}",
    )
    t.start()
    try:
        return fn()
    finally:
        done.set()


def _call_with_timeout(fn, *, timeout_sec, excel_proc=None, label="COM-call"):
    """Wrap a Python callable (e.g., `excel.CalculateFull`, `wb.Save`)
    in the COM watchdog. See `_with_watchdog` for kill semantics.
    """
    return _with_watchdog(fn, timeout_sec=timeout_sec,
                          excel_proc=excel_proc, label=label)


def _run_macro_with_timeout(excel, call_args, *, timeout_sec,
                             excel_proc=None, label="Application.Run"):
    """Wrap `excel.Application.Run(*call_args)` in the COM watchdog.
    Raises TimeoutError on watchdog fire, or the underlying COM exception
    if Application.Run raises before the timeout.
    """
    return _with_watchdog(
        lambda: excel.Application.Run(*call_args),
        timeout_sec=timeout_sec, excel_proc=excel_proc, label=label,
    )


class _NormCell(NamedTuple):
    sheet: str
    address: str  # may contain "{col}" template placeholder


class _NormTask(NamedTuple):
    col: str
    offset: int
    name: str
    idx_cell: CellAddress | dict
    read_cells: tuple[_NormCell, ...]


def _norm_cell(c: object) -> _NormCell:
    if hasattr(c, "address"):
        return _NormCell(sheet=c.sheet, address=c.address)
    return _NormCell(sheet=c["sheet"], address=c["address"])


def _norm_task(t: object) -> _NormTask:
    if hasattr(t, "project_col_letter"):
        cells = tuple(_norm_cell(c) for c in t.read_cells)
        return _NormTask(
            col=t.project_col_letter,
            offset=int(t.project_offset),
            name=t.project_name,
            idx_cell=t.project_index_cell,
            read_cells=cells,
        )
    cells = tuple(_norm_cell(c) for c in t.get("read_cells", []))
    return _NormTask(
        col=t["project_col_letter"],
        offset=int(t["project_offset"]),
        name=t["project_name"],
        idx_cell=t["project_index_cell"],
        read_cells=cells,
    )


class _StatusWriter:
    """Buffered solver-status writer.

    Holds the immutable run-level fields (workbook, project list, start time)
    so each phase update only constructs a small delta dict before writing.
    Output is still JSON to STATUS_FILE for the Streamlit tracker.
    """

    __slots__ = ("_base", "_start", "_path", "_worker_id")

    def __init__(
        self,
        *,
        workbook_path: str,
        proj_names: list[str],
        start: float,
        path: Path = STATUS_FILE,
        worker_id: int | None = None,
    ) -> None:
        self._base = {
            "workbook": workbook_path,
            "total_projects": len(proj_names),
            "_proj_names": proj_names,
        }
        self._start = start
        self._path = path
        self._worker_id = worker_id

    def update(
        self,
        phase: str,
        *,
        per_project_status: str | None = None,
        projects: list[dict] | None = None,
        **extras: object,
    ) -> None:
        if projects is None and per_project_status is not None:
            projects = [
                {"name": n, "status": per_project_status}
                for n in self._base["_proj_names"]
            ]
        payload = {
            "phase": phase,
            "workbook": self._base["workbook"],
            "total_projects": self._base["total_projects"],
            "projects": projects or [],
            "elapsed_sec": time.time() - self._start,
        }
        if self._worker_id is not None:
            payload["worker_id"] = self._worker_id
        payload.update(extras)
        # Atomic swap with bounded retry so a transient Windows
        # PermissionError (AV scan, dashboard reader, file-index service)
        # doesn't drop a status update silently. Helper retries on a
        # short backoff schedule and logs only on final failure — quiet
        # in the common case, loud when something is genuinely stuck.
        atomic_write_json(self._path, payload)


def run_direct(
    workbook_path: str,
    tasks: list[SolveTask],
    *,
    original_f2: int = 1,
    timeout_sec: int = 600,
    use_chunked: bool = False,
    checkpoint_callback: Callable[[object, dict], None] | None = None,
    save_solved: bool = True,
    skip_output_recalc: bool = False,
    strip_sheets: tuple[str, ...] = (),
    output_path: str | None = None,
    status_path: Path | None = None,
    excel_threads: int | None = None,
    worker_id: int | None = None,
) -> dict:
    """Open Excel, run SolveHeadless, read results, close. All in-process.

    When `use_chunked` is True, the macro runs through the
    InitSolveEnvHL / SolveOneProjectByColHL / FinalizeSolveEnvHL entry
    points instead of the single-shot SolveHeadless. Each project becomes
    its own COM Application.Run call so no single invocation can exceed
    Excel's ~900s RPC timeout — the win on cold portfolios that today
    crash before completion. Per-project progress is also surfaced live
    via the status JSON between calls.

    `checkpoint_callback`, if provided and `use_chunked` is True, fires
    after each successful per-project SolveOneProjectByColHL with
    `(_NormTask, dict-of-__SolverResults-row)`. Use it to persist
    partial results so a mid-portfolio crash leaves an audit trail.
    Exceptions raised inside the callback are caught and logged so a
    persistence failure cannot stall the solve.

    When `save_solved` is False, the `<workbook>_SOLVED.xlsm` SaveAs at
    end of run is skipped. Useful for fast iteration on Box-mounted
    workbooks where the save is a nontrivial fixed cost.

    When `skip_output_recalc` is True, VBA skips CalcOutputSheetsHL at
    finalize — the 7 downstream output sheets (Portfolio, AT Returns_WIP,
    Corp Model Output, Cust Prop, Dashboard, Table, Waterfall Sensitivity)
    are left un-recalculated. Excel recalcs them lazily on next interactive
    open. Saves 10-30s on workbooks with #REF!-heavy Dashboard / Waterfall.

    `strip_sheets` is a tuple of sheet names to DELETE from the temp copy
    via openpyxl before Excel opens it. Use to drop non-essential output
    sheets (e.g. Dashboard with 50K #REF! cells) so SaveAs and the post-
    export validation scan don't pay to serialize them. Original workbook
    is untouched (direct_runner always operates on a tempfile copy).
    Deletion is more aggressive than skip_output_recalc — only safe when
    no core-sheet formula references the deleted sheets.

    `output_path` overrides the default `<workbook>_SOLVED.xlsm` location.
    Used by the parallel runner so each worker saves to its own temp dir
    instead of all workers racing on the same path next to the source.

    `status_path` overrides the default solver_status.json location. Used
    by the parallel runner so each worker writes its own status file
    (solver_status_w{id}.json) instead of multiple workers stomping on
    the canonical one.

    `excel_threads` caps Excel's MultiThreadedCalculation.ThreadCount.
    Default None uses Excel's default (= Environment.ProcessorCount). Set
    to cpu_count // workers in parallel mode so N workers don't oversubscribe
    the CPU (4 workers × 8 threads = 32 OS threads competing).
    """
    import pythoncom
    import win32com.client

    start = time.time()
    tmp_dir = Path(tempfile.mkdtemp(prefix="38dn_com_"))
    temp_path = tmp_dir / Path(workbook_path).name
    shutil.copy2(workbook_path, str(temp_path))

    # Optional: strip non-essential sheets from the temp copy before Excel
    # opens it. Saves SaveAs time, post-export validation time, and removes
    # the recalc cost for sheets the user doesn't review per solve. The
    # original workbook on Box/OneDrive is unaffected — we always operate
    # on a tempfile copy.
    if strip_sheets:
        # Hard guardrail: refuse to strip sheets the solver / deal team
        # depend on every run. Even if the user passes one of these names,
        # silently skip it and log a warning rather than corrupt the run.
        _CRITICAL_SHEETS = {
            "Dashboard", "Table", "PT Returns", "NPP Calc", "Appraisal",
            "Perm Debt", "Tax Equity", "CL", "Project Inputs",
        }
        requested = list(strip_sheets)
        rejected = [s for s in requested if s in _CRITICAL_SHEETS]
        safe_to_strip = [s for s in requested if s not in _CRITICAL_SHEETS]
        if rejected:
            log.warning(
                "Refusing to strip critical sheet(s): %s. These sheets "
                "must recalc every run. Stripping ignored for these names.",
                ", ".join(rejected),
            )
        if safe_to_strip:
            # KNOWN-RISKY PATTERN: openpyxl save on an .xlsm pre-solve
            # can corrupt workbook state (data validation extLst,
            # calcChain, cond formatting) in ways the macro tolerates
            # for direct writes but rejects during the first GoalSeek.
            #
            # Safety net: the strip runs ONCE, before any solve. If the
            # first macro call hits an auto-recoverable HRESULT,
            # auto_recovery.with_recovery closes the workbook and
            # invokes import_vba_module.py, whose Excel COM SaveAs
            # rewrites the file. That rewrite clears the openpyxl
            # artifacts. The strip is NOT re-run on retry, so the
            # recovery loop cannot re-introduce the corruption.
            # Stripped sheets remain stripped (intended); the artifacts
            # that broke the first solve are gone.
            #
            # DO NOT add new openpyxl.save() sites on .xlsm files. Use
            # dn38_solver.com.com_edit.edit_xlsm — it routes the
            # mutation through Excel COM SaveAs (FileFormat=52), which
            # has been verified not to corrupt state.
            try:
                import openpyxl
                wb_strip = openpyxl.load_workbook(str(temp_path), keep_vba=True)
                removed = []
                for sheet_name in safe_to_strip:
                    if sheet_name in wb_strip.sheetnames:
                        del wb_strip[sheet_name]
                        removed.append(sheet_name)
                if removed:
                    wb_strip.save(str(temp_path))
                    log.info("  Stripped %d sheet(s) from temp copy: %s",
                             len(removed), ", ".join(removed))
                wb_strip.close()
            except Exception as strip_exc:
                log.warning(
                    "Sheet strip failed (%s) — proceeding with full workbook. "
                    "Common cause: a core-sheet formula references the deleted "
                    "sheet, which would create #REF! errors on save.", strip_exc,
                )

    excel = None
    wb = None

    result: dict = {
        "status": "error",
        "project_results": [],
        "duration_sec": 0.0,
        "saved_to": None,
        "error": None,
        "macro_used": None,
    }

    try:
        pythoncom.CoInitialize()
        # DispatchEx forces a fresh COM instance so we never attach to a
        # zombie EXCEL.EXE left over from a prior crash, and never share
        # process state with an Excel session the user has open for their
        # own work.
        excel = win32com.client.DispatchEx("Excel.Application")
        # msoAutomationSecurityLow = 1. With the default ByUI security
        # (=2), Excel treats headless automation as no-consent and blocks
        # macros configured as "Disable with notification" in Trust
        # Center — surfacing as "Cannot run the macro ... macros may be
        # disabled." Setting Low scopes the override to this COM session
        # only, so the user's interactive Excel settings are unchanged.
        # Surface any failure rather than swallow it — a silent failure
        # here looks identical to "macro not in workbook" downstream.
        try:
            excel.AutomationSecurity = 1
            log.debug("  AutomationSecurity set to Low (=1)")
        except Exception as as_exc:
            log.warning(
                "Could not set AutomationSecurity=1 (%s) — Excel will use "
                "Trust Center default and may block macro execution.",
                as_exc,
            )
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        # Cap multi-threaded calc threads if requested. In parallel mode
        # we want each Excel instance to use fewer threads so N workers
        # don't oversubscribe the CPU (4 workers × default 8 threads = 32
        # OS threads thrashing). Single-process runs leave it at Excel's
        # default (= cpu_count) for max single-instance throughput.
        if excel_threads is not None and excel_threads > 0:
            try:
                excel.MultiThreadedCalculation.Enabled = True
                excel.MultiThreadedCalculation.ThreadCount = excel_threads
                log.debug(
                    "  MultiThreadedCalculation.ThreadCount capped at %d",
                    excel_threads,
                )
            except Exception as mtc_exc:
                log.warning(
                    "Could not cap MultiThreadedCalculation threads (%s)",
                    mtc_exc,
                )

        # Normalize task payloads once so the per-project loop is branch-free.
        norm_tasks = [_norm_task(t) for t in tasks]
        proj_names = [nt.name for nt in norm_tasks]

        status = _StatusWriter(
            workbook_path=workbook_path,
            proj_names=proj_names,
            start=start,
            path=status_path if status_path is not None else STATUS_FILE,
            worker_id=worker_id,
        )
        status.update("opening", per_project_status="pending")

        log.debug("  Opening workbook via COM...")
        wb = excel.Workbooks.Open(
            str(temp_path),
            ReadOnly=False,
            UpdateLinks=0,
        )
        open_time = time.time() - start
        log.debug("  Opened in %.1fs", open_time)

        warmup_time = 0.0

        # Capture the Excel process handle as soon as the workbook is
        # open so setup-call timeouts (SET_SKIP_OUTPUT_RECALC) can use
        # the watchdog kill path too. None if psutil unavailable or both
        # Hwnd-lookup and child-scan failed.
        excel_proc = _capture_excel_proc(excel)
        if excel_proc is None:
            log.warning(
                "  _capture_excel_proc returned None — per-call watchdog "
                "will time out but cannot kill Excel. Hangs will leak."
            )
        else:
            log.debug("  Captured Excel pid=%s for watchdog", excel_proc.pid)

        # Propagate the skip-output-recalc flag to VBA before either macro
        # path runs. Both single-shot SolveHeadless and chunked Finalize
        # funnel through CalcOutputSheetsHL, which checks mSkipOutputRecalc
        # at its top. Older workbooks without the setter silently ignore
        # the call (the macro defaults to recalc, preserving prior behavior).
        if skip_output_recalc:
            try:
                _run_macro_with_timeout(
                    excel,
                    (vba_call_str(wb.Name, SET_SKIP_OUTPUT_RECALC), True),
                    timeout_sec=120,  # setter is trivial; 2 min is generous
                    excel_proc=excel_proc,
                    label="SetSkipOutputRecalc",
                )
                log.info("  Output-sheet recalc disabled for this run")
            except Exception as set_exc:
                log.warning(
                    "Could not set SkipOutputRecalc (%s) — older VBA module; "
                    "output sheets will be recalculated as usual.",
                    set_exc,
                )

        # --- Run the VBA macro ---
        macro_used: str | None = None
        macro_error: str | None = None
        # C1: True once auto-recovery resumed a crashed chunked run. The
        # orchestrator keeps per-project checkpoints (rather than clearing
        # them on a clean run) when this is set, so the correct in-flight
        # telemetry for pre-failure projects survives for forensics even if
        # the pre-recovery COM save did not.
        run_recovered = False

        log.info(
            "  Running macro (%s)...",
            "chunked: per-project" if use_chunked else "single-shot SolveHeadless",
        )
        status.update(
            "solving",
            per_project_status="solving",
            macro_used="SolveHeadless",
            chunked=use_chunked,
        )

        t0 = time.time()
        if use_chunked:
            macro_used, has_switch, macro_error, failed_idx = _run_chunked(
                excel, wb, norm_tasks, original_f2, status,
                checkpoint_callback=checkpoint_callback,
                excel_proc=excel_proc,
            )
        else:
            macro_used, has_switch, macro_error = _run_single_shot(
                excel, wb, excel_proc=excel_proc,
            )
            failed_idx = None

        # Auto-recovery: a generic VBA error inside Application.Run
        # (e.g. 0x80048028) almost always means the workbook's state is
        # stale relative to the embedded macro — typically because an
        # openpyxl save (or another non-Excel writer) touched the .xlsm
        # since the last macro import. A fresh `import_vba_module.py`
        # pass via Excel COM SaveAs rewrites the file and clears the
        # corruption. Wrap that recovery as a one-shot retry here.
        #
        # PRE-RECOVERY SNAPSHOT: __SolverResults rows for projects that
        # converged BEFORE the failing idx are still in memory on the
        # open workbook, but `wb.Close(SaveChanges=False)` below would
        # discard them. Snapshot now so the post-recovery read-pass can
        # merge them in — otherwise pre-failure converged projects show
        # as `not_attempted` in the run record.
        pre_recovery_results: dict[int, dict] = {}
        pre_recovery_heartbeat: str | None = None
        if (
            use_chunked
            and macro_error
            and failed_idx is not None
            and failed_idx > 0
            and decode_com_error(macro_error).auto_recoverable
        ):
            try:
                pre_recovery_results, pre_recovery_heartbeat = (
                    _read_solver_results_map(wb)
                )
                if pre_recovery_results:
                    log.info(
                        "  Pre-recovery snapshot: captured %d converged "
                        "project(s) before retry close.",
                        len(pre_recovery_results),
                    )
            except Exception as snap_exc:
                # Non-fatal: recovery proceeds, but partial results are lost.
                # Logged so an operator forensicing a partial run sees why
                # the pre-error projects show not_attempted.
                log.warning(
                    "  Pre-recovery snapshot failed (%s); partial results "
                    "from converged-before-error projects will not be "
                    "preserved.", snap_exc,
                )

        if (
            use_chunked
            and macro_error
            and decode_com_error(macro_error).auto_recoverable
        ):
            decoded = decode_com_error(macro_error)
            log.warning(
                "  Macro failed with an auto-recoverable error. Attempting "
                "recovery (re-import + retry once).\n%s",
                format_decoded(decoded),
            )
            status.update(
                "recovering",
                per_project_status="recovering",
                macro_used=macro_used,
                recovery_reason=decoded.summary,
            )

            retry_used: list = [macro_used]
            retry_has_switch: list = [has_switch]
            retry_error: list = [macro_error]

            def _close_wb() -> None:
                nonlocal wb
                if wb is not None:
                    # C1: Save the scratch temp copy via Excel COM BEFORE
                    # closing, so the in-memory converged stamps for every
                    # project solved BEFORE the failure (Project Inputs
                    # columns 32/33/38/39/371 and their __SolverResults rows)
                    # survive into the reopened file. Without this, the
                    # SaveChanges=False close discarded them; the retry
                    # resumed from the failed project and never re-stamped the
                    # earlier ones; and the read pass then stamped PRE-SOLVE
                    # values into those columns while labeling them converged
                    # — silent wrong NPP on a green run. COM Save (not
                    # openpyxl) is the sanctioned write path and the temp file
                    # is scratch, so the save is safe. Suppressed on failure:
                    # we fall back to the pre-recovery __SolverResults snapshot
                    # and the retained in-flight checkpoints.
                    try:
                        _call_with_timeout(
                            wb.Save,
                            timeout_sec=300,
                            excel_proc=excel_proc,
                            label="Save[pre-recovery-preserve]",
                        )
                        log.info(
                            "  Saved pre-failure converged state to temp copy "
                            "before recovery close."
                        )
                    except Exception as save_exc:
                        log.warning(
                            "  Pre-recovery save failed (%s) — pre-failure "
                            "projects recovered from the __SolverResults "
                            "snapshot + checkpoints, not re-stamped cells.",
                            save_exc,
                        )
                    with contextlib.suppress(Exception):
                        wb.Close(SaveChanges=False)
                    wb = None

            def _reopen_wb() -> None:
                nonlocal wb
                wb = excel.Workbooks.Open(
                    str(temp_path),
                    ReadOnly=False,
                    UpdateLinks=0,
                )
                if skip_output_recalc:
                    with contextlib.suppress(Exception):
                        _run_macro_with_timeout(
                            excel,
                            (vba_call_str(wb.Name, SET_SKIP_OUTPUT_RECALC), True),
                            timeout_sec=120,
                            excel_proc=excel_proc,
                            label="SetSkipOutputRecalc[recovery]",
                        )

            def _retry() -> None:
                # Resume from the failed project, not project 1 — the
                # pre-recovery snapshot already captured the converged
                # results from earlier projects.
                resume_from = failed_idx if failed_idx is not None else 0
                used, switch, err, retry_failed_idx = _run_chunked(
                    excel, wb, norm_tasks, original_f2, status,
                    checkpoint_callback=checkpoint_callback,
                    start_idx=resume_from,
                    excel_proc=excel_proc,
                )
                retry_used[0] = used
                retry_has_switch[0] = switch
                retry_error[0] = err
                if err:
                    # If the retry failed at the SAME project, the bug is
                    # deterministic — re-import doesn't help, and a
                    # second retry would just re-burn the wall clock.
                    # Data-driven failures (e.g., placeholder columns
                    # producing #DIV/0!) belong with the preflight
                    # warnings, not the recovery loop.
                    if (
                        retry_failed_idx is not None
                        and failed_idx is not None
                        and retry_failed_idx == failed_idx
                    ):
                        raise RuntimeError(
                            f"{err} (recovery failed at same project idx={failed_idx} "
                            "-- bug is data-driven, not workbook-state corruption; "
                            "no further retries)"
                        )
                    raise RuntimeError(err)

            try:
                recovered = with_recovery(
                    workbook_path=temp_path,
                    close_open_handle=_close_wb,
                    reopen_handle=_reopen_wb,
                    retry_callable=_retry,
                    already_recovered=False,
                )
            except AutoRecoveryUnavailable as ar_exc:
                log.warning("  Auto-recovery unavailable: %s", ar_exc)
                recovered = False

            if recovered:
                macro_used = retry_used[0]
                has_switch = retry_has_switch[0]
                macro_error = None
                run_recovered = True
            else:
                # Surface the post-recovery error so the operator knows
                # this wasn't a first-pass failure. Keep the original
                # decoded hint visible upstream.
                macro_error = (
                    f"auto-recovery did not converge — original: {macro_error} "
                    f"| retry: {retry_error[0]}"
                )

        solve_time = time.time() - t0
        log.info("  Macro '%s' completed in %.1fs", macro_used, solve_time)

        if macro_used is None:
            # No macro to run is unrecoverable — there's nothing in
            # __SolverResults to read and no output worth saving.
            result["error"] = "No solver macro found in workbook"
            return result

        # macro_error and timeout are non-fatal: a chunked Finalize failure
        # or a mid-portfolio crash still leaves valid rows in __SolverResults
        # and converged values on Project Inputs / PT Returns. We always
        # proceed to read whatever landed and save the xlsx so the partial
        # state is recoverable. The error is propagated on the result so
        # the orchestrator can mark the run accordingly and keep
        # checkpoints rather than silently clearing them.
        timeout_error: str | None = None
        if timeout_sec > 0 and solve_time > timeout_sec:
            timeout_error = (
                f"Macro execution exceeded timeout_sec={timeout_sec} "
                f"(actual={solve_time:.1f}s)"
            )

        # --- Read results per project ---
        # SolveHeadless leaves calc in MANUAL mode.
        # has_switch indicates whether SwitchProjectAndRecalc lives
        # alongside the macro that ran; the runner helpers report it
        # rather than inferring from a name string.
        log.debug("  Reading results for %d project(s)...", len(tasks))
        status.update(
            "reading",
            per_project_status="reading",
            macro_used=macro_used,
            macro_time_sec=solve_time,
        )
        t0 = time.time()
        project_results = []
        solver_results, heartbeat = _read_solver_results_map(wb)

        # Merge in any pre-recovery snapshot (captured before the auto-
        # recovery close discarded in-memory __SolverResults rows). The
        # post-recovery read takes priority for any offset it covers —
        # rows written after recovery reflect actual retry outcomes; the
        # snapshot only fills in offsets the recovered workbook lost.
        # Empty dict when no recovery occurred, so the no-error path is
        # unchanged.
        if pre_recovery_results:
            for offset, meta in pre_recovery_results.items():
                solver_results.setdefault(offset, meta)
            heartbeat = heartbeat or pre_recovery_heartbeat

        for nt in norm_tasks:
            meta = solver_results.get(nt.offset)
            if meta is None:
                # Mid-portfolio failure stopped before this project ran.
                # Don't read its cells — they'd reflect either uninitialized
                # state or another project's converged values, both of
                # which would be misleading downstream.
                project_results.append({
                    "project_name": nt.name,
                    "project_offset": nt.offset,
                    "status": "not_attempted",
                    "solved_values": {},
                    "iterations_used": 0,
                    "duration_sec": 0,
                    "meta": {"project_offset": nt.offset},
                    "_summary": {"npp": None, "dev_fee": None, "fmv": None},
                })
                continue

            stamp_failed = False

            # Switch F2 with targeted recalc. The fallback path sets F2
            # directly via Range.Value, which does NOT trigger recalc under
            # xlCalculationManual — reads after that would see the previous
            # project's values. Force a full recalc when we fall back so
            # the legacy-macro path doesn't silently mix projects.
            #
            # Both the macro and fallback paths get the watchdog timeout —
            # SwitchProjectAndRecalc calls CalcModelCoreHL which can spin
            # on a corrupted workbook (same hang mode as the per-project
            # solve). Per-call timeout caps the read-pass at N * timeout
            # in worst case, where N is project count.
            switched_with_recalc = False
            if has_switch:
                try:
                    _run_macro_with_timeout(
                        excel,
                        (vba_call_str(wb.Name, SWITCH_PROJECT_AND_RECALC), nt.offset),
                        timeout_sec=DEFAULT_PER_CALL_TIMEOUT_SEC,
                        excel_proc=excel_proc,
                        label=f"SwitchProjectAndRecalc[{nt.name}]",
                    )
                    switched_with_recalc = True
                except Exception as switch_exc:
                    log.warning(
                        "  Read-pass switch failed for %s: %s — falling back to direct F2",
                        nt.name, switch_exc,
                    )
                    _set_f2(wb, nt.idx_cell, nt.offset)
            else:
                _set_f2(wb, nt.idx_cell, nt.offset)

            if not switched_with_recalc:
                # Bare CalculateFull() could hang 5-15 min on a workbook
                # whose iterative calc was left in a non-converging state
                # by the failed switch. Wrap in a short watchdog so the
                # read-pass for a single project can't burn the wall
                # clock — 60s is generous for a recalc that normally
                # completes in 1-3s on this workbook size.
                with contextlib.suppress(Exception):
                    _call_with_timeout(
                        excel.CalculateFull,
                        timeout_sec=60,
                        excel_proc=excel_proc,
                        label=f"CalculateFull[fallback-recalc:{nt.name}]",
                    )

            # Stamp the active project's per-column convergence cells
            # NOW, while F2 is pinned to this project and the workbook
            # is in its post-all-solves consistent state. Replaces the
            # in-solve hard-stamps that captured a transient cross-
            # project state and produced merged-file values diverging
            # from what we're about to read in the next loop. Skip on
            # the legacy fallback path that doesn't have the helper.
            #
            # Skip the per-column stamp when an upstream macro_error
            # already broke the run. Stamping invokes CalculateFull,
            # which on a corrupted workbook spins the circular sticky-IF
            # cells in rows 31/37 against #DIV/0 propagation for the
            # full MaxIterations budget per project. The values we'd
            # read are unreliable anyway when the solve already failed.
            #
            # When stamping IS run, never suppress its exception: a
            # silent stamp failure leaves the per-column cell as a
            # circular IF formula that openpyxl-merge reads as None.
            #
            # Pass this project's DSCR so the stamp can restore PT!F129
            # before the CalculateFull that pins rows 31/37. PT!F129 is
            # a single live cell that GoalSeek overwrites once per
            # project; without the restore, every project except the
            # worker's last gets its IRR computed against the wrong
            # DSCR. Use 0.0 sentinel when no DSCR was captured (the VBA
            # side reads >0 as "perform restore", 0 as "skip").
            if has_switch and not macro_error:
                col_idx = nt.offset + BASE_COL
                dscr_for_stamp = safe_float(meta.get("dscr")) or 0.0
                try:
                    _run_macro_with_timeout(
                        excel,
                        (
                            vba_call_str(wb.Name, STAMP_ACTIVE_PROJECT_COLUMN),
                            int(col_idx),
                            float(dscr_for_stamp),
                        ),
                        timeout_sec=DEFAULT_PER_CALL_TIMEOUT_SEC,
                        excel_proc=excel_proc,
                        label=f"StampActiveProjectColumnHL[{nt.name}]",
                    )
                except Exception as stamp_exc:
                    # C2: localize the failure. A stamp error leaves THIS
                    # project's per-column cells as unresolved circular-IF
                    # formulas (openpyxl reads them as None). Previously this
                    # propagated to the function-level catch-all, which
                    # returned empty project_results and skipped the SaveAs —
                    # discarding EVERY already-converged project in the open
                    # workbook (the SolarStone 2026-06-04 incident, and a
                    # direct contradiction of this module's "always salvage
                    # partial results" contract). Flag this one project and
                    # keep going so the rest of the portfolio and the SaveAs
                    # still land.
                    log.error(
                        "  StampActiveProjectColumnHL FAILED for %s: %s — "
                        "marking project stamp_failed; remaining projects and "
                        "the SaveAs proceed.", nt.name, stamp_exc,
                    )
                    stamp_failed = True

            # Read cells
            solved: dict[str, float | str | None] = {}
            npp = dev_fee = fmv = None
            for cell in nt.read_cells:
                addr = cell.address.replace("{col}", nt.col)
                key = f"{cell.sheet}!{addr}"
                val = _read_cell(wb, cell.sheet, addr)
                solved[key] = val
                # Capture summary scalars at read time so we don't re-scan
                # solved_values later for the tracker payload.
                if addr.endswith("38"):
                    npp = safe_float(val)
                elif addr.endswith("32"):
                    dev_fee = safe_float(val)
                elif addr.endswith("33"):
                    fmv = safe_float(val)

            # Prefer per-project DSCR captured during solve loop to avoid
            # last-project F129 bleed in multi-project runs.
            dscr_key = "PT Returns!F129"
            if "dscr" in meta:
                solved[dscr_key] = meta["dscr"]

            # Trust VBA's converged_flag in column I rather than assuming
            # every row read means convergence. A row exists for every
            # project the macro attempted, including ones that timed out
            # of their inner loop without hitting tolerance.
            converged_flag = meta.get("converged_flag")
            is_converged = bool(converged_flag) if converged_flag is not None else False

            # Stamp project_offset into meta so downstream consumers
            # (parallel_runner.run_parallel re-keying across workers) can
            # align results without relying on project names, which may
            # not be unique across an intake portfolio.
            meta = dict(meta) if meta else {}
            meta["project_offset"] = nt.offset

            # Distinguish skipped placeholders from real convergence
            # failures. VBA's early-skip writes a sentinel like
            # "skipped:no_rc1_revenue" to __SolverResults col 12
            # (surfaced as meta["mode"]). Surfacing "skipped" as a
            # distinct status lets the CLI summary, dashboard tracker,
            # and post-merge verifier each render it correctly without
            # re-parsing the mode string.
            mode = meta.get("mode")
            if stamp_failed:
                # C2: the per-column stamp didn't land, so this project's
                # read-back values are untrusted. Surface a distinct status
                # and force the tier to "none" — convergence_label keys off
                # the tier, not converged_flag, so a leftover "strict" tier
                # would otherwise render this failed project as OK.
                project_status = "stamp_failed"
                meta["conv_tier"] = "none"
            elif isinstance(mode, str) and mode.startswith("skipped:"):
                project_status = "skipped"
            elif is_converged:
                project_status = "converged"
            else:
                project_status = "not_converged"

            project_results.append({
                "project_name": nt.name,
                "project_offset": nt.offset,
                "status": project_status,
                "solved_values": solved,
                "iterations_used": int(meta.get("iterations") or 0),
                "duration_sec": 0,
                "meta": meta,
                "_summary": {"npp": npp, "dev_fee": dev_fee, "fmv": fmv},
            })

        read_time = time.time() - t0
        log.debug("  Read %d project(s) in %.1fs", len(tasks), read_time)

        # Restore original F2. Watchdog-wrapped so a silent hang here
        # doesn't delay workbook close and the SaveAs that produces
        # _SOLVED.xlsm. 60s cap is generous for the F2 reset + recalc.
        if norm_tasks and has_switch:
            with contextlib.suppress(Exception):
                _run_macro_with_timeout(
                    excel,
                    (
                        vba_call_str(wb.Name, SWITCH_PROJECT_AND_RECALC),
                        int(original_f2),
                    ),
                    timeout_sec=60,
                    excel_proc=excel_proc,
                    label="SwitchProjectAndRecalc[restore-F2]",
                )

        # Save solved workbook (opt-out via save_solved=False for fast
        # iteration runs; the post-export validation gate only runs when
        # there's actually a saved file to scan).
        # Box/OneDrive sync conflicts can silently fail SaveAs — log so a
        # multi-hour solve doesn't vanish without a trail.
        saved_to = None
        if save_solved:
            if output_path is not None:
                solved_path = Path(output_path)
            else:
                wb_path = Path(workbook_path)
                solved_name = wb_path.stem + "_SOLVED" + wb_path.suffix
                solved_path = wb_path.parent / solved_name
            try:
                wb.SaveAs(str(solved_path))
                saved_to = str(solved_path)
            except Exception as save_exc:
                log.warning(
                    "SaveAs failed for %s: %s — converged values remain in "
                    "__SolverResults and the open workbook but no _SOLVED file "
                    "was written.", solved_path, save_exc,
                )

        # Post-export formula-error gate: scan the just-saved file for
        # cached Excel error tokens (#REF! / #DIV/0! / #VALUE! / etc.).
        # Pure-Python via openpyxl — no LibreOffice dependency.
        validation = None
        if saved_to is not None:
            try:
                validation = scan_workbook_errors(saved_to)
            except Exception as val_exc:
                log.warning(
                    "Post-export validation scan failed for %s: %s — solve "
                    "results are still valid but #REF!/#DIV/0! cells were not "
                    "checked.", saved_to, val_exc,
                )

        total = time.time() - start

        # Build tracker payload from pre-computed per-project summaries.
        tracker_projects = [
            {
                "name": pr["project_name"],
                "status": pr["status"],
                **pr["_summary"],
            }
            for pr in project_results
        ]
        # _summary was a transport-only field for the tracker — drop it from
        # the returned project_results so downstream consumers see a clean shape.
        for pr in project_results:
            pr.pop("_summary", None)

        # Compose batch-level status from any error surfaced during the
        # macro / timeout path. project_results is populated either way so
        # the orchestrator can decide what to persist; the error string
        # tells it whether to keep checkpoints around for forensics.
        if macro_error:
            batch_status = "error"
            batch_error = f"Macro {macro_used} failed: {macro_error}"
        elif timeout_error:
            batch_status = "error"
            batch_error = timeout_error
        else:
            batch_status = "converged"
            batch_error = None

        result = {
            "status": batch_status,
            "project_results": project_results,
            "duration_sec": round(total, 2),
            "saved_to": saved_to,
            "error": batch_error,
            "macro_used": macro_used,
            "open_time_sec": round(open_time, 2),
            "warmup_time_sec": round(warmup_time, 2),
            "solve_time_sec": round(solve_time, 2),
            "read_time_sec": round(read_time, 2),
            "solver_heartbeat": heartbeat,
            "validation": validation,
            "recovered": run_recovered,
        }

        status.update(
            "complete" if batch_error is None else "error",
            projects=tracker_projects,
            total_time_sec=total,
            macro_used=macro_used,
            open_time_sec=round(open_time, 2),
            macro_time_sec=round(solve_time, 2),
            read_time_sec=round(read_time, 2),
            solver_heartbeat=heartbeat,
            error=batch_error,
        )

    except Exception as exc:
        result = {
            "status": "error",
            "project_results": [],
            "duration_sec": round(time.time() - start, 2),
            "saved_to": None,
            "error": f"{type(exc).__name__}: {exc}",
        }

    finally:
        # Close the workbook before flipping calc back to automatic. If the
        # close hangs (file lock, OneDrive sync), the prior order would
        # leave Excel evaluating volatile formulas while the file was
        # still open — making a slow hang slower. Each step is suppressed
        # so a single failure doesn't block the rest of the cleanup.
        if wb is not None:
            with contextlib.suppress(Exception):
                wb.Close(SaveChanges=False)
        if excel is not None:
            with contextlib.suppress(Exception):
                excel.Calculation = -4105  # xlCalculationAutomatic
            with contextlib.suppress(Exception):
                excel.ScreenUpdating = True
            with contextlib.suppress(Exception):
                excel.EnableEvents = True
            with contextlib.suppress(Exception):
                excel.Quit()
        # Log rather than swallow — a leaked tmp_dir typically means a
        # zombie EXCEL.EXE still holds a file handle, and that's worth
        # surfacing rather than silently filling up %TEMP%.
        try:
            shutil.rmtree(tmp_dir)
        except OSError as rm_exc:
            log.warning("Failed to remove tmp_dir %s: %s", tmp_dir, rm_exc)
        with contextlib.suppress(Exception):
            pythoncom.CoUninitialize()

    return result


def _run_single_shot(
    excel: object,
    wb: object,
    *,
    excel_proc=None,
    per_call_timeout_sec: int = DEFAULT_PER_CALL_TIMEOUT_SEC,
) -> tuple[str | None, bool, str | None]:
    """Legacy single-invocation macro path. Tries SolveHeadless first;
    falls back to the original SolveMinEquityWithHoldCo if the headless
    wrapper isn't present in the workbook (older models).

    Returns (macro_used, has_switch, error). has_switch is True when the
    SwitchProjectAndRecalc helper lives alongside the macro that ran --
    only the SolveHeadless module ships it, so the legacy fallback path
    reports False and the post-solve read uses the F2-direct fallback.

    The Application.Run call is wrapped in `_run_macro_with_timeout`
    so a hung VBA loop on this path can't burn the wall clock.
    """
    macro_names = ("SolveHeadless", "SolveMinEquityWithHoldCo")
    for macro_name in macro_names:
        try:
            _run_macro_with_timeout(
                excel,
                (f"'{wb.Name}'!{macro_name}",),
                timeout_sec=per_call_timeout_sec,
                excel_proc=excel_proc,
                label=f"{macro_name}[single-shot]",
            )
            return macro_name, macro_name == "SolveHeadless", None
        except TimeoutError as e:
            return macro_name, macro_name == "SolveHeadless", (
                f"{macro_name} TIMEOUT (single-shot): {e}"
            )
        except Exception as e:
            err_str = str(e).lower()
            if "macro may not be available" in err_str or "cannot run" in err_str:
                continue
            return macro_name, macro_name == "SolveHeadless", str(e)
    return None, False, None


def _run_chunked(
    excel: object,
    wb: object,
    norm_tasks: list[_NormTask],
    original_f2: int,
    status: _StatusWriter,
    *,
    checkpoint_callback: Callable[[object, dict], None] | None = None,
    start_idx: int = 0,
    excel_proc=None,
    per_call_timeout_sec: int = DEFAULT_PER_CALL_TIMEOUT_SEC,
) -> tuple[str | None, bool, str | None, int | None]:
    """Chunked macro path: Init + per-project SolveOneProjectByColHL + Finalize.

    Each project becomes its own COM Application.Run call so a long
    cold-start solve cannot push a single COM call past Excel's ~900s
    RPC timeout. Status is updated between projects so the dashboard
    can show live "N of M" progress.

    When `checkpoint_callback` is supplied, the helper reads that
    project's row from __SolverResults right after it lands and fires
    the callback with the parsed dict. Callback exceptions are caught
    and logged but do not stop the solve.

    Returns (macro_used, has_switch, error). has_switch is always True
    here since the chunked entry points only live in the SolveHeadless
    module, alongside SwitchProjectAndRecalc.
    """
    macro_used = "SolveHeadless"  # The chunked entry points live in this module
    has_switch = True
    # Init only on a fresh run, not on resume — InitSolveEnvHL resets the
    # __SolverResults sheet and clears prior per-project results. On resume
    # we want the already-converged rows preserved.
    if start_idx == 0:
        try:
            # InitSolveEnvHL resets __SolverResults and primes calc tier
            # state. A prior run that crashed mid-write can leave the
            # sheet in a state where Init spins on teardown; the
            # watchdog caps that at the per-call timeout.
            _run_macro_with_timeout(
                excel,
                (vba_call_str(wb.Name, INIT_SOLVE_ENV),),
                timeout_sec=per_call_timeout_sec,
                excel_proc=excel_proc,
                label="InitSolveEnvHL",
            )
        except Exception as e:
            return macro_used, has_switch, f"InitSolveEnvHL failed: {e}", None

    n = len(norm_tasks)
    chunked_error: str | None = None
    failed_idx: int | None = None
    if start_idx > 0:
        log.info("  Resuming chunked solve from project %d/%d", start_idx + 1, n)
    for idx in range(start_idx, n):
        nt = norm_tasks[idx]
        col_idx = nt.offset + BASE_COL
        results_row = idx + 2  # row 1 holds headers
        log.info("  [%d/%d] Solving %s (col %d)...", idx + 1, n, nt.name, col_idx)
        # Two-phase status writes per project let an external observer
        # tell "Python invoked COM, waiting on Excel" from "Excel
        # returned, Python prepping next" — without parsing VBA
        # heartbeats. Pre/post the COM call carry call_phase="invoking"
        # / "returned" respectively.
        call_start = time.time()
        status.update(
            "solving",
            per_project_status="solving",
            current_index=idx + 1,
            current_total=n,
            current_project=nt.name,
            macro_used=macro_used,
            chunked=True,
            call_phase="invoking",
            call_col=int(col_idx),
        )
        try:
            _run_macro_with_timeout(
                excel,
                (
                    vba_call_str(wb.Name, SOLVE_ONE_PROJECT_BY_COL),
                    int(col_idx),
                    str(nt.name),
                    int(results_row),
                ),
                timeout_sec=per_call_timeout_sec,
                excel_proc=excel_proc,
                label=f"SolveOneProjectByColHL[{nt.name}]",
            )
            status.update(
                "solving",
                per_project_status="solving",
                current_index=idx + 1,
                current_total=n,
                current_project=nt.name,
                macro_used=macro_used,
                chunked=True,
                call_phase="returned",
                call_col=int(col_idx),
                last_call_secs=round(time.time() - call_start, 2),
            )
        except TimeoutError as e:
            chunked_error = (
                f"SolveOneProjectByColHL TIMEOUT at "
                f"{nt.name} (col {col_idx}, row {results_row}, idx {idx}): {e}"
            )
            failed_idx = idx
            break
        except Exception as e:
            chunked_error = (
                f"SolveOneProjectByColHL failed at "
                f"{nt.name} (col {col_idx}, row {results_row}, idx {idx}): {e}"
            )
            failed_idx = idx
            break

        if checkpoint_callback is not None:
            try:
                meta = _read_one_solver_result_row(wb, results_row)
                checkpoint_callback(nt, meta)
            except Exception as cb_exc:
                # Persistence failure must never stall the solve. Log and
                # carry on; the project's data is still in the workbook.
                # Include row index so forensic recovery from logs alone
                # can match the failure back to its __SolverResults row.
                log.warning(
                    "  Checkpoint callback failed for %s (row=%d): %s",
                    nt.name, results_row, cb_exc,
                )

    # Finalize is best-effort — it restores F2 and re-enables non-core
    # sheets, so we always try to call it even after a per-project error.
    # Finalize also gets the timeout watchdog: a finalize that hangs would
    # have the same silent-stall failure mode as a per-project hang.
    try:
        _run_macro_with_timeout(
            excel,
            (vba_call_str(wb.Name, FINALIZE_SOLVE_ENV), int(original_f2)),
            timeout_sec=per_call_timeout_sec,
            excel_proc=excel_proc,
            label="FinalizeSolveEnvHL",
        )
    except Exception as e:
        if chunked_error is None:
            chunked_error = f"FinalizeSolveEnvHL failed: {e}"

    return macro_used, has_switch, chunked_error, failed_idx


def _parse_solver_result_row(row_vals: tuple) -> dict:
    """Decode one __SolverResults row tuple into the canonical result dict.

    Single source of truth for the cols A-T schema. Callers pass the
    raw tuple from `Range.Value` (already flattened from the
    `((v1, v2, ...),)` single-row shape if needed). Older workbook
    builds without the phase-telemetry columns return short rows; the
    `len(row_vals) > N` guards keep those cases None-valued rather
    than IndexError.
    """
    tier_raw = row_vals[19] if len(row_vals) > 19 else None
    return {
        "project_name": safe_str_or_float(row_vals[1]),
        "dscr": safe_float(row_vals[2]),
        "npp": safe_float(row_vals[3]),
        "dev_fee": safe_float(row_vals[4]),
        "equity_pct": safe_float(row_vals[5]),
        "irr_gap": safe_float(row_vals[6]),
        "appr_gap": safe_float(row_vals[7]),
        "converged_flag": safe_str_or_float(row_vals[8]),
        "calc_tier": safe_str_or_float(row_vals[9]),
        "gs_retry_limit": safe_str_or_float(row_vals[10]),
        "mode": safe_str_or_float(row_vals[11]),
        "solve_seconds": safe_float(row_vals[12]),
        "heartbeat": safe_str_or_float(row_vals[13]),
        "calc_secs_dscr": safe_float(row_vals[14]) if len(row_vals) > 14 else None,
        "calc_secs_npp": safe_float(row_vals[15]) if len(row_vals) > 15 else None,
        "calc_secs_appr": safe_float(row_vals[16]) if len(row_vals) > 16 else None,
        "calc_secs_full": safe_float(row_vals[17]) if len(row_vals) > 17 else None,
        "iterations": _to_int(row_vals[18]) if len(row_vals) > 18 else None,
        "conv_tier": tier_raw if isinstance(tier_raw, str) else "none",
    }


def _read_one_solver_result_row(wb: object, results_row: int) -> dict:
    """Read a single __SolverResults row used by the chunked checkpoint
    hook so each callback fires with the same shape the post-solve bulk
    read produces.
    """
    try:
        ws = wb.Sheets(SOLVER_RESULTS_SHEET)
        row_vals = ws.Range(f"A{results_row}:T{results_row}").Value
    except Exception:
        return {}
    if row_vals is None:
        return {}
    if row_vals and isinstance(row_vals[0], tuple):
        row_vals = row_vals[0]
    if len(row_vals) < 14:
        return {}
    return _parse_solver_result_row(row_vals)


def _read_cell(wb: object, sheet: str, address: str) -> float | str | None:
    return safe_value(wb.Sheets(sheet).Range(address).Value)


def _set_f2(wb: object, idx_cell: object, offset: int) -> None:
    """Set F2 directly (fallback when SwitchProjectAndRecalc unavailable)."""
    if hasattr(idx_cell, "sheet"):
        wb.Sheets(idx_cell.sheet).Range(idx_cell.address).Value = offset
    else:
        wb.Sheets(idx_cell["sheet"]).Range(idx_cell["address"]).Value = offset


def _read_solver_results_map(
    wb: object,
) -> tuple[dict[int, dict[str, float | str | None]], str | None]:
    """Read per-project solve telemetry captured by SolveHeadless VBA.

    One bulk Range.Value read covering A2:T{2 + _RESULTS_BULK_ROWS - 1} replaces
    the prior per-cell while loop (~20 COM round-trips per project, ~1200 for a
    60-project portfolio). The block is sparse-tolerant: rows whose A column is
    blank are treated as end-of-data.

    Columns A–N are the original schema (offset, name, DSCR, NPP, Dev Fee,
    equity_pct, gaps, converged, calc tier, retry limit, mode, solve_seconds,
    heartbeat). Columns O–R carry per-phase calc-time telemetry written by
    CalcForPhase: cumulative seconds spent recalculating in the DSCR / NPP /
    Appraisal / Full scopes for that project. Column S is the actual outer-loop
    iteration count captured by the solve path. Column T is the convergence
    tier ("strict" / "relaxed" / "none") written by ClassifyConvergenceHL.
    """
    out: dict[int, dict[str, float | str | None]] = {}
    try:
        ws = wb.Sheets(SOLVER_RESULTS_SHEET)
    except Exception:
        return out, None

    heartbeat = safe_str_or_float(ws.Range("N1").Value)
    if not isinstance(heartbeat, str):
        heartbeat = None

    last_row = 1 + _RESULTS_BULK_ROWS
    try:
        block = ws.Range(f"A2:T{last_row}").Value
    except Exception:
        return out, heartbeat
    if block is None:
        return out, heartbeat

    # Single-row Range.Value comes back as a flat tuple; multi-row as
    # tuple-of-tuples. Normalize to the latter.
    if block and not isinstance(block[0], tuple):
        block = (block,)

    for row_vals in block:
        offset_raw = row_vals[0]
        if offset_raw is None or offset_raw == "":
            break
        try:
            offset = int(offset_raw)
        except (ValueError, TypeError):
            continue
        out[offset] = _parse_solver_result_row(row_vals)
    return out, heartbeat


def _to_int(v: object) -> int | None:
    """Convert a COM scalar to int, returning None on failure.

    Modeled on safe_float — VBA writes an Integer to column S, but the
    pywin32 marshaller surfaces it as float when the cell is part of a
    Range.Value bulk read. Coerce via float to handle both cases without
    raising on a None / blank row.
    """
    if v is None or v == "":
        return None
    with contextlib.suppress(ValueError, TypeError):
        return int(float(v))
    return None
