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
import json
import logging
import shutil
import tempfile
import time
from pathlib import Path
from typing import Callable, NamedTuple

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

    __slots__ = ("_base", "_start", "_path")

    def __init__(
        self,
        *,
        workbook_path: str,
        proj_names: list[str],
        start: float,
        path: Path = STATUS_FILE,
    ) -> None:
        self._base = {
            "workbook": workbook_path,
            "total_projects": len(proj_names),
            "_proj_names": proj_names,
        }
        self._start = start
        self._path = path

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
        payload.update(extras)
        try:
            # Atomic swap so a Streamlit reader can never observe a half-
            # written JSON. tmp file lives next to the target so .replace()
            # stays on the same filesystem.
            tmp_path = self._path.with_suffix(self._path.suffix + ".tmp")
            tmp_path.write_text(json.dumps(payload, default=str))
            tmp_path.replace(self._path)
        except OSError:
            pass


def run_direct(
    workbook_path: str,
    tasks: list[SolveTask],
    *,
    original_f2: int = 1,
    timeout_sec: int = 600,
    use_chunked: bool = False,
    checkpoint_callback: Callable[[object, dict], None] | None = None,
    save_solved: bool = True,
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
    """
    import pythoncom
    import win32com.client

    start = time.time()
    tmp_dir = Path(tempfile.mkdtemp(prefix="38dn_com_"))
    temp_path = tmp_dir / Path(workbook_path).name
    shutil.copy2(workbook_path, str(temp_path))

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
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.ScreenUpdating = False
        excel.EnableEvents = False

        # Normalize task payloads once so the per-project loop is branch-free.
        norm_tasks = [_norm_task(t) for t in tasks]
        proj_names = [nt.name for nt in norm_tasks]

        status = _StatusWriter(
            workbook_path=workbook_path,
            proj_names=proj_names,
            start=start,
        )
        status.update("opening", per_project_status="pending")

        log.info("  Opening workbook via COM...")
        wb = excel.Workbooks.Open(
            str(temp_path),
            ReadOnly=False,
            UpdateLinks=0,
        )
        open_time = time.time() - start
        log.info("  Opened in %.1fs", open_time)

        warmup_time = 0.0

        # --- Run the VBA macro ---
        macro_used: str | None = None
        macro_error: str | None = None

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
            macro_used, has_switch, macro_error = _run_chunked(
                excel, wb, norm_tasks, original_f2, status,
                checkpoint_callback=checkpoint_callback,
            )
        else:
            macro_used, has_switch, macro_error = _run_single_shot(excel, wb)

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
        log.info("  Reading results for %d project(s)...", len(tasks))
        status.update(
            "reading",
            per_project_status="reading",
            macro_used=macro_used,
            macro_time_sec=solve_time,
        )
        t0 = time.time()
        project_results = []
        solver_results, heartbeat = _read_solver_results_map(wb)

        for nt in norm_tasks:
            meta = solver_results.get(nt.offset)
            if meta is None:
                # Mid-portfolio failure stopped before this project ran.
                # Don't read its cells — they'd reflect either uninitialized
                # state or another project's converged values, both of
                # which would be misleading downstream.
                project_results.append({
                    "project_name": nt.name,
                    "status": "not_attempted",
                    "solved_values": {},
                    "iterations_used": 0,
                    "duration_sec": 0,
                    "meta": {},
                    "_summary": {"npp": None, "dev_fee": None, "fmv": None},
                })
                continue

            # Switch F2 with targeted recalc
            if has_switch:
                try:
                    excel.Application.Run(
                        f"'{wb.Name}'!SwitchProjectAndRecalc",
                        nt.offset,
                    )
                except Exception:
                    _set_f2(wb, nt.idx_cell, nt.offset)
            else:
                _set_f2(wb, nt.idx_cell, nt.offset)

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

            project_results.append({
                "project_name": nt.name,
                "status": "converged" if is_converged else "not_converged",
                "solved_values": solved,
                "iterations_used": int(meta.get("iterations") or 0),
                "duration_sec": 0,
                "meta": meta or {},
                "_summary": {"npp": npp, "dev_fee": dev_fee, "fmv": fmv},
            })

        read_time = time.time() - t0
        log.info("  Read %d project(s) in %.1fs", len(tasks), read_time)

        # Restore original F2
        if norm_tasks and has_switch:
            with contextlib.suppress(Exception):
                excel.Application.Run(
                    f"'{wb.Name}'!SwitchProjectAndRecalc",
                    int(original_f2),
                )

        # Save solved workbook (opt-out via save_solved=False for fast
        # iteration runs; the post-export validation gate only runs when
        # there's actually a saved file to scan).
        saved_to = None
        if save_solved:
            wb_path = Path(workbook_path)
            solved_name = wb_path.stem + "_SOLVED" + wb_path.suffix
            solved_path = wb_path.parent / solved_name
            with contextlib.suppress(Exception):
                wb.SaveAs(str(solved_path))
                saved_to = str(solved_path)

        # Post-export formula-error gate: scan the just-saved file for
        # cached Excel error tokens (#REF! / #DIV/0! / #VALUE! / etc.).
        # Pure-Python via openpyxl — no LibreOffice dependency.
        validation = None
        if saved_to is not None:
            with contextlib.suppress(Exception):
                validation = scan_workbook_errors(saved_to)

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
) -> tuple[str | None, bool, str | None]:
    """Legacy single-invocation macro path. Tries SolveHeadless first;
    falls back to the original SolveMinEquityWithHoldCo if the headless
    wrapper isn't present in the workbook (older models).

    Returns (macro_used, has_switch, error). has_switch is True when the
    SwitchProjectAndRecalc helper lives alongside the macro that ran --
    only the SolveHeadless module ships it, so the legacy fallback path
    reports False and the post-solve read uses the F2-direct fallback.
    """
    macro_names = ("SolveHeadless", "SolveMinEquityWithHoldCo")
    for macro_name in macro_names:
        try:
            excel.Application.Run(f"'{wb.Name}'!{macro_name}")
            return macro_name, macro_name == "SolveHeadless", None
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
) -> tuple[str | None, bool, str | None]:
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
    try:
        excel.Application.Run(f"'{wb.Name}'!InitSolveEnvHL")
    except Exception as e:
        return macro_used, has_switch, f"InitSolveEnvHL failed: {e}"

    n = len(norm_tasks)
    chunked_error: str | None = None
    for idx, nt in enumerate(norm_tasks):
        col_idx = nt.offset + BASE_COL
        results_row = idx + 2  # row 1 holds headers
        log.info("  [%d/%d] Solving %s (col %d)...", idx + 1, n, nt.name, col_idx)
        status.update(
            "solving",
            per_project_status="solving",
            current_index=idx + 1,
            current_total=n,
            current_project=nt.name,
            macro_used=macro_used,
            chunked=True,
        )
        try:
            excel.Application.Run(
                f"'{wb.Name}'!SolveOneProjectByColHL",
                int(col_idx),
                str(nt.name),
                int(results_row),
            )
        except Exception as e:
            chunked_error = (
                f"SolveOneProjectByColHL failed at "
                f"{nt.name} (col {col_idx}, row {results_row}): {e}"
            )
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
    try:
        excel.Application.Run(
            f"'{wb.Name}'!FinalizeSolveEnvHL",
            int(original_f2),
        )
    except Exception as e:
        if chunked_error is None:
            chunked_error = f"FinalizeSolveEnvHL failed: {e}"

    return macro_used, has_switch, chunked_error


def _read_one_solver_result_row(wb: object, results_row: int) -> dict:
    """Read a single row from __SolverResults in one COM round-trip.

    Mirrors the bulk read schema (cols A–S) but for a single project,
    used by the chunked checkpoint hook so each callback fires with the
    same shape the post-solve bulk read produces.
    """
    try:
        ws = wb.Sheets(SOLVER_RESULTS_SHEET)
        row_vals = ws.Range(f"A{results_row}:T{results_row}").Value
    except Exception:
        return {}
    if row_vals is None:
        return {}
    # Single-row Range.Value comes back as ((v1, v2, ...),) — flatten.
    if row_vals and isinstance(row_vals[0], tuple):
        row_vals = row_vals[0]
    if len(row_vals) < 14:
        return {}
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
        # Older workbook builds without the phase-telemetry columns return
        # short rows (or None in O–R); use safe_float so a missing value
        # surfaces as None rather than raising.
        tier_raw = row_vals[19] if len(row_vals) > 19 else None
        out[offset] = {
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
