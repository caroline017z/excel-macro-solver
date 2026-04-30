"""dn38_solver.solver.orchestrator — Main hybrid shadow solve loop.

KEY OPTIMIZATION: Sends all projects to ONE COM subprocess that opens
Excel once, solves all projects by switching F2, then closes.
This mirrors the VBA macro's approach and avoids cold-start per project.
"""
from __future__ import annotations

import logging
import time
import uuid
from pathlib import Path

from dn38_solver.types import (
    ProjectInfo,
    ProjectResult,
    RunRecord,
    SolveStatus,
)
from dn38_solver.config import (
    OUTPUT_ROWS,
)
from dn38_solver.convert import safe_float
from dn38_solver.shadow.reader import WorkbookReader
from dn38_solver.shadow.validation import (
    format_validation_report,
    scan_workbook_errors,
)
from dn38_solver.com.direct_runner import run_direct
from dn38_solver.solver.sequence import build_solve_task
from dn38_solver.storage.database import get_connection, now_iso, save_run

log = logging.getLogger(__name__)


def _parse_project_result(
    project: ProjectInfo,
    raw: dict,
) -> ProjectResult:
    """Map a raw project result dict to a ProjectResult struct."""
    sv = raw.get("solved_values", {})
    col = project.col_letter

    def get(sheet: str, addr: str) -> float | None:
        key = f"{sheet}!{addr}"
        val = sv.get(key)
        return safe_float(val) if val is not None else None

    eq_val = safe_float(sv.get("PT Returns!C128"))
    uses_val = safe_float(sv.get("PT Returns!C130"))
    eq_pct = (
        eq_val / uses_val
        if eq_val is not None and uses_val is not None and uses_val != 0
        else None
    )

    return ProjectResult(
        name=project.name,
        col=project.col,
        col_letter=col,
        npp_per_w=get("Project Inputs", f"{col}38"),
        npp_total=get("Project Inputs", f"{col}39"),
        dev_fee_per_w=get("Project Inputs", f"{col}32"),
        fmv_per_w=get("Project Inputs", f"{col}33"),
        target_irr=get("Project Inputs", "F36"),
        live_irr=get("Project Inputs", "F37"),
        appraisal_live=get("Project Inputs", "F31"),
        wacc_target=get("Project Inputs", "F30"),
        dscr_multiple=get("PT Returns", "F129"),
        equity_pct=eq_pct,
        converged=raw.get("status") == "converged",
        iterations=raw.get("iterations_used", 0),
    )


def solve_all(
    workbook_path: Path,
    *,
    batch_id: str | None = None,
    dry_run: bool = False,
    timeout_sec: int = 600,
    strict_validation: bool = False,
) -> RunRecord:
    """Main entry point for the Hybrid Shadow solver.

    1. Pre-flight validation (openpyxl error scan) — fail-fast on broken input
    2. Shadow read (openpyxl) — extract projects, read current values
    3. Build SolveTasks for all projects
    4. Send ALL tasks to ONE COM subprocess (single Excel session)
    5. Parse results, persist to SQLite

    Args:
        strict_validation: when True, pre-flight or post-export formula
            errors abort the run with status=ERROR. When False (default),
            errors are logged as warnings but the solve proceeds.
    """
    if batch_id is None:
        batch_id = uuid.uuid4().hex[:8]

    start = time.time()

    log.info("=" * 60)
    log.info("  38DN Hybrid Shadow Solver")
    log.info("  Workbook: %s", workbook_path.name)
    log.info("  Mode: %s", "DRY RUN" if dry_run else "LIVE")
    log.info("  Batch: %s", batch_id)
    log.info("=" * 60)

    # Phase 0: Pre-flight formula-error scan on the input workbook.
    # Cheap (~1s on the IL pricing model) and catches a class of broken
    # inputs before we pay the COM startup + multi-minute solve cost.
    log.info("[Phase 0] Pre-flight formula-error scan...")
    pre_validation = scan_workbook_errors(workbook_path)
    log.info("  %s", format_validation_report(pre_validation, "Input"))
    if not pre_validation.ok and strict_validation:
        return RunRecord(
            workbook_name=workbook_path.name,
            run_timestamp=now_iso(),
            batch_id=batch_id,
            solver_mode="hybrid_shadow",
            projects=(),
            total_duration_sec=time.time() - start,
            status=SolveStatus.ERROR.value,
            error=(
                f"Pre-flight validation failed: {pre_validation.total_errors} "
                f"formula error(s) in input workbook"
            ),
        )

    # Phase 1: Shadow pre-read
    log.info("[Phase 1] Reading workbook with openpyxl...")
    with WorkbookReader(workbook_path) as reader:
        projects = reader.extract_active_projects()
        # Read original F2 to restore after solve
        original_f2 = reader.cell_value("Project Inputs", 2, 6)

    if not projects:
        log.warning("No active projects found (row 7 toggle = 1)")
        return RunRecord(
            workbook_name=workbook_path.name,
            run_timestamp=now_iso(),
            batch_id=batch_id,
            solver_mode="hybrid_shadow",
            projects=(),
            total_duration_sec=time.time() - start,
            status=SolveStatus.ERROR.value,
            error="No active projects found",
        )

    log.info("  Found %d active project(s):", len(projects))
    for p in projects:
        log.info("    - %s (col %s, offset %d)", p.name, p.col_letter, p.offset)

    if dry_run:
        log.info("[DRY RUN] Current values (no COM):")
        with WorkbookReader(workbook_path) as reader:
            for p in projects:
                outputs = reader.read_output_rows(p.col)
                log.info("  %s:", p.name)
                for row, label in OUTPUT_ROWS.items():
                    val = outputs.get(row)
                    log.info("    %s: %s", label, f"{val:.4f}" if val is not None else "\u2014")
        return RunRecord(
            workbook_name=workbook_path.name,
            run_timestamp=now_iso(),
            batch_id=batch_id,
            solver_mode="hybrid_shadow",
            projects=(),
            total_duration_sec=time.time() - start,
            status=SolveStatus.DRY_RUN.value,
        )

    # Phase 2: Build tasks for ALL projects
    tasks = [build_solve_task(p, str(workbook_path)) for p in projects]

    # Phase 3: Run VBA macro via direct COM (no subprocess)
    log.info("[Phase 2] Running VBA macro via direct COM (%d projects)...", len(projects))
    batch_result = run_direct(
        workbook_path=str(workbook_path),
        tasks=tasks,
        original_f2=int(original_f2) if original_f2 else 1,
        timeout_sec=timeout_sec,
    )

    # Phase 4: Parse results
    project_results: list[ProjectResult] = []
    raw_results = batch_result.get("project_results", [])

    if len(raw_results) != len(projects):
        log.warning(
            "Result count mismatch: expected %d projects, got %d results",
            len(projects), len(raw_results),
        )

    for project, raw in zip(projects, raw_results):
        pr = _parse_project_result(project, raw)
        project_results.append(pr)

        match raw.get("status"):
            case "converged":
                log.info("  %s: CONVERGED in %d iter (%.1fs)",
                         pr.name, pr.iterations, raw.get("duration_sec", 0))
            case "not_converged":
                log.warning("  %s: NOT CONVERGED after %d iter", pr.name, pr.iterations)
            case _:
                log.error("  %s: %s", pr.name, raw.get("status", "unknown"))

        log.info("    NPP=$%.4f  DevFee=$%.4f  FMV=$%.4f  DSCR=%.4fx",
                 pr.npp_per_w or 0, pr.dev_fee_per_w or 0,
                 pr.fmv_per_w or 0, pr.dscr_multiple or 0)

        meta = raw.get("meta") or {}
        phase_secs = (
            meta.get("calc_secs_dscr"),
            meta.get("calc_secs_npp"),
            meta.get("calc_secs_appr"),
            meta.get("calc_secs_full"),
        )
        if any(v is not None for v in phase_secs):
            dscr_s, npp_s, appr_s, full_s = (v or 0.0 for v in phase_secs)
            log.info(
                "    calc_secs: DSCR=%.1f  NPP=%.1f  Appr=%.1f  Full=%.1f  (total=%.1f)",
                dscr_s, npp_s, appr_s, full_s,
                dscr_s + npp_s + appr_s + full_s,
            )

    # Handle batch-level errors
    if batch_result.get("status") == "error" and not project_results:
        log.error("COM worker error: %s", batch_result.get("error"))
        return RunRecord(
            workbook_name=workbook_path.name,
            run_timestamp=now_iso(),
            batch_id=batch_id,
            solver_mode="hybrid_shadow",
            projects=(),
            total_duration_sec=time.time() - start,
            status=SolveStatus.ERROR.value,
            error=batch_result.get("error"),
        )

    # Post-export gate: scan the saved _SOLVED.xlsx for formula errors.
    post_validation = batch_result.get("validation")
    if post_validation is not None:
        log.info("  %s", format_validation_report(post_validation, "Solved"))

    all_converged = all(pr.converged for pr in project_results)
    post_failed = (
        strict_validation
        and post_validation is not None
        and not post_validation.ok
    )
    total_time = time.time() - start

    if post_failed:
        status_value = SolveStatus.ERROR.value
        error_msg = (
            f"Post-export validation failed: "
            f"{post_validation.total_errors} formula error(s) in saved workbook"
        )
    elif all_converged:
        status_value = SolveStatus.CONVERGED.value
        error_msg = None
    else:
        status_value = SolveStatus.NOT_CONVERGED.value
        error_msg = None

    record = RunRecord(
        workbook_name=workbook_path.name,
        run_timestamp=now_iso(),
        batch_id=batch_id,
        solver_mode="hybrid_shadow",
        projects=tuple(project_results),
        total_duration_sec=round(total_time, 2),
        status=status_value,
        error=error_msg,
    )

    # Persist
    conn = get_connection()
    row_id = save_run(conn, record)
    conn.close()

    # Summary
    log.info("=" * 60)
    log.info("  COMPLETE - %d project(s) in %.1fs (COM: %.1fs)",
             len(project_results), total_time, batch_result.get("duration_sec", 0))
    if batch_result.get("saved_to"):
        log.info("  Solved workbook: %s", batch_result["saved_to"])
    log.info("  Run id=%d | Status: %s", row_id, record.status)
    log.info("=" * 60)
    log.info("  %-28s %10s %10s %10s %8s %12s",
             "Project", "NPP $/W", "Dev Fee", "FMV", "DSCR", "Status")
    log.info("  %s %s %s %s %s %s", "-"*28, "-"*10, "-"*10, "-"*10, "-"*8, "-"*12)
    for r in project_results:
        npp = f"${r.npp_per_w:.3f}" if r.npp_per_w is not None else "\u2014"
        dev = f"${r.dev_fee_per_w:.3f}" if r.dev_fee_per_w is not None else "\u2014"
        fmv = f"${r.fmv_per_w:.3f}" if r.fmv_per_w is not None else "\u2014"
        dscr = f"{r.dscr_multiple:.2f}x" if r.dscr_multiple is not None else "\u2014"
        stat = "OK" if r.converged else "CHECK"
        log.info("  %-28s %10s %10s %10s %8s %12s", r.name, npp, dev, fmv, dscr, stat)

    return record
