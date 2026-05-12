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
    RELAXED_LEGEND,
    RunRecord,
    SolveStatus,
    convergence_label,
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
from dn38_solver.storage.database import (
    clear_project_checkpoints,
    get_connection,
    now_iso,
    save_project_checkpoint,
    save_run,
)

log = logging.getLogger(__name__)


def _parse_project_result(
    project: ProjectInfo,
    raw: dict,
) -> ProjectResult:
    """Map a raw project result dict to a ProjectResult struct.

    The convergence tier is sourced from __SolverResults!T (written by
    VBA's ClassifyConvergenceHL). Older runs that predate the column
    fall back to inferring the tier from the strict converged flag so
    historical SQLite rows render with sensible values rather than
    blanking out the new field.
    """
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

    is_converged = raw.get("status") == SolveStatus.CONVERGED.value
    meta = raw.get("meta") or {}
    tier_raw = meta.get("conv_tier")
    if isinstance(tier_raw, str) and tier_raw in {"strict", "relaxed", "none"}:
        tier = tier_raw
    else:
        # Pre-tier rows: infer from strict converged flag so older
        # SQLite-loaded ProjectResults still classify cleanly.
        tier = "strict" if is_converged else "none"

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
        converged=is_converged,
        convergence_tier=tier,
        iterations=raw.get("iterations_used", 0),
    )


def solve_all(
    workbook_path: Path,
    *,
    batch_id: str | None = None,
    dry_run: bool = False,
    timeout_sec: int = 600,
    strict_validation: bool = False,
    use_chunked: bool = False,
    allow_relaxed: bool = False,
    save_solved: bool = True,
    skip_output_recalc: bool = False,
    strip_sheets: tuple[str, ...] = (),
    workers: int = 1,
    excel_threads_per_worker: int | None = None,
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
        use_chunked: when True, run the macro through the per-project
            chunked entry points (InitSolveEnvHL / SolveOneProjectByColHL /
            FinalizeSolveEnvHL) instead of single-shot SolveHeadless. Each
            project is its own COM Application.Run call so no single
            invocation can exceed the ~900s COM RPC timeout — the win on
            cold portfolios that today crash before completion. Adds live
            per-project progress to the status JSON.
        allow_relaxed: when True, projects whose convergence tier is
            "relaxed" (equity within +/-1pp, inner gaps <= 5x tol) count
            toward the run-level CONVERGED status. Per-project
            ProjectResult.converged stays strict-only regardless. Default
            False preserves prior strict-only run-level reporting.
        save_solved: when True (default), persist a `<workbook>_SOLVED.xlsm`
            copy at end of run. Set False for fast iteration on Box-
            mounted workbooks where the save is a nontrivial fixed cost.
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

    # Open the DB up front so chunked checkpoints can be written as
    # projects converge, not only at the end of the run. The same
    # connection is reused for the final save_run write.
    conn = get_connection()

    # Build a project-by-name lookup so the checkpoint callback can
    # recover the col / col_letter for the ProjectResult struct from
    # whatever _NormTask is passed in.
    project_lookup = {p.name: p for p in projects}

    def _checkpoint(nt: object, meta: dict) -> None:
        """Persist a partial ProjectResult after each per-project solve.

        Wired through run_direct -> _run_chunked. Only called when the
        chunked path is in use (use_chunked=True). Reads only the
        scalars __SolverResults captured in-VBA; richer fields (target/
        live IRR, FMV, NPP $) are left None and will be filled in by
        the post-solve cell read at end of run.
        """
        raw_name = getattr(nt, "name", None)
        proj = project_lookup.get(raw_name) if raw_name else None
        if proj is None and isinstance(raw_name, str):
            # Trailing/leading whitespace from VBA is the likely culprit
            # for a name mismatch. Retry with the stripped form before
            # giving up so a cosmetic difference doesn't drop the row.
            proj = project_lookup.get(raw_name.strip())
        if proj is None:
            log.warning(
                "Checkpoint name lookup miss (%r) — partial result not "
                "persisted for this project; subsequent projects continue.",
                raw_name,
            )
            return
        converged_flag = meta.get("converged_flag")
        tier_raw = meta.get("conv_tier")
        tier = tier_raw if isinstance(tier_raw, str) and tier_raw in {
            "strict", "relaxed", "none"
        } else "none"
        partial = ProjectResult(
            name=proj.name,
            col=proj.col,
            col_letter=proj.col_letter,
            npp_per_w=safe_float(meta.get("npp")),
            dev_fee_per_w=safe_float(meta.get("dev_fee")),
            dscr_multiple=safe_float(meta.get("dscr")),
            equity_pct=safe_float(meta.get("equity_pct")),
            converged=bool(converged_flag) if converged_flag is not None else False,
            convergence_tier=tier,
        )
        save_project_checkpoint(
            conn,
            batch_id=batch_id,
            workbook_name=workbook_path.name,
            project=partial,
        )

    # Phase 3: Run VBA macro via direct COM (no subprocess)
    log.info("[Phase 2] Running VBA macro via direct COM (%d projects)...", len(projects))
    if workers > 1:
        # Parallel path: each worker is its own subprocess with its own
        # Excel COM session and a round-robin slice of projects. Parent
        # merges per-project output cells from each worker's _SOLVED.xlsm
        # into a single canonical _SOLVED.xlsm next to the source.
        # Portfolio-level aggregates may be stale post-merge (acceptable
        # per Issue #8 design — Excel recalcs them on next interactive
        # open since the project-column cells are correct).
        from dn38_solver.com.parallel_runner import run_parallel
        batch_result = run_parallel(
            workbook_path=str(workbook_path),
            tasks=tasks,
            workers=workers,
            original_f2=int(original_f2) if original_f2 else 1,
            timeout_sec=timeout_sec,
            use_chunked=use_chunked,
            skip_output_recalc=skip_output_recalc,
            strip_sheets=strip_sheets,
            excel_threads_per_worker=excel_threads_per_worker,
        )
        # Persist per-project checkpoints for the parallel path. The chunked
        # single-worker path uses an in-flight checkpoint_callback; parallel
        # workers can't share the parent's SQLite connection across processes,
        # so we batch-write here after run_parallel returns. Restores the
        # documented `--show-checkpoints <batch_id>` forensics path which
        # would otherwise show empty for any parallel run.
        for raw in batch_result.get("project_results", []):
            project_obj = project_lookup.get(raw.get("project_name"))
            if project_obj is None:
                continue
            partial = _parse_project_result(project_obj, raw)
            try:
                save_project_checkpoint(
                    conn,
                    batch_id=batch_id,
                    workbook_name=workbook_path.name,
                    project=partial,
                )
            except Exception as cp_exc:
                log.warning(
                    "Parallel checkpoint write failed for %s: %s",
                    project_obj.name, cp_exc,
                )
    else:
        batch_result = run_direct(
            workbook_path=str(workbook_path),
            tasks=tasks,
            original_f2=int(original_f2) if original_f2 else 1,
            timeout_sec=timeout_sec,
            use_chunked=use_chunked,
            checkpoint_callback=_checkpoint if use_chunked else None,
            save_solved=save_solved,
            skip_output_recalc=skip_output_recalc,
            strip_sheets=strip_sheets,
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

        meta = raw.get("meta") or {}
        # The chunked path hardcodes raw["duration_sec"]=0; the real per-
        # project time is captured by VBA in __SolverResults!M (surfaced
        # as meta["solve_seconds"]). Fall back to raw["duration_sec"] for
        # the single-shot path's batch-level number.
        solve_secs = meta.get("solve_seconds") or raw.get("duration_sec") or 0.0

        match raw.get("status"):
            case SolveStatus.CONVERGED.value:
                # tier is always "strict" here (status=converged implies
                # bConverged from VBA, which only fires on strict).
                log.info("  %s: CONVERGED in %d iter (%.1fs)",
                         pr.name, pr.iterations, solve_secs)
            case SolveStatus.NOT_CONVERGED.value:
                # Distinguish relaxed-tier hits from genuine non-convergence
                # so the log reflects what's actually investment-grade vs.
                # what needs a re-solve.
                if pr.convergence_tier == "relaxed":
                    log.info(
                        "  %s: CONVERGED (relaxed) in %d iter (%.1fs)",
                        pr.name, pr.iterations, solve_secs,
                    )
                else:
                    log.warning("  %s: NOT CONVERGED after %d iter", pr.name, pr.iterations)
            case _:
                log.error("  %s: %s", pr.name, raw.get("status", "unknown"))

        log.info("    NPP=$%.4f  DevFee=$%.4f  FMV=$%.4f  DSCR=%.4fx",
                 pr.npp_per_w or 0, pr.dev_fee_per_w or 0,
                 pr.fmv_per_w or 0, pr.dscr_multiple or 0)

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

        # Tier 3 = Application.CalculateFull fallback. Reaching it means
        # the cheaper recalc tiers didn't propagate dirty dependents
        # cleanly. Surface it so a regression in cold-start convergence
        # isn't buried in __SolverResults!J.
        try:
            calc_tier = int(meta.get("calc_tier") or 0)
        except (TypeError, ValueError):
            calc_tier = 0
        if calc_tier >= 3:
            log.warning(
                "    %s reached Tier 3 (CalculateFull) — cheaper recalc "
                "tiers did not propagate cleanly; check model for new "
                "OFFSET/INDIRECT volatility.", pr.name,
            )

    # Handle batch-level errors
    if batch_result.get("status") == SolveStatus.ERROR.value and not project_results:
        log.error("COM worker error: %s", batch_result.get("error"))
        conn.close()
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

    # Run-level convergence applies the --allow-relaxed policy. The
    # per-project ProjectResult.converged field stays strict-only --
    # downstream consumers reading individual records get the unchanged
    # strict semantics. Only the rolled-up run status is affected here.
    def _ok_at_run_level(pr: ProjectResult) -> bool:
        if pr.converged:
            return True
        return allow_relaxed and pr.convergence_tier == "relaxed"

    all_converged = all(_ok_at_run_level(pr) for pr in project_results)
    post_failed = (
        strict_validation
        and post_validation is not None
        and not post_validation.ok
    )
    # batch_error is set when the macro itself raised but partial data
    # was still salvaged from __SolverResults (chunked Finalize failure,
    # mid-portfolio crash). The run is marked ERROR so checkpoints are
    # retained for forensics, even though some / all projects converged.
    batch_error = batch_result.get("error")
    total_time = time.time() - start

    if post_failed:
        status_value = SolveStatus.ERROR.value
        error_msg = (
            f"Post-export validation failed: "
            f"{post_validation.total_errors} formula error(s) in saved workbook"
        )
    elif batch_error:
        status_value = SolveStatus.ERROR.value
        error_msg = batch_error
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

    # Persist (uses the connection opened earlier so chunked checkpoints
    # already written under the same batch_id stay alongside the final
    # RunRecord).
    row_id = save_run(conn, record)
    if use_chunked and record.status == SolveStatus.CONVERGED.value:
        # Run finished cleanly — drop the per-project checkpoint rows
        # so the audit trail only retains incidents worth keeping.
        cleared = clear_project_checkpoints(conn, batch_id)
        if cleared:
            log.debug("  Cleared %d per-project checkpoint(s) on clean run", cleared)
    conn.close()

    # Summary
    log.info("=" * 60)
    log.info("  COMPLETE - %d project(s) in %.1fs (COM: %.1fs)",
             len(project_results), total_time, batch_result.get("duration_sec", 0))

    # Parallel-mode signals: speedup vs estimated sequential and the merge
    # path that produced the canonical _SOLVED.xlsm. Both are no-ops in
    # single-worker mode (run_direct doesn't return them).
    if workers > 1:
        est_seq = batch_result.get("estimated_sequential_sec")
        if est_seq and total_time > 0:
            speedup = est_seq / total_time
            log.info(
                "  Parallel speedup: %.2fx (%.1fs parallel vs %.1fs estimated sequential)",
                speedup, total_time, est_seq,
            )
        merge_path = batch_result.get("merge_path")
        if merge_path == "openpyxl":
            log.info("  Merge path: openpyxl (per-project columns authoritative)")
        elif merge_path == "vba_fallback":
            log.warning(
                "  Merge path: VBA-helper fallback used \u2014 openpyxl could not "
                "round-trip the macro project. File is correct; flag if recurring."
            )
        elif merge_path == "copy_master":
            log.error(
                "  Merge path: copy_master fallback \u2014 only worker-0's columns "
                "are converged. Consult per-worker outputs in the temp dir."
            )

    if batch_result.get("saved_to"):
        log.info("  Solved workbook: %s", batch_result["saved_to"])
    log.info("  Run id=%d | Status: %s", row_id, record.status)

    # Convergence-tier rollup so the user sees at a glance how many
    # projects hit strict vs relaxed vs no convergence \u2014 one number per
    # tier rather than scanning the per-project table.
    tier_counts: dict[str, int] = {"strict": 0, "relaxed": 0, "none": 0}
    for r in project_results:
        tier_counts[r.convergence_tier] = tier_counts.get(r.convergence_tier, 0) + 1
    log.info(
        "  Convergence: %d strict / %d relaxed / %d none (of %d total)",
        tier_counts["strict"], tier_counts["relaxed"], tier_counts["none"],
        len(project_results),
    )

    log.info("=" * 60)
    log.info("  %-28s %10s %10s %10s %8s %12s",
             "Project", "NPP $/W", "Dev Fee", "FMV", "DSCR", "Status")
    log.info("  %s %s %s %s %s %s", "-"*28, "-"*10, "-"*10, "-"*10, "-"*8, "-"*12)
    has_relaxed = False
    for r in project_results:
        npp = f"${r.npp_per_w:.3f}" if r.npp_per_w is not None else "\u2014"
        dev = f"${r.dev_fee_per_w:.3f}" if r.dev_fee_per_w is not None else "\u2014"
        fmv = f"${r.fmv_per_w:.3f}" if r.fmv_per_w is not None else "\u2014"
        dscr = f"{r.dscr_multiple:.2f}x" if r.dscr_multiple is not None else "\u2014"
        stat = convergence_label(r)
        if stat == "OK*":
            has_relaxed = True
        log.info("  %-28s %10s %10s %10s %8s %12s", r.name, npp, dev, fmv, dscr, stat)
    if has_relaxed:
        log.info("  %s", RELAXED_LEGEND)

    return record
