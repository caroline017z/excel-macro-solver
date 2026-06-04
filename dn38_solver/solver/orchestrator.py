"""dn38_solver.solver.orchestrator — Main hybrid shadow solve loop.

KEY OPTIMIZATION: Sends all projects to ONE COM subprocess that opens
Excel once, solves all projects by switching F2, then closes.
This mirrors the VBA macro's approach and avoids cold-start per project.
"""
from __future__ import annotations

import contextlib
import logging
import time
import uuid
from pathlib import Path

from dn38_solver.types import (
    ProjectInfo,
    ProjectResult,
    RELAXED_LEGEND,
    RunMetrics,
    RunRecord,
    SolveStatus,
    convergence_label,
)
from dn38_solver.config import (
    OUTPUT_ROWS,
)
from dn38_solver.convert import safe_float
from dn38_solver.shadow.reader import WorkbookReader
from dn38_solver.shadow.preflight import (
    apply_auto_fixes,
    check_macro_hash,
    check_macro_signatures,
    check_macro_version,
    format_preflight_report,
    run_preflight,
)
from dn38_solver.shadow.validation import (
    format_validation_report,
)
from dn38_solver.com.direct_runner import run_direct
from dn38_solver.solver.sequence import build_solve_task
from dn38_solver.storage.database import (
    clear_project_checkpoints,
    get_connection,
    now_iso,
    save_project_checkpoint,
    save_run,
    save_run_metrics,
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
    # ConvergenceTier Literal validates at struct-construction time, so
    # pre-validate here against the same set rather than letting msgspec
    # raise inside ProjectResult().
    _VALID_TIERS = {"strict", "relaxed", "none", "not_attempted", "skipped"}
    raw_status = raw.get("status")
    if raw_status == "not_attempted":
        # Worker crashed before reaching this project. Distinguish from
        # genuine non-convergence so the end-of-run rollup doesn't
        # overstate the failure rate (and so the speedup math can
        # exclude these from the wall-time-vs-sequential comparison).
        tier = "not_attempted"
    elif raw_status == "skipped":
        # VBA fast-skip bypassed this project pre-solve (placeholder
        # column with no RC1 revenue, or MWdc=0). Distinct from
        # not_attempted (worker crash) — these are deliberate bypasses
        # and shouldn't count against batch-level convergence rollup.
        tier = "skipped"
    elif isinstance(tier_raw, str) and tier_raw in _VALID_TIERS:
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
    strict_preflight: bool = False,
    auto_fix: bool = False,
    auto_import_macro: bool = False,
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

    # Phase 0: Bank-grade pre-flight pass.
    # Three categories of checks (calc props / structure / critical-path
    # errors / input bounds). Cheap (~2-3s on a 13MB IL pricing model)
    # vs the multi-minute COM cost we'd otherwise burn on a workbook that
    # can't converge. Errors abort the run unconditionally; warnings abort
    # only with --strict-preflight. --auto-fix patches fixable findings
    # into a sibling _FIXED.xlsm and proceeds against the patched copy.
    log.info("[Phase 0] Pre-flight checks...")
    preflight = run_preflight(workbook_path)
    log.info("\n%s", format_preflight_report(preflight))

    # Phase 0.5 — Auto-import macro into SOURCE workbook.
    # Destructive path: --auto-import-macro alone (without --auto-fix) is
    # the explicit "mutate the source" mode. Kept for back-compat with
    # callers that need the side effect (e.g., a one-off macro refresh on
    # a workbook the operator wants to keep as the canonical artifact).
    # The preferred path for everyday solves is --auto-fix, which routes
    # the same re-import into the _FIXED.xlsm sibling — see Phase 0.6.
    macro_import_codes = {"D15", "D17", "D18"}
    needs_macro_import = any(
        f.code in macro_import_codes for f in preflight.findings
    )
    if auto_import_macro and not auto_fix and needs_macro_import:
        try:
            from dn38_solver.com.auto_recovery import (
                AutoRecoveryUnavailable,
                reimport_macro_subprocess,
            )
            log.info(
                "  Auto-import-macro: re-importing macro into %s "
                "(D15/D17 fired in pre-flight)...",
                workbook_path.name,
            )
            reimport_macro_subprocess(workbook_path)
            # Macro re-import only mutates xl/vbaProject.bin and docProps/
            # custom.xml (the D15 function-presence target and the D17 hash
            # stamp). Every other pre-flight finding (A/B/C/E tiers) is
            # invariant under this mutation — the calcPr block, sheet
            # structure, cell values, RC config, etc. cannot have changed.
            # Re-running the full preflight here would burn another ~3-4
            # min of openpyxl loading for guaranteed-identical results.
            # Instead, re-check ONLY the D-tier (cheap zip-level reads,
            # ~10ms) and filter the stale D-codes out of the existing
            # preflight result.
            log.info("  Auto-import-macro: re-checking D15/D17/D18 against updated workbook...")
            new_d_findings: list = []
            new_d_findings.extend(check_macro_version(workbook_path))
            new_d_findings.extend(check_macro_hash(workbook_path))
            new_d_findings.extend(check_macro_signatures(workbook_path))
            d_tier_codes = {"D15", "D16", "D17", "D18"}
            filtered = tuple(
                f for f in preflight.findings if f.code not in d_tier_codes
            )
            preflight = type(preflight)(
                workbook_path=preflight.workbook_path,
                findings=filtered + tuple(new_d_findings),
                error_scan=preflight.error_scan,
            )
            log.info("\n%s", format_preflight_report(preflight))
        except AutoRecoveryUnavailable as ar_exc:
            log.error("  Auto-import-macro FAILED: %s", ar_exc)
            return RunRecord(
                workbook_name=workbook_path.name,
                run_timestamp=now_iso(),
                batch_id=batch_id,
                solver_mode="hybrid_shadow",
                projects=(),
                total_duration_sec=time.time() - start,
                status=SolveStatus.ERROR.value,
                error=(
                    f"Auto-import-macro failed: {ar_exc}. Close the workbook "
                    "if open in Excel, verify Trust Center 'Trust access to "
                    "the VBA project object model' is enabled, and retry."
                ),
            )

    # Phase 0.6 — Auto-fix sibling path.
    # Write _FIXED.xlsm and resume against it. Preserves the original
    # file path in the run record so the audit trail names the source
    # workbook, not the patched copy.
    #
    # Per Tranche 7.13, --auto-fix covers TWO classes of auto-fixable
    # findings:
    #   * A1  — workbook calcPr / iterateDelta missing or out-of-bound.
    #           Patched at the xl/workbook.xml zip layer (no COM).
    #   * D15 / D17 — embedded macro missing required Subs or .bas hash
    #           drift. Patched via reimport_macro_subprocess (Excel COM
    #           SaveAs against the sibling — the source remains
    #           untouched).
    #
    # Both paths target the _FIXED.xlsm sibling, never the source. After
    # the patch pass, we re-check only the affected codes (A1 from the
    # findings filter; D15/D17 by re-running the D-tier scan against the
    # patched file). We do NOT re-run the full preflight — its A/B/C/E
    # tiers are invariant under both A1 and D15/D17 mutations.
    original_workbook_path = workbook_path
    if auto_fix and preflight.auto_fixable:
        fixed_path = workbook_path.with_name(
            workbook_path.stem + "_FIXED" + workbook_path.suffix
        )
        auto_fixable_codes = {f.code for f in preflight.auto_fixable}
        log.info(
            "  Auto-fix: applying %d fixable finding(s) -> %s",
            len(preflight.auto_fixable), fixed_path.name,
        )
        fixed_path, applied_codes = apply_auto_fixes(
            workbook_path, fixed_path, preflight.findings
        )
        if applied_codes:
            log.info("  Auto-fix: applied codes %s (zip-layer)", applied_codes)

        # Macro re-import on the sibling. apply_auto_fixes already
        # copied the source to fixed_path (and patched A1 if needed),
        # so the sibling exists and is ready to be SaveAs'd by the
        # macro import subprocess. The source workbook is untouched.
        macro_codes_fired = auto_fixable_codes & macro_import_codes
        if macro_codes_fired:
            try:
                from dn38_solver.com.auto_recovery import (
                    AutoRecoveryUnavailable,
                    reimport_macro_subprocess,
                )
                log.info(
                    "  Auto-fix: re-importing macro into sibling %s "
                    "(%s fired in pre-flight)...",
                    fixed_path.name, ", ".join(sorted(macro_codes_fired)),
                )
                reimport_macro_subprocess(fixed_path)
                applied_codes = list(applied_codes) + sorted(macro_codes_fired)
                log.info(
                    "  Auto-fix: re-checking D15/D17/D18 against patched sibling..."
                )
                new_d_findings: list = []
                new_d_findings.extend(check_macro_version(fixed_path))
                new_d_findings.extend(check_macro_hash(fixed_path))
                new_d_findings.extend(check_macro_signatures(fixed_path))
                d_tier_codes = {"D15", "D16", "D17", "D18"}
                filtered = tuple(
                    f for f in preflight.findings if f.code not in d_tier_codes
                )
                # Re-stitch into a preflight result against the patched
                # path. The A/B/C/E findings still reference the source's
                # path attribute, but their codes are what gates the run.
                preflight = type(preflight)(
                    workbook_path=str(fixed_path),
                    findings=filtered + tuple(new_d_findings),
                    error_scan=preflight.error_scan,
                )
            except AutoRecoveryUnavailable as ar_exc:
                log.error("  Auto-fix macro re-import FAILED: %s", ar_exc)
                return RunRecord(
                    workbook_name=workbook_path.name,
                    run_timestamp=now_iso(),
                    batch_id=batch_id,
                    solver_mode="hybrid_shadow",
                    projects=(),
                    total_duration_sec=time.time() - start,
                    status=SolveStatus.ERROR.value,
                    error=(
                        f"Auto-fix macro re-import failed: {ar_exc}. Close "
                        "the workbook if open in Excel, verify Trust Center "
                        "'Trust access to the VBA project object model' is "
                        "enabled, and retry."
                    ),
                )

        workbook_path = fixed_path  # swap to the fixed copy for the solve

        # Filter the codes we just resolved out of the cached preflight
        # result so the bank-grade error gate below sees a clean slate.
        applied_set = set(applied_codes)
        preflight = type(preflight)(
            workbook_path=str(fixed_path),
            findings=tuple(f for f in preflight.findings if f.code not in applied_set),
            error_scan=preflight.error_scan,
        )
        log.info("\n%s", format_preflight_report(preflight))

    # Bank-grade gate: errors always block. Warnings block only if
    # --strict-preflight. The legacy --strict-validation flag continues
    # to gate on the formula-error scan only (back-compat); --strict-
    # preflight is the new comprehensive gate.
    if preflight.errors:
        codes = ", ".join(f.code for f in preflight.errors)
        return RunRecord(
            workbook_name=original_workbook_path.name,
            run_timestamp=now_iso(),
            batch_id=batch_id,
            solver_mode="hybrid_shadow",
            projects=(),
            total_duration_sec=time.time() - start,
            status=SolveStatus.ERROR.value,
            error=(
                f"Pre-flight FAILED: {len(preflight.errors)} blocking error(s) "
                f"({codes}). See preflight report above for remediation."
            ),
        )
    if strict_preflight and preflight.warnings:
        codes = ", ".join(f.code for f in preflight.warnings)
        return RunRecord(
            workbook_name=original_workbook_path.name,
            run_timestamp=now_iso(),
            batch_id=batch_id,
            solver_mode="hybrid_shadow",
            projects=(),
            total_duration_sec=time.time() - start,
            status=SolveStatus.ERROR.value,
            error=(
                f"Pre-flight strict mode: {len(preflight.warnings)} warning(s) "
                f"treated as failure ({codes}). Run without --strict-preflight "
                f"to proceed past warnings."
            ),
        )

    # Back-compat: --strict-validation still gates on the bare formula
    # scan. Kept alongside the new preflight gate so existing CI/scripts
    # don't break.
    pre_validation = preflight.error_scan
    if not pre_validation.ok and strict_validation:
        return RunRecord(
            workbook_name=original_workbook_path.name,
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
    try:

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
            try:
                batch_result = run_parallel(
                    workbook_path=str(workbook_path),
                    tasks=tasks,
                    workers=workers,
                    original_f2=int(original_f2) if original_f2 else 1,
                    timeout_sec=timeout_sec,
                    use_chunked=use_chunked,
                    save_solved=save_solved,
                    skip_output_recalc=skip_output_recalc,
                    strip_sheets=strip_sheets,
                    excel_threads_per_worker=excel_threads_per_worker,
                )
            except BaseException as exc:
                # KeyboardInterrupt is BaseException, not Exception — catch
                # it here so we still close the SQLite connection and emit
                # a RunRecord(ERROR) instead of leaking conn and dropping
                # the run from the audit trail. Re-raising would also leave
                # the user without the batch_id to diagnose with.
                log.exception("Parallel runner crashed (or interrupted)")
                with contextlib.suppress(Exception):
                    conn.close()
                return RunRecord(
                    workbook_name=workbook_path.name,
                    run_timestamp=now_iso(),
                    batch_id=batch_id,
                    solver_mode="hybrid_shadow",
                    projects=(),
                    total_duration_sec=time.time() - start,
                    status=SolveStatus.ERROR.value,
                    error=f"parallel runner failed: {type(exc).__name__}: {exc}",
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
        # conn.close handled by outer finally below
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
        # downstream consumers reading individual records get the
        # unchanged strict semantics. Only the rolled-up run status is
        # affected here.
        #
        # Skipped projects (VBA fast-skip for placeholders) are treated
        # as run-level OK so a workbook with 5 real + 10 placeholders
        # reports CONVERGED if all 5 reals converged. Operator already
        # sees SKIP* labels in the per-project table.
        def _ok_at_run_level(pr: ProjectResult) -> bool:
            if pr.converged:
                return True
            if pr.convergence_tier == "skipped":
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

        # Sidecar metrics (merge_path, parallel speedup, worker count).
        # Must NOT abort the solve on failure — the run is already in
        # solver_runs; metrics are nice-to-have audit detail. The whole
        # block is wrapped in try so a schema-version mismatch on an
        # ancient DB doesn't kill an otherwise-clean run.
        try:
            metrics = RunMetrics(
                run_id=row_id,
                workers_used=int(batch_result.get("workers_used") or workers or 1),
                merge_path=batch_result.get("merge_path"),
                estimated_sequential_sec=batch_result.get("estimated_sequential_sec"),
                wall_time_sec=round(total_time, 2),
            )
            save_run_metrics(conn, metrics)
        except Exception as metrics_exc:
            log.warning("RunMetrics persistence failed: %s", metrics_exc)
        if use_chunked and record.status == SolveStatus.CONVERGED.value:
            # Run finished cleanly — drop the per-project checkpoint rows
            # so the audit trail only retains incidents worth keeping.
            cleared = clear_project_checkpoints(conn, batch_id)
            if cleared:
                log.debug("  Cleared %d per-project checkpoint(s) on clean run", cleared)
        # conn.close handled by outer finally below

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
                # Caveat the speedup line on small portfolios \u2014 fixed startup
                # cost (worker spawn + Excel open + status threads) dominates
                # below ~4 projects and the bare number reads as "parallel is
                # broken" without context.
                caveat = ""
                if len(project_results) < 4:
                    caveat = (
                        "  (small portfolio \u2014 parallel overhead dominates; "
                        "expect <1x below ~4 projects)"
                    )
                log.info(
                    "  Parallel speedup: %.2fx "
                    "(%.1fs parallel vs %.1fs estimated sequential)%s",
                    speedup, total_time, est_seq, caveat,
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
                    "are converged. DO NOT SHIP the merged file; use the per-worker "
                    "_SOLVED.xlsm files in the parent_tmp directory (path was logged "
                    "by parallel_runner above) instead."
                )

        log.info("  Run id=%d | Status: %s", row_id, record.status)
        # Print "Solved workbook" path AFTER status so an ERROR run doesn't
        # tempt the user to copy the path before noticing the run failed.
        if batch_result.get("saved_to"):
            if record.status == SolveStatus.CONVERGED.value:
                log.info("  Solved workbook: %s", batch_result["saved_to"])
            else:
                log.warning(
                    "  Solved workbook (NOT ship-ready, see Status above): %s",
                    batch_result["saved_to"],
                )

        # Convergence-tier rollup so the user sees at a glance how many
        # projects hit strict vs relaxed vs no convergence \u2014 one number per
        # tier rather than scanning the per-project table.
        tier_counts: dict[str, int] = {
            "strict": 0, "relaxed": 0, "none": 0, "not_attempted": 0,
        }
        for r in project_results:
            tier_counts[r.convergence_tier] = tier_counts.get(r.convergence_tier, 0) + 1
        n_total = len(project_results)
        n_ship = tier_counts["strict"] + (
            tier_counts["relaxed"] if allow_relaxed else 0
        )
        # Lead with the IC-relevant number \u2014 Caroline's first question is
        # always "can I send this file?" Tier breakdown follows for context.
        log.info(
            "  Ship-ready: %d/%d projects%s",
            n_ship, n_total,
            "  (relaxed counted, --allow-relaxed)" if allow_relaxed else "",
        )
        log.info(
            "  Convergence: %d strict / %d relaxed / %d none / %d not_attempted",
            tier_counts["strict"], tier_counts["relaxed"],
            tier_counts["none"], tier_counts["not_attempted"],
        )
        if tier_counts["not_attempted"]:
            log.warning(
                "  %d project(s) were not attempted \u2014 likely a worker crash "
                "before that slice; check parent_tmp for forensics.",
                tier_counts["not_attempted"],
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
    finally:
        # Always close conn even if a KeyboardInterrupt or unhandled
        # exception fires anywhere in the body (per-project checkpoint
        # loop, _parse_project_result, save_run, summary block, etc.).
        # The previous structure only closed conn on selected paths and
        # leaked it on Ctrl+C - fixed in round-3 review.
        with contextlib.suppress(Exception):
            conn.close()
