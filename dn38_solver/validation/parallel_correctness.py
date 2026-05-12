"""dn38_solver.validation.parallel_correctness — Sequential vs parallel diff.

Runs `solve_all` twice on the same workbook (single-worker baseline, then
parallel) and compares per-project outputs. Issue #8 acceptance gate.

Fields compared:
- npp_per_w
- dev_fee_per_w
- fmv_per_w
- dscr_multiple
- live_irr
- appraisal_live
- equity_pct

Tolerance defaults to 1e-4 per the issue spec. Different projects can
move independently — we don't compare portfolio aggregates because per
Caroline's design, the merged _SOLVED.xlsm leaves those stale for Excel
to recompute on next interactive open.
"""
from __future__ import annotations

import logging
import uuid
from dataclasses import dataclass, field
from pathlib import Path

from dn38_solver.solver.orchestrator import solve_all
from dn38_solver.types import ProjectResult, RunRecord

log = logging.getLogger(__name__)

COMPARED_FIELDS: tuple[str, ...] = (
    "npp_per_w",
    "dev_fee_per_w",
    "fmv_per_w",
    "dscr_multiple",
    "live_irr",
    "appraisal_live",
    "equity_pct",
)


@dataclass
class FieldDiff:
    field_name: str
    sequential: float | None
    parallel: float | None
    abs_diff: float | None
    within_tolerance: bool


@dataclass
class ProjectDiff:
    name: str
    status: str  # "PASS" | "FAIL" | "MISSING_SEQ" | "MISSING_PAR"
    max_abs_diff: float | None
    field_diffs: list[FieldDiff] = field(default_factory=list)


@dataclass
class ValidationReport:
    workbook_name: str
    workers: int
    tolerance: float
    sequential_batch_id: str
    parallel_batch_id: str
    sequential_duration: float
    parallel_duration: float
    speedup: float
    project_diffs: list[ProjectDiff]
    all_pass: bool
    failing_projects: list[str]


def _by_name(record: RunRecord) -> dict[str, ProjectResult]:
    return {p.name: p for p in record.projects}


def _diff_one_project(
    name: str,
    seq: ProjectResult | None,
    par: ProjectResult | None,
    tolerance: float,
) -> ProjectDiff:
    if seq is None:
        return ProjectDiff(name=name, status="MISSING_SEQ", max_abs_diff=None)
    if par is None:
        return ProjectDiff(name=name, status="MISSING_PAR", max_abs_diff=None)

    diffs: list[FieldDiff] = []
    max_abs: float = 0.0
    all_within = True

    for fname in COMPARED_FIELDS:
        sval = getattr(seq, fname, None)
        pval = getattr(par, fname, None)
        if sval is None or pval is None:
            # Treat missing as a tolerance failure unless both are None
            within = sval is None and pval is None
            diffs.append(FieldDiff(
                field_name=fname, sequential=sval, parallel=pval,
                abs_diff=None, within_tolerance=within,
            ))
            if not within:
                all_within = False
            continue
        abs_diff = abs(float(sval) - float(pval))
        within = abs_diff <= tolerance
        diffs.append(FieldDiff(
            field_name=fname, sequential=float(sval), parallel=float(pval),
            abs_diff=abs_diff, within_tolerance=within,
        ))
        if abs_diff > max_abs:
            max_abs = abs_diff
        if not within:
            all_within = False

    return ProjectDiff(
        name=name,
        status="PASS" if all_within else "FAIL",
        max_abs_diff=max_abs,
        field_diffs=diffs,
    )


def validate_parallel(
    workbook_path: Path,
    *,
    workers: int = 2,
    tolerance: float = 1e-4,
    timeout_sec: int = 3600,
    use_chunked: bool = True,
    allow_relaxed: bool = True,
) -> ValidationReport:
    """Run sequential baseline, then parallel; diff per-project results.

    Returns a ValidationReport. Caller decides what to do with FAILs.
    """
    if workers < 2:
        raise ValueError("validate_parallel requires workers >= 2")

    seq_batch = f"validate-seq-{uuid.uuid4().hex[:8]}"
    par_batch = f"validate-par-{uuid.uuid4().hex[:8]}"

    log.info("=" * 60)
    log.info("  Parallel correctness validation")
    log.info("  Workbook: %s", workbook_path.name)
    log.info("  Workers:  %d", workers)
    log.info("  Tolerance: %.0e", tolerance)
    log.info("=" * 60)

    log.info("\n[Pass 1] Sequential baseline (single worker)...")
    seq_record = solve_all(
        workbook_path,
        batch_id=seq_batch,
        timeout_sec=timeout_sec,
        use_chunked=use_chunked,
        allow_relaxed=allow_relaxed,
        save_solved=False,  # don't write _SOLVED.xlsm during validation
    )

    log.info("\n[Pass 2] Parallel run (%d workers)...", workers)
    par_record = solve_all(
        workbook_path,
        batch_id=par_batch,
        timeout_sec=timeout_sec,
        use_chunked=use_chunked,
        allow_relaxed=allow_relaxed,
        save_solved=False,
        workers=workers,
    )

    seq_by_name = _by_name(seq_record)
    par_by_name = _by_name(par_record)
    all_names = sorted(set(seq_by_name) | set(par_by_name))

    project_diffs = [
        _diff_one_project(n, seq_by_name.get(n), par_by_name.get(n), tolerance)
        for n in all_names
    ]
    failing = [p.name for p in project_diffs if p.status != "PASS"]

    speedup = (
        seq_record.total_duration_sec / par_record.total_duration_sec
        if par_record.total_duration_sec > 0 else 0.0
    )

    return ValidationReport(
        workbook_name=workbook_path.name,
        workers=workers,
        tolerance=tolerance,
        sequential_batch_id=seq_batch,
        parallel_batch_id=par_batch,
        sequential_duration=seq_record.total_duration_sec,
        parallel_duration=par_record.total_duration_sec,
        speedup=speedup,
        project_diffs=project_diffs,
        all_pass=not failing,
        failing_projects=failing,
    )


def format_report(report: ValidationReport) -> str:
    """Render a human-readable summary of the validation result."""
    lines: list[str] = []
    lines.append("=" * 80)
    lines.append(f"  PARALLEL CORRECTNESS VALIDATION — {report.workbook_name}")
    lines.append("=" * 80)
    lines.append(f"  Workers:           {report.workers}")
    lines.append(f"  Tolerance:         {report.tolerance:.0e}")
    lines.append(f"  Sequential time:   {report.sequential_duration:.1f}s "
                 f"(batch {report.sequential_batch_id})")
    lines.append(f"  Parallel time:     {report.parallel_duration:.1f}s "
                 f"(batch {report.parallel_batch_id})")
    lines.append(f"  Speedup:           {report.speedup:.2f}x")
    lines.append(f"  Projects compared: {len(report.project_diffs)}")
    lines.append(
        f"  Result:            {'PASS' if report.all_pass else 'FAIL'} "
        f"({len(report.failing_projects)} failing)"
    )
    lines.append("")
    lines.append(f"  {'Project':<32} {'Status':>8} {'Max |diff|':>14}")
    lines.append(f"  {'-'*32} {'-'*8} {'-'*14}")
    for pd in report.project_diffs:
        max_d = f"{pd.max_abs_diff:.3e}" if pd.max_abs_diff is not None else "—"
        lines.append(f"  {pd.name[:32]:<32} {pd.status:>8} {max_d:>14}")

    if report.failing_projects:
        lines.append("")
        lines.append("  --- FAILING PROJECT DETAILS ---")
        for pd in report.project_diffs:
            if pd.status == "PASS":
                continue
            lines.append(f"  {pd.name} [{pd.status}]:")
            for fd in pd.field_diffs:
                if fd.within_tolerance:
                    continue
                seq_s = f"{fd.sequential:.6f}" if fd.sequential is not None else "—"
                par_s = f"{fd.parallel:.6f}" if fd.parallel is not None else "—"
                diff_s = f"{fd.abs_diff:.3e}" if fd.abs_diff is not None else "—"
                lines.append(
                    f"    {fd.field_name:<18} seq={seq_s:>14}  "
                    f"par={par_s:>14}  diff={diff_s:>12}"
                )
    lines.append("=" * 80)
    return "\n".join(lines)
