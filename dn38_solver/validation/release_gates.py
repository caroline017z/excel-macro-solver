"""dn38_solver.validation.release_gates — pure gate-decision logic.

The warm/cold/stress regression harness (tests/test_excel_regression.py,
run_release_gate.ps1) drives real Excel solves; this module holds the
*decisions* it makes about the results, as pure functions over RunRecord
so they unit-test without Excel. Keeping the thresholds and pass/fail rules
here — not inline in the integration tests — is what lets the riskiest
60% of the system (chunked loop, watchdog, recovery, parallel merge) be
gated by a rule set that is itself covered by fast tests.

Gates (per ENGINEERING_REVIEW.md Sprint 1):
  * warm  — a pre-solved workbook must re-converge fast (wall-time budget,
            optional P95-per-project budget). Guards speed regressions.
  * cold  — an unsolved workbook must reach 100% STRICT convergence.
            Guards correctness regressions (relaxed-only is not enough for
            the release gate, even though it may ship a bid).
  * stress— a multi-project portfolio solved under parallel workers must
            produce N consecutive clean runs with zero cross-run cell diff.
            Guards RPC resilience and merge correctness.
"""
from __future__ import annotations

from typing import NamedTuple, Sequence

from dn38_solver.types import RunRecord, SolveStatus


class GateResult(NamedTuple):
    gate: str
    passed: bool
    detail: str
    metrics: dict[str, float]


def percentile(values: Sequence[float], pct: float) -> float | None:
    """Linear-interpolation percentile. Returns None on empty input.

    pct is 0..100. P95 of one value is that value; of two is interpolated.
    """
    xs = sorted(v for v in values if v is not None)
    if not xs:
        return None
    if len(xs) == 1:
        return float(xs[0])
    k = (len(xs) - 1) * (pct / 100.0)
    lo = int(k)
    hi = min(lo + 1, len(xs) - 1)
    frac = k - lo
    return float(xs[lo] + (xs[hi] - xs[lo]) * frac)


def p95(values: Sequence[float]) -> float | None:
    return percentile(values, 95.0)


def _scored_projects(record: RunRecord) -> list:
    """Projects that count toward a gate — deliberate skips and worker-crash
    not_attempted rows are excluded (they're not convergence outcomes).
    """
    return [
        p for p in record.projects
        if p.convergence_tier not in ("skipped", "not_attempted")
    ]


def evaluate_cold_gate(record: RunRecord) -> GateResult:
    """Cold fixture: every scored project must be STRICT-converged."""
    scored = _scored_projects(record)
    if not scored:
        return GateResult(
            "cold", False, "no scored projects in record", {"n_projects": 0.0}
        )
    non_strict = [p.name for p in scored if p.convergence_tier != "strict"]
    passed = not non_strict
    detail = (
        f"all {len(scored)} project(s) strict"
        if passed
        else f"{len(non_strict)}/{len(scored)} not strict: "
             + ", ".join(non_strict[:8])
    )
    return GateResult(
        "cold", passed, detail,
        {"n_projects": float(len(scored)), "n_non_strict": float(len(non_strict))},
    )


def evaluate_warm_gate(
    record: RunRecord,
    *,
    max_total_sec: float,
    max_p95_project_sec: float | None = None,
    project_durations: Sequence[float] | None = None,
) -> GateResult:
    """Warm fixture (pre-solved): must converge within the time budget.

    Always gates on run status == converged and total wall time. When
    per-project durations are supplied, also gates the P95-per-project time.
    """
    reasons: list[str] = []
    if record.status != SolveStatus.CONVERGED.value:
        reasons.append(f"status={record.status}")
    if record.total_duration_sec > max_total_sec:
        reasons.append(
            f"total {record.total_duration_sec:.0f}s > {max_total_sec:.0f}s budget"
        )
    metrics: dict[str, float] = {"total_sec": record.total_duration_sec}
    if project_durations:
        p95_val = p95(project_durations)
        if p95_val is not None:
            metrics["p95_project_sec"] = p95_val
            if max_p95_project_sec is not None and p95_val > max_p95_project_sec:
                reasons.append(
                    f"P95/project {p95_val:.0f}s > {max_p95_project_sec:.0f}s"
                )
    passed = not reasons
    return GateResult(
        "warm", passed,
        "within speed budget" if passed else "; ".join(reasons),
        metrics,
    )


def evaluate_stress_gate(
    records: Sequence[RunRecord],
    *,
    parallel_diffs: Sequence[int],
    min_clean_runs: int = 3,
) -> GateResult:
    """Stress fixture: N consecutive clean parallel runs, zero cross-run diff.

    records: one RunRecord per repeated run. parallel_diffs: the mismatched-
    cell count from validate_parallel for each run (0 = clean). Passes iff at
    least `min_clean_runs` runs, all converged, all diffs zero.
    """
    reasons: list[str] = []
    if len(records) < min_clean_runs:
        reasons.append(f"only {len(records)} run(s), need {min_clean_runs}")
    non_conv = [
        i for i, r in enumerate(records)
        if r.status != SolveStatus.CONVERGED.value
    ]
    if non_conv:
        reasons.append(f"runs not converged: {non_conv}")
    dirty = [i for i, d in enumerate(parallel_diffs) if d != 0]
    if dirty:
        reasons.append(f"runs with nonzero parallel diff: {dirty}")
    passed = not reasons
    return GateResult(
        "stress", passed,
        f"{len(records)} clean run(s), zero diff" if passed else "; ".join(reasons),
        {"n_runs": float(len(records)), "n_dirty": float(len(dirty))},
    )
