"""Unit tests for the release-gate decision logic (no Excel).

These cover the rules the warm/cold/stress regression harness applies to
real solve results, so the gate logic itself is regression-protected even
though the Excel-driving tier (test_excel_regression.py) only runs opt-in.
"""
from __future__ import annotations

from dn38_solver.types import ProjectResult, RunRecord, SolveStatus
from dn38_solver.validation.release_gates import (
    evaluate_cold_gate,
    evaluate_stress_gate,
    evaluate_warm_gate,
    p95,
    percentile,
)


def _proj(name: str, tier: str, **kw) -> ProjectResult:
    return ProjectResult(
        name=name, col=8, col_letter="H",
        converged=(tier == "strict"), convergence_tier=tier, **kw,
    )


def _record(projects, *, status=SolveStatus.CONVERGED.value, total=100.0) -> RunRecord:
    return RunRecord(
        workbook_name="wb.xlsm", run_timestamp="2026-06-09T00:00:00",
        batch_id="test", solver_mode="hybrid_shadow",
        projects=tuple(projects), total_duration_sec=total, status=status,
    )


# --- percentile / p95 --------------------------------------------------------

def test_percentile_empty_is_none():
    assert percentile([], 95) is None
    assert p95([]) is None


def test_percentile_single_value():
    assert p95([42.0]) == 42.0


def test_percentile_interpolates():
    # P50 of 1..5 is 3; P95 lands near the top.
    assert percentile([1, 2, 3, 4, 5], 50) == 3.0
    assert percentile([1, 2, 3, 4, 5], 100) == 5.0
    assert 4.0 < p95([1, 2, 3, 4, 5]) <= 5.0


def test_percentile_ignores_none():
    assert percentile([None, 10.0, None, 20.0], 50) == 15.0


# --- cold gate: 100% strict --------------------------------------------------

def test_cold_gate_passes_all_strict():
    r = _record([_proj("A", "strict"), _proj("B", "strict")])
    res = evaluate_cold_gate(r)
    assert res.passed and res.gate == "cold"
    assert res.metrics["n_non_strict"] == 0.0


def test_cold_gate_fails_on_relaxed():
    # Relaxed is shippable for a bid but NOT acceptable for the cold gate.
    r = _record([_proj("A", "strict"), _proj("B", "relaxed")])
    res = evaluate_cold_gate(r)
    assert not res.passed
    assert "B" in res.detail


def test_cold_gate_ignores_skipped_and_not_attempted():
    r = _record([
        _proj("A", "strict"),
        _proj("PH", "skipped"),
        _proj("Crash", "not_attempted"),
    ])
    res = evaluate_cold_gate(r)
    assert res.passed
    assert res.metrics["n_projects"] == 1.0  # only the real one scored


def test_cold_gate_empty_record_fails():
    res = evaluate_cold_gate(_record([_proj("PH", "skipped")]))
    assert not res.passed


# --- warm gate: speed budget -------------------------------------------------

def test_warm_gate_passes_under_budget():
    r = _record([_proj("A", "strict")], total=120.0)
    res = evaluate_warm_gate(r, max_total_sec=600)
    assert res.passed
    assert res.metrics["total_sec"] == 120.0


def test_warm_gate_fails_over_total_budget():
    r = _record([_proj("A", "strict")], total=900.0)
    res = evaluate_warm_gate(r, max_total_sec=600)
    assert not res.passed
    assert "budget" in res.detail


def test_warm_gate_fails_on_non_converged_status():
    r = _record([_proj("A", "relaxed")], status=SolveStatus.NOT_CONVERGED.value, total=10.0)
    res = evaluate_warm_gate(r, max_total_sec=600)
    assert not res.passed
    assert "status" in res.detail


def test_warm_gate_p95_project_budget():
    r = _record([_proj("A", "strict")], total=100.0)
    # P95 of these is well over 150 -> fail when a per-project cap is set.
    res = evaluate_warm_gate(
        r, max_total_sec=600, max_p95_project_sec=150,
        project_durations=[100, 120, 130, 400],
    )
    assert not res.passed
    assert "P95" in res.detail
    assert "p95_project_sec" in res.metrics


# --- stress gate: N clean parallel runs --------------------------------------

def test_stress_gate_passes_three_clean_runs():
    runs = [_record([_proj("A", "strict")]) for _ in range(3)]
    res = evaluate_stress_gate(runs, parallel_diffs=[0, 0, 0])
    assert res.passed and res.gate == "stress"


def test_stress_gate_fails_too_few_runs():
    runs = [_record([_proj("A", "strict")]) for _ in range(2)]
    res = evaluate_stress_gate(runs, parallel_diffs=[0, 0])
    assert not res.passed
    assert "need 3" in res.detail


def test_stress_gate_fails_on_nonzero_diff():
    runs = [_record([_proj("A", "strict")]) for _ in range(3)]
    res = evaluate_stress_gate(runs, parallel_diffs=[0, 5, 0])
    assert not res.passed
    assert "diff" in res.detail


def test_stress_gate_fails_on_non_converged_run():
    runs = [
        _record([_proj("A", "strict")]),
        _record([_proj("A", "relaxed")], status=SolveStatus.ERROR.value),
        _record([_proj("A", "strict")]),
    ]
    res = evaluate_stress_gate(runs, parallel_diffs=[0, 0, 0])
    assert not res.passed
    assert "not converged" in res.detail
