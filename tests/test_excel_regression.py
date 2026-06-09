"""Warm / cold / stress regression tier — opt-in, drives real Excel solves.

Skipped entirely unless DN38_EXCEL_TESTS=1, and each test additionally skips
if its fixture path is not provided / not present. This tier never runs in
the cross-platform CI matrix (no Excel there); it is the release gate run on
a Windows+Excel box before shipping a new SolveHeadless.bas, via
run_release_gate.ps1. The pass/fail rules live in
dn38_solver.validation.release_gates (unit-tested in test_release_gates.py).

Environment contract
--------------------
  DN38_EXCEL_TESTS=1          master switch (required)
  DN38_WARM_FIXTURE=<path>    pre-solved workbook (speed gate)
  DN38_COLD_FIXTURE=<path>    unsolved workbook (100% strict convergence)
  DN38_STRESS_FIXTURE=<path>  multi-project workbook (parallel resilience)
  DN38_WARM_MAX_SEC=600       warm total wall-time budget (optional)
  DN38_WARM_P95_SEC=<sec>     warm P95-per-project budget (optional)
  DN38_STRESS_WORKERS=2       parallel workers for the stress run (optional)
  DN38_STRESS_RUNS=3          consecutive clean runs required (optional)
  DN38_STRESS_NPP_TOL=0.005   max cross-run NPP drift ($/W) treated as clean
"""
from __future__ import annotations

import os
from pathlib import Path

import pytest

from dn38_solver.solver.orchestrator import solve_all
from dn38_solver.types import SolveStatus
from dn38_solver.validation.release_gates import (
    evaluate_cold_gate,
    evaluate_stress_gate,
    evaluate_warm_gate,
)

pytestmark = pytest.mark.excel_integration


def _require_excel_tier() -> None:
    if os.environ.get("DN38_EXCEL_TESTS") != "1":
        pytest.skip("Excel integration tier off (set DN38_EXCEL_TESTS=1 to run)")


def _fixture(env_var: str) -> Path:
    _require_excel_tier()
    raw = os.environ.get(env_var)
    if not raw:
        pytest.skip(f"{env_var} not set")
    p = Path(raw)
    if not p.exists():
        pytest.skip(f"{env_var}={raw} does not exist")
    return p


def _env_float(name: str, default: float) -> float:
    raw = os.environ.get(name)
    return float(raw) if raw else default


def _env_int(name: str, default: int) -> int:
    raw = os.environ.get(name)
    return int(raw) if raw else default


def test_warm_speed_gate() -> None:
    """A pre-solved workbook must re-converge within the wall-time budget."""
    wb = _fixture("DN38_WARM_FIXTURE")
    record = solve_all(wb, auto_fix=True, use_chunked=True, allow_relaxed=True)
    res = evaluate_warm_gate(
        record,
        max_total_sec=_env_float("DN38_WARM_MAX_SEC", 600.0),
        max_p95_project_sec=(
            _env_float("DN38_WARM_P95_SEC", 0.0) or None
        ),
    )
    assert res.passed, f"warm gate FAILED: {res.detail} | metrics={res.metrics}"


def test_cold_strict_convergence_gate() -> None:
    """An unsolved workbook must reach 100% STRICT convergence."""
    wb = _fixture("DN38_COLD_FIXTURE")
    # No allow_relaxed: the cold gate demands strict, and the evaluator
    # checks per-project tiers regardless, but keeping the run strict avoids
    # a misleading run-level CONVERGED built from relaxed projects.
    record = solve_all(wb, auto_fix=True, use_chunked=True, allow_relaxed=False)
    res = evaluate_cold_gate(record)
    assert res.passed, f"cold gate FAILED: {res.detail}"


def test_stress_parallel_resilience_gate() -> None:
    """A multi-project portfolio solved under parallel workers must produce
    N consecutive converged runs with stable per-project NPP across runs.
    """
    wb = _fixture("DN38_STRESS_FIXTURE")
    workers = _env_int("DN38_STRESS_WORKERS", 2)
    n_runs = _env_int("DN38_STRESS_RUNS", 3)
    npp_tol = _env_float("DN38_STRESS_NPP_TOL", 0.005)

    records = []
    npp_by_run: list[dict[str, float]] = []
    for _ in range(n_runs):
        rec = solve_all(
            wb, auto_fix=True, use_chunked=True, allow_relaxed=True,
            workers=workers,
        )
        records.append(rec)
        npp_by_run.append({
            p.name: p.npp_per_w
            for p in rec.projects
            if p.npp_per_w is not None
        })

    # Cross-run determinism: a run is "dirty" if any project's NPP drifted
    # more than tol vs the first run. Stands in for validate_parallel's
    # sequential-vs-parallel diff with a cheaper run-to-run stability check.
    diffs: list[int] = []
    base = npp_by_run[0] if npp_by_run else {}
    for run_npps in npp_by_run:
        dirty = sum(
            1 for name, npp in run_npps.items()
            if name in base and abs(npp - base[name]) > npp_tol
        )
        diffs.append(dirty)

    res = evaluate_stress_gate(records, parallel_diffs=diffs, min_clean_runs=n_runs)
    assert res.passed, (
        f"stress gate FAILED: {res.detail} | per-run NPP-drift counts={diffs}"
    )
    # Belt-and-braces: every run must have actually converged.
    assert all(r.status == SolveStatus.CONVERGED.value for r in records)
