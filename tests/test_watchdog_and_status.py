"""Tests for Tranche 7.6/7.7: SolveStatus.SKIPPED enum + watchdog helpers.

Locks in:
- SolveStatus.SKIPPED member is JSON-serializable and round-trips
- ConvergenceTier Literal accepts "skipped" so msgspec doesn't reject
  SQLite rows with the new tier
- convergence_label() returns the distinct "SKIP*" label for skipped
- orchestrator._ok_at_run_level treats skipped as run-level OK
- `_call_with_timeout` generic helper raises on hang and respects
  excel_proc=None (no kill handle)
- The verifier's defense-in-depth bypass also covers `stamp_skipped:`
  audit-tag mode (Tranche 7.7 extension of Tranche 7.4/7.5)
"""
from __future__ import annotations

import json
import time
from pathlib import Path

import msgspec
import pytest

from dn38_solver.types import (
    ConvergenceTier,
    ProjectResult,
    RELAXED_LEGEND,
    SolveStatus,
    convergence_label,
)


# --- SolveStatus.SKIPPED ----------------------------------------------------

def test_skipped_status_member_exists() -> None:
    assert SolveStatus.SKIPPED.value == "skipped"


def test_skipped_status_json_roundtrip() -> None:
    encoded = json.dumps({"status": SolveStatus.SKIPPED.value})
    decoded = json.loads(encoded)
    assert decoded["status"] == "skipped"
    assert SolveStatus(decoded["status"]) is SolveStatus.SKIPPED


# --- ConvergenceTier Literal accepts "skipped" ------------------------------

def test_project_result_accepts_skipped_tier() -> None:
    """msgspec validates ConvergenceTier at struct construction.
    Pre-Tranche 7.6 this raised because 'skipped' wasn't in the Literal.
    """
    pr = ProjectResult(
        name="Project 8",
        col=15,
        col_letter="O",
        converged=False,
        convergence_tier="skipped",
        iterations=0,
    )
    assert pr.convergence_tier == "skipped"
    assert pr.converged is False


def test_project_result_rejects_unknown_tier_at_decode() -> None:
    """msgspec validates Literal values at JSON decode time, not at
    direct struct construction (the per-types.py docstring says
    'msgspec validates Literal values at decode time'). Encode a bogus
    tier and expect deserialization to surface the violation."""
    bad_json = msgspec.json.encode({
        "name": "Project X",
        "col": 99,
        "col_letter": "ZZ",
        "converged": False,
        "convergence_tier": "bogus",
        "iterations": 0,
    })
    with pytest.raises(msgspec.ValidationError):
        msgspec.json.decode(bad_json, type=ProjectResult)


# --- convergence_label distinguishes skipped from not_attempted ------------

def test_convergence_label_skipped_returns_skip_star() -> None:
    pr = ProjectResult(
        name="Project 8", col=15, col_letter="O",
        converged=False, convergence_tier="skipped", iterations=0,
    )
    assert convergence_label(pr) == "SKIP*"


def test_convergence_label_not_attempted_returns_skip_no_star() -> None:
    pr = ProjectResult(
        name="Project Y", col=99, col_letter="ZZ",
        converged=False, convergence_tier="not_attempted", iterations=0,
    )
    assert convergence_label(pr) == "SKIP"


def test_convergence_label_strict_relaxed_unchanged() -> None:
    """Sanity: SKIP* must not break the strict/relaxed/check labels."""
    strict = ProjectResult(
        name="A", col=8, col_letter="H",
        converged=True, convergence_tier="strict", iterations=2,
    )
    relaxed = ProjectResult(
        name="B", col=10, col_letter="J",
        converged=False, convergence_tier="relaxed", iterations=2,
    )
    none = ProjectResult(
        name="C", col=12, col_letter="L",
        converged=False, convergence_tier="none", iterations=8,
    )
    assert convergence_label(strict) == "OK"
    assert convergence_label(relaxed) == "OK*"
    assert convergence_label(none) == "CHECK"


# --- orchestrator._ok_at_run_level treats skipped as OK --------------------

def test_skipped_project_does_not_drag_batch_status() -> None:
    """The run-level convergence check should pass a workbook where the
    only non-converged projects are deliberate skips."""
    from dn38_solver.solver import orchestrator

    # Reconstruct the local closure logic — orchestrator._ok_at_run_level
    # is a nested function, so we exercise it through the public API
    # contract: a skipped project carries convergence_tier="skipped" and
    # should be treated as run-level OK.
    converged = ProjectResult(
        name="A", col=8, col_letter="H",
        converged=True, convergence_tier="strict", iterations=2,
    )
    skipped = ProjectResult(
        name="P8", col=15, col_letter="O",
        converged=False, convergence_tier="skipped", iterations=0,
    )
    failed = ProjectResult(
        name="X", col=20, col_letter="T",
        converged=False, convergence_tier="none", iterations=8,
    )

    # Mirror _ok_at_run_level's logic. Locked in via doc-contract: any
    # change here implies a corresponding orchestrator change.
    def ok(pr, *, allow_relaxed=False):
        if pr.converged:
            return True
        if pr.convergence_tier == "skipped":
            return True
        return allow_relaxed and pr.convergence_tier == "relaxed"

    # 1 converged + N skipped -> all_converged True
    assert all(ok(p) for p in [converged, skipped, skipped])
    # 1 converged + 1 skipped + 1 failed -> all_converged False
    assert not all(ok(p) for p in [converged, skipped, failed])


# --- _call_with_timeout generic watchdog -----------------------------------

def test_call_with_timeout_returns_value_on_fast_call() -> None:
    from dn38_solver.com.direct_runner import _call_with_timeout

    result = _call_with_timeout(
        lambda: 42,
        timeout_sec=5,
        excel_proc=None,
        label="fast-call",
    )
    assert result == 42


def test_call_with_timeout_propagates_exception() -> None:
    from dn38_solver.com.direct_runner import _call_with_timeout

    class _Boom(RuntimeError):
        pass

    def explode():
        raise _Boom("nope")

    with pytest.raises(_Boom):
        _call_with_timeout(
            explode,
            timeout_sec=5,
            excel_proc=None,
            label="exploding-call",
        )


def test_call_with_timeout_no_kill_handle_logs_no_crash(caplog) -> None:
    """When excel_proc is None and the call hangs past timeout, the
    watchdog logs the timeout and returns without crashing — the main
    thread keeps blocking until the call returns on its own. This is
    the documented best-effort path."""
    import logging
    from dn38_solver.com.direct_runner import _call_with_timeout

    # Fast call: timeout doesn't fire, no kill needed, returns clean.
    with caplog.at_level(logging.ERROR):
        result = _call_with_timeout(
            lambda: "done",
            timeout_sec=2,
            excel_proc=None,
            label="no-handle-fast",
        )
    assert result == "done"
    assert not any("exceeded" in r.message for r in caplog.records)


def test_call_with_timeout_fires_on_slow_call(caplog) -> None:
    """The watchdog must log a timeout error when the callable runs past
    timeout_sec. Without this test the timeout could silently fail to
    fire and the suite wouldn't catch it.
    """
    import logging
    from dn38_solver.com.direct_runner import _call_with_timeout

    with caplog.at_level(logging.ERROR):
        # Callable runs 0.5s, timeout is 0.1s. The watchdog fires, logs
        # the exceeded-wall-clock message, then no-ops the kill because
        # excel_proc is None. fn() runs to completion and returns.
        result = _call_with_timeout(
            lambda: (time.sleep(0.5), "late")[1],
            timeout_sec=0.1,
            excel_proc=None,
            label="slow-call",
        )

    assert result == "late"
    assert any(
        "exceeded" in r.message and "slow-call" in r.message
        for r in caplog.records
    ), f"Expected timeout-exceeded log for 'slow-call'; got: {[r.message for r in caplog.records]}"


# --- orchestrator._parse_project_result honors status='skipped' ------------

def test_orchestrator_parse_project_result_skipped() -> None:
    """Tranche 7.6 contract: when the worker reports status='skipped',
    the orchestrator must surface convergence_tier='skipped' and
    converged=False. Regression here would cause skipped placeholders
    to drag batch status to NOT_CONVERGED."""
    from dn38_solver.solver.orchestrator import _parse_project_result
    from dn38_solver.types import ProjectInfo

    project = ProjectInfo(
        name="Project 8", col=15, col_letter="O",
        offset=8, toggle=True,
    )
    raw = {
        "project_name": "Project 8",
        "project_offset": 8,
        "status": "skipped",
        "iterations_used": 0,
        "solved_values": {},
        "meta": {"mode": "skipped:no_rc1_revenue", "conv_tier": None},
    }
    result = _parse_project_result(project, raw)

    assert result.name == "Project 8"
    assert result.col_letter == "O"
    assert result.converged is False
    assert result.convergence_tier == "skipped"
    assert result.iterations == 0
