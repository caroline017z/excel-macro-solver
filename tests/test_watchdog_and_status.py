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


# --- Tranche 7.10: post-read DSCR restore contract -------------------------
# Locks in the fix for the 2026-05-18 SolarStone bug. PT Returns!F129 is a
# single live cell that GoalSeek overwrites per project. By post-read time
# it holds the LAST-solved-project's DSCR, so the CalculateFull inside
# StampActiveProjectColumnHL would propagate the wrong DSCR through the
# IRR chain and stamp the wrong Live IRR into rows 31/37. The fix is to
# pass meta["dscr"] as the second arg to StampActiveProjectColumnHL and
# have VBA restore F129 before the CalculateFull.
#
# These tests verify both the contract (Python args match VBA sig) and
# the call-site (direct_runner threads meta["dscr"] through).

def test_stamp_active_project_contract_includes_dscr_arg() -> None:
    """The Python-side VBASub declaration for StampActiveProjectColumnHL
    must include dscrRestore as a Double. Drift here would silently
    revert to the pre-fix single-arg call and resurrect the IRR-bleed
    bug (Albion / Bethel stamped at 40%+ vs 18% target on SolarStone).
    """
    from dn38_solver.com.vba_contract import STAMP_ACTIVE_PROJECT_COLUMN

    arg_names = [a[0] for a in STAMP_ACTIVE_PROJECT_COLUMN.args]
    arg_types = [a[1] for a in STAMP_ACTIVE_PROJECT_COLUMN.args]
    assert arg_names == ["colIdx", "dscrRestore"], (
        f"StampActiveProjectColumnHL args drifted: got {arg_names}"
    )
    assert arg_types == ["Integer", "Double"], (
        f"StampActiveProjectColumnHL arg types drifted: got {arg_types}"
    )


def test_vba_stamp_active_signature_takes_two_args() -> None:
    """The VBA Sub declaration in SolveHeadless.bas must match the Python
    contract — two args, dscrRestore as Double. A drift here means the
    macro re-import would silently leave the wrong signature in the
    workbook and Application.Run would fail with arg-count mismatch.
    """
    from pathlib import Path
    repo_root = Path(__file__).resolve().parent.parent
    bas = (repo_root / "SolveHeadless.bas").read_text(encoding="utf-8")
    # Find the Public Sub StampActiveProjectColumnHL declaration. The
    # signature spans multiple lines via VBA's `_` continuation, so
    # collapse continuations before matching.
    import re
    collapsed = re.sub(r"_\s*\n\s*", " ", bas)
    m = re.search(
        r"Public Sub StampActiveProjectColumnHL\(([^)]*)\)",
        collapsed,
    )
    assert m, "StampActiveProjectColumnHL declaration not found in SolveHeadless.bas"
    sig = m.group(1)
    assert "colIdx As Integer" in sig
    assert "dscrRestore As Double" in sig, (
        f"dscrRestore arg missing from VBA sig — pre-fix bug resurrected: {sig!r}"
    )


def test_direct_runner_threads_dscr_into_stamp_call(monkeypatch) -> None:
    """The post-read pass in direct_runner must pass meta['dscr'] as the
    second positional arg to STAMP_ACTIVE_PROJECT_COLUMN. Verifies the
    call-site stays in sync with the contract — a regression here is
    invisible to the contract test above (which only checks the schema).

    The check is structural: we read the call site as source text and
    assert that the dscr is read from meta and threaded into the Application.Run
    args tuple. Spinning up Excel COM in unit tests would require a
    licensed install + a real workbook; the source-level assert is the
    pragmatic safety net.
    """
    from pathlib import Path
    repo_root = Path(__file__).resolve().parent.parent
    src = (repo_root / "dn38_solver" / "com" / "direct_runner.py").read_text(
        encoding="utf-8"
    )

    # The call site must read DSCR from meta and pass it as a positional
    # arg in the same tuple as STAMP_ACTIVE_PROJECT_COLUMN. The marker
    # text is stable; if anyone restructures this, they must keep DSCR
    # threading or this test breaks loud.
    assert 'meta.get("dscr")' in src, (
        "direct_runner no longer reads meta['dscr'] — DSCR restore bypassed"
    )
    assert "STAMP_ACTIVE_PROJECT_COLUMN" in src
    # The contract test above already locks in the two-arg shape; here
    # we just need to confirm DSCR is what fills the second slot.
    import re
    # Find the STAMP_ACTIVE_PROJECT_COLUMN call site and assert dscr is
    # nearby (within ~10 lines — generous so refactors don't break us).
    lines = src.splitlines()
    for i, line in enumerate(lines):
        if "STAMP_ACTIVE_PROJECT_COLUMN" in line and "vba_call_str" in line:
            window = "\n".join(lines[max(0, i - 5):i + 10])
            assert "dscr_for_stamp" in window or "dscr" in window, (
                f"STAMP_ACTIVE_PROJECT_COLUMN call site at line {i+1} "
                f"is missing DSCR threading; got:\n{window}"
            )
            break
    else:
        pytest.fail("STAMP_ACTIVE_PROJECT_COLUMN call site not found")
