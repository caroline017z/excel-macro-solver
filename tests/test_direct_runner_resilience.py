"""Structural regression guards for direct_runner failure-isolation fixes.

These paths need a live Excel COM session to exercise end-to-end, so —
following the repo's established pattern for the COM boundary
(see test_watchdog_and_status.test_direct_runner_threads_dscr_into_stamp_call)
— we assert the source structure that makes the fix work. They fail loud
if a refactor strips the isolation, which is the regression that matters.
"""
from __future__ import annotations

import re
import time
from pathlib import Path

import pytest

_REPO = Path(__file__).resolve().parent.parent
_SRC = (_REPO / "dn38_solver" / "com" / "direct_runner.py").read_text(encoding="utf-8")
_ORCH = (_REPO / "dn38_solver" / "solver" / "orchestrator.py").read_text(encoding="utf-8")


def test_stamp_call_is_isolated_per_project() -> None:
    """C2: the per-project StampActiveProjectColumnHL call must be wrapped
    in its own try/except that sets stamp_failed, NOT left to propagate to
    the function-level catch-all (which discards the whole run's results
    and skips the SaveAs — the SolarStone 2026-06-04 incident).
    """
    lines = _SRC.splitlines()
    idx = next(
        (
            i for i, ln in enumerate(lines)
            if "STAMP_ACTIVE_PROJECT_COLUMN" in ln and "vba_call_str" in ln
        ),
        None,
    )
    assert idx is not None, "STAMP_ACTIVE_PROJECT_COLUMN call site not found"
    window = "\n".join(lines[max(0, idx - 8):idx + 32])
    assert "try:" in window, (
        "STAMP_ACTIVE_PROJECT_COLUMN call is not guarded by try/ — a stamp "
        "failure would propagate to the catch-all and discard the run"
    )
    assert "stamp_failed = True" in window, (
        "stamp_failed flag not set on stamp failure — isolation removed"
    )


def test_stamp_failed_status_forces_non_converged_tier() -> None:
    """C2: a stamp_failed project must surface a distinct status AND force
    conv_tier to 'none' — convergence_label() renders 'OK' for any 'strict'
    tier regardless of the converged flag, so a leftover tier would show a
    failed project as converged.
    """
    assert '"stamp_failed"' in _SRC or "'stamp_failed'" in _SRC, (
        "stamp_failed status not surfaced"
    )
    assert 'meta["conv_tier"] = "none"' in _SRC, (
        "stamp_failed path does not force conv_tier='none' — failed project "
        "could render as OK in the convergence table"
    )


def test_recovery_close_saves_before_discarding() -> None:
    """C1: the recovery-path workbook close must Save the temp copy first,
    so pre-failure projects' converged stamps survive into the reopened
    file instead of being discarded and replaced with pre-solve values.
    """
    lines = _SRC.splitlines()
    idx = next(
        (i for i, ln in enumerate(lines) if "def _close_wb()" in ln), None
    )
    assert idx is not None, "_close_wb (recovery close) not found"
    body = "\n".join(lines[idx:idx + 40])
    assert "wb.Save" in body, (
        "_close_wb no longer saves before close — pre-failure converged "
        "stamps would be discarded (C1 silent-wrong-NPP regression)"
    )
    # The Save must precede the Close in the function body.
    assert body.index("wb.Save") < body.index("wb.Close"), (
        "Save must run BEFORE Close in _close_wb"
    )


def test_recovered_run_retains_checkpoints() -> None:
    """C1: a run completed via auto-recovery must NOT clear its per-project
    checkpoints — they are the forensic record of what converged before the
    crash. run_direct surfaces a 'recovered' flag; the orchestrator gates
    the checkpoint clear on it.
    """
    assert '"recovered": run_recovered' in _SRC, (
        "run_direct result dict no longer reports the 'recovered' flag"
    )
    assert 'batch_result.get("recovered")' in _ORCH, (
        "orchestrator no longer reads the 'recovered' flag"
    )
    assert "not run_was_recovered" in _ORCH, (
        "orchestrator clears checkpoints unconditionally on CONVERGED — "
        "recovered-run checkpoints would be lost"
    )


# --- Timeout family (C3, C8, C9, C10) ---------------------------------------

def test_watchdog_raises_real_timeouterror_when_it_fires() -> None:
    """C9: when the watchdog fires, the wrapped call surfaces a generic COM
    error from the killed Excel; _with_watchdog must re-raise it as a
    TimeoutError so the callers' `except TimeoutError` branches (which build
    labeled TIMEOUT diagnostics) are actually reachable. No Excel needed —
    excel_proc=None exercises the no-handle path while still setting the
    timed_out flag.
    """
    from dn38_solver.com.direct_runner import _with_watchdog

    def slow_then_com_error():
        time.sleep(0.25)
        raise RuntimeError("0x800706BA RPC server unavailable")

    with pytest.raises(TimeoutError):
        _with_watchdog(
            slow_then_com_error, timeout_sec=0.05,
            excel_proc=None, label="unit",
        )


def test_watchdog_passes_fast_success_through() -> None:
    """C10: a call that completes well within the budget returns normally
    and is never misclassified as a timeout.
    """
    from dn38_solver.com.direct_runner import _with_watchdog
    out = _with_watchdog(lambda: 42, timeout_sec=5, excel_proc=None, label="unit")
    assert out == 42


def test_watchdog_non_timeout_exception_propagates_unchanged() -> None:
    """A genuine error before the deadline must propagate as itself, not be
    masked as a TimeoutError.
    """
    from dn38_solver.com.direct_runner import _with_watchdog

    def boom():
        raise ValueError("genuine failure")

    with pytest.raises(ValueError):
        _with_watchdog(boom, timeout_sec=5, excel_proc=None, label="unit")


def test_per_call_timeout_exceeds_vba_project_budget() -> None:
    """C8: the Python per-call watchdog must sit ABOVE the VBA-side
    PROJECT_TIMEOUT_SECONDS so a slow cold project hits VBA's graceful
    per-project timeout (which still writes a telemetry row) rather than
    being hard-killed, which loses the whole session's read pass.
    """
    from dn38_solver.com.direct_runner import DEFAULT_PER_CALL_TIMEOUT_SEC
    bas = (_REPO / "SolveHeadless.bas").read_text(encoding="latin-1")
    m = re.search(r"PROJECT_TIMEOUT_SECONDS\s+As Double\s*=\s*(\d+)", bas)
    assert m, "PROJECT_TIMEOUT_SECONDS not found in SolveHeadless.bas"
    vba_cap = int(m.group(1))
    assert DEFAULT_PER_CALL_TIMEOUT_SEC > vba_cap, (
        f"per-call watchdog ({DEFAULT_PER_CALL_TIMEOUT_SEC}s) must exceed the "
        f"VBA project budget ({vba_cap}s) so VBA's graceful timeout wins"
    )


def test_orchestrator_auto_enables_chunked_for_multiproject() -> None:
    """C3: a multi-project run must auto-upgrade to the chunked path — the
    single-shot whole-portfolio call would be hard-killed by the per-project
    watchdog on any portfolio larger than a few projects.
    """
    assert "len(projects) > 1" in _ORCH, (
        "orchestrator no longer auto-enables chunked for multi-project runs"
    )
    idx = _ORCH.index("len(projects) > 1")
    window = _ORCH[idx:idx + 400]
    assert "use_chunked = True" in window, (
        "multi-project branch does not set use_chunked = True"
    )
