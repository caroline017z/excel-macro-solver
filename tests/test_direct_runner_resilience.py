"""Structural regression guards for direct_runner failure-isolation fixes.

These paths need a live Excel COM session to exercise end-to-end, so —
following the repo's established pattern for the COM boundary
(see test_watchdog_and_status.test_direct_runner_threads_dscr_into_stamp_call)
— we assert the source structure that makes the fix work. They fail loud
if a refactor strips the isolation, which is the regression that matters.
"""
from __future__ import annotations

from pathlib import Path

_SRC = (
    Path(__file__).resolve().parent.parent
    / "dn38_solver" / "com" / "direct_runner.py"
).read_text(encoding="utf-8")


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
