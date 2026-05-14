"""Tests for dn38_solver.com.auto_recovery — close/reimport/reopen/retry flow.

The integration with Excel COM and the import_vba_module subprocess is
mocked out — these tests cover the orchestration logic only. End-to-end
recovery is exercised by the live solver runs against RP Puma fixtures.
"""
from __future__ import annotations

from pathlib import Path

import pytest

from dn38_solver.com.auto_recovery import (
    AutoRecoveryUnavailable,
    with_recovery,
)


class _CallLog:
    """Tiny ordered call recorder for asserting orchestration sequence."""

    def __init__(self) -> None:
        self.events: list[str] = []

    def mark(self, name: str) -> None:
        self.events.append(name)


def test_already_recovered_raises(monkeypatch, tmp_path):
    """Passing already_recovered=True must short-circuit with the
    AutoRecoveryUnavailable signal — auto-recovery is one-shot per
    workbook by design (no infinite loops on a workbook that's broken
    in ways re-import can't fix).
    """
    wb = tmp_path / "x.xlsm"
    wb.write_bytes(b"not-a-real-xlsm")  # never opened in this branch

    with pytest.raises(AutoRecoveryUnavailable):
        with_recovery(
            workbook_path=wb,
            close_open_handle=lambda: None,
            reopen_handle=lambda: None,
            retry_callable=lambda: None,
            already_recovered=True,
        )


def test_recovery_orchestration_order(monkeypatch, tmp_path):
    """Close → re-import → reopen → retry, in that order."""
    log = _CallLog()
    wb = tmp_path / "x.xlsm"
    wb.write_bytes(b"x")

    # Replace the subprocess call so the test doesn't need Excel.
    import dn38_solver.com.auto_recovery as ar
    monkeypatch.setattr(ar, "reimport_macro_subprocess",
                        lambda p, timeout_sec=120: log.mark("reimport"))

    ok = with_recovery(
        workbook_path=wb,
        close_open_handle=lambda: log.mark("close"),
        reopen_handle=lambda: log.mark("reopen"),
        retry_callable=lambda: log.mark("retry"),
        already_recovered=False,
    )
    assert ok is True
    assert log.events == ["close", "reimport", "reopen", "retry"]


def test_recovery_returns_false_when_retry_fails(monkeypatch, tmp_path):
    """A failing retry must be reported (return False) not raised. The
    caller is expected to surface the original COM error rather than
    pretend recovery worked.
    """
    wb = tmp_path / "x.xlsm"
    wb.write_bytes(b"x")

    import dn38_solver.com.auto_recovery as ar
    monkeypatch.setattr(ar, "reimport_macro_subprocess",
                        lambda p, timeout_sec=120: None)

    def _failing_retry() -> None:
        raise RuntimeError("simulated chunked failure")

    ok = with_recovery(
        workbook_path=wb,
        close_open_handle=lambda: None,
        reopen_handle=lambda: None,
        retry_callable=_failing_retry,
        already_recovered=False,
    )
    assert ok is False


def test_reimport_unavailable_bubbles_up(monkeypatch, tmp_path):
    """If the re-import subprocess can't run (missing script, non-zero
    exit), the unavailability surfaces as AutoRecoveryUnavailable rather
    than being silently swallowed — caller must fall back to the
    original COM error so the operator knows we didn't pretend.
    """
    wb = tmp_path / "x.xlsm"
    wb.write_bytes(b"x")

    import dn38_solver.com.auto_recovery as ar

    def _boom(p, timeout_sec=120):
        raise AutoRecoveryUnavailable("simulated missing script")

    monkeypatch.setattr(ar, "reimport_macro_subprocess", _boom)

    with pytest.raises(AutoRecoveryUnavailable):
        with_recovery(
            workbook_path=wb,
            close_open_handle=lambda: None,
            reopen_handle=lambda: None,
            retry_callable=lambda: None,
            already_recovered=False,
        )
