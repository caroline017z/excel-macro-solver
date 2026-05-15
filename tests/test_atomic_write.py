"""Tests for dn38_solver.com.status_aggregator.atomic_write_json.

Covers the retry-on-PermissionError fix introduced after the Queen City
2026-05-14 incident: the old one-shot `Path.replace` silently dropped
status writes when Windows AV / a Streamlit reader briefly held the
target open, even though both workers' macros had completed. A bounded
retry recovers the run in the common case while still surfacing
genuinely-stuck contention.
"""
from __future__ import annotations

import json
from pathlib import Path

import pytest

from dn38_solver.com import status_aggregator


def test_happy_path_writes_payload(tmp_path: Path) -> None:
    target = tmp_path / "status.json"
    payload = {"phase": "complete", "projects": [{"name": "P1"}]}

    status_aggregator.atomic_write_json(target, payload)

    assert json.loads(target.read_text(encoding="utf-8")) == payload
    assert not target.with_suffix(target.suffix + ".tmp").exists()


def test_retry_succeeds_after_transient_permission_error(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """Mimic the Queen City failure: first rename hits PermissionError,
    second succeeds. The helper must retry and the payload must land.
    """
    target = tmp_path / "status.json"
    payload = {"phase": "solving"}

    original_replace = Path.replace
    calls = {"n": 0}

    def flaky_replace(self: Path, target_path: Path) -> None:
        calls["n"] += 1
        if calls["n"] == 1:
            raise PermissionError("simulated AV lock on first attempt")
        return original_replace(self, target_path)

    monkeypatch.setattr(Path, "replace", flaky_replace)
    # Shorten retry delays so the test runs fast — the real schedule is
    # ~50ms..800ms; for unit-test purposes the timing doesn't matter, only
    # that retry happens.
    monkeypatch.setattr(
        status_aggregator,
        "_RENAME_RETRY_DELAYS_SEC",
        (0.0, 0.0, 0.0, 0.0, 0.0),
    )

    with caplog.at_level("INFO", logger="dn38_solver.com.status_aggregator"):
        status_aggregator.atomic_write_json(target, payload)

    assert calls["n"] == 2, "expected exactly one retry"
    assert json.loads(target.read_text(encoding="utf-8")) == payload
    assert any(
        "succeeded on attempt 2" in rec.getMessage()
        for rec in caplog.records
    ), "retry path should log a single info-level success line"


def test_exhausted_retries_leaves_tmp_for_recovery(
    tmp_path: Path,
    monkeypatch: pytest.MonkeyPatch,
    caplog: pytest.LogCaptureFixture,
) -> None:
    """When every replace attempt fails, the .tmp file must survive so
    the next successful write recovers state and operators have forensics.
    A warning (not silent failure) must be logged.
    """
    target = tmp_path / "status.json"
    tmp = target.with_suffix(target.suffix + ".tmp")
    payload = {"phase": "complete"}

    def always_fail(self: Path, target_path: Path) -> None:
        raise PermissionError("simulated persistent lock")

    monkeypatch.setattr(Path, "replace", always_fail)
    monkeypatch.setattr(
        status_aggregator,
        "_RENAME_RETRY_DELAYS_SEC",
        (0.0, 0.0),
    )

    with caplog.at_level("WARNING", logger="dn38_solver.com.status_aggregator"):
        status_aggregator.atomic_write_json(target, payload)

    assert not target.exists(), "target should not have been created"
    assert tmp.exists(), ".tmp must survive on exhaustion for next-write recovery"
    assert json.loads(tmp.read_text(encoding="utf-8")) == payload
    assert any(
        "failed after" in rec.getMessage() and rec.levelname == "WARNING"
        for rec in caplog.records
    ), "exhaustion path must surface a warning, not fail silently"
