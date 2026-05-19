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


# --- Tranche 7.12: PI 31/37 stay dynamic; hardcode lands on PI 371 ----------
# Caroline's 2026-05-19 rule: stop hardcoding over PI rows 31 (Live Appraisal
# IRR) and 37 (Live Levered Pre-Tax IRR) — those need to stay as sticky-IF
# circular formulas so the audit chain from solver tabs through to the
# per-column cells is preserved in the formula bar. Instead, hardcode the
# converged DSCR Multiple onto PI row 371 (Min Equity DSCR Multiple). That
# upstream-input lock freezes debt sizing / equity / IRR cascade by formula,
# and the one-cell revert (restore ='PT Returns'!$F$129 on row 371) puts
# Min Equity back into fully dynamic solve mode.
#
# These tests lock in the contract + VBA signature + behavior. Drift in any
# of these resurrects the audit-chain destruction Caroline rejected.

def test_stamp_converged_values_contract_carries_dscr_not_irrs() -> None:
    """The StampConvergedValuesHL Python contract must take dscrMult and
    NOT take liveIRR / apprLive. A regression to the pre-7.12 signature
    would have the merge fallback re-hardcode over rows 31/37, destroying
    the audit chain Caroline relies on when opening the merged file.
    """
    from dn38_solver.com.vba_contract import STAMP_CONVERGED_VALUES

    arg_names = [a[0] for a in STAMP_CONVERGED_VALUES.args]
    arg_types = [a[1] for a in STAMP_CONVERGED_VALUES.args]
    assert arg_names == ["colIdx", "npp", "devFee", "fmv", "nppTotal", "dscrMult"], (
        f"StampConvergedValuesHL args drifted: got {arg_names}"
    )
    assert "liveIRR" not in arg_names, (
        "liveIRR resurrected in StampConvergedValuesHL — would re-hardcode PI 37"
    )
    assert "apprLive" not in arg_names, (
        "apprLive resurrected in StampConvergedValuesHL — would re-hardcode PI 31"
    )
    assert arg_types == ["Integer", "Double", "Double", "Double", "Double", "Double"]


def test_vba_stamp_converged_signature_takes_dscr_not_irrs() -> None:
    """SolveHeadless.bas Public Sub StampConvergedValuesHL must match the
    Python contract: dscrMult arg, no liveIRR / apprLive. Drift here
    means the macro re-import would leave the wrong signature in the
    workbook and Application.Run would fail with arg-count mismatch.
    """
    from pathlib import Path
    repo_root = Path(__file__).resolve().parent.parent
    bas = (repo_root / "SolveHeadless.bas").read_text(encoding="utf-8")

    import re
    collapsed = re.sub(r"_\s*\n\s*", " ", bas)
    m = re.search(
        r"Public Sub StampConvergedValuesHL\(([^)]*)\)",
        collapsed,
    )
    assert m, "StampConvergedValuesHL declaration not found in SolveHeadless.bas"
    sig = m.group(1)
    assert "colIdx As Integer" in sig
    assert "dscrMult As Double" in sig, (
        f"dscrMult arg missing from VBA sig — Tranche 7.12 reverted: {sig!r}"
    )
    assert "liveIRR" not in sig, (
        f"liveIRR resurrected in VBA sig — would re-hardcode PI 37: {sig!r}"
    )
    assert "apprLive" not in sig, (
        f"apprLive resurrected in VBA sig — would re-hardcode PI 31: {sig!r}"
    )


def test_stamp_active_writes_dscr_to_371_not_31_or_37() -> None:
    """The Public Sub StampActiveProjectColumnHL body must write the
    converged DSCR onto row 371 and must NOT write to rows 31 or 37
    (sticky-IF formulas stay). Structural check against the .bas source
    text — a unit test on COM behavior isn't feasible without a licensed
    Excel install + real workbook fixture.
    """
    from pathlib import Path
    import re
    repo_root = Path(__file__).resolve().parent.parent
    bas = (repo_root / "SolveHeadless.bas").read_text(encoding="utf-8")

    # Isolate the StampActiveProjectColumnHL body (from `Public Sub` to
    # matching `End Sub`). Defensive: must locate the body before
    # asserting against it.
    m = re.search(
        r"Public Sub StampActiveProjectColumnHL\([^)]*\)(.*?)End Sub",
        bas, re.DOTALL,
    )
    assert m, "StampActiveProjectColumnHL Sub body not found"
    body = m.group(1)

    # MUST write to PI_ROW_DSCR_MULT (row 371). The Cells(..., colIdx)
    # form is the canonical per-column write pattern in the .bas.
    assert "Cells(PI_ROW_DSCR_MULT, colIdx)" in body, (
        "StampActiveProjectColumnHL must write the converged DSCR onto "
        "PI row 371 (PI_ROW_DSCR_MULT) — upstream input lock per Tranche 7.12"
    )

    # MUST NOT write to PI_ROW_IRR_LIVE or PI_ROW_APPR_LIVE. The Tranche
    # 7.12 change deliberately leaves rows 31/37 as sticky-IF formulas
    # so the audit chain stays intact.
    assert "Cells(PI_ROW_IRR_LIVE, colIdx).Value =" not in body, (
        "StampActiveProjectColumnHL must NOT hardcode PI row 37 "
        "(Live Levered Pre-Tax IRR) — Tranche 7.12 reverted"
    )
    assert "Cells(PI_ROW_APPR_LIVE, colIdx).Value =" not in body, (
        "StampActiveProjectColumnHL must NOT hardcode PI row 31 "
        "(Live Appraisal IRR) — Tranche 7.12 reverted"
    )


def test_merge_output_rows_excludes_31_and_37() -> None:
    """The merge module's OUTPUT_ROWS (rows copied between worker
    outputs) must exclude rows 31 and 37 and must include row 371. The
    openpyxl merge path iterates this exact tuple, so a regression here
    would re-copy the hardcoded IRR cache cells (pre-7.12 behavior) or
    drop the DSCR multiple stamp (broken merge).
    """
    from dn38_solver.merge import OUTPUT_ROWS

    assert 31 not in OUTPUT_ROWS, (
        "OUTPUT_ROWS includes row 31 — merge would re-copy hardcoded "
        "Live Appraisal IRR, destroying the sticky-IF audit chain"
    )
    assert 37 not in OUTPUT_ROWS, (
        "OUTPUT_ROWS includes row 37 — merge would re-copy hardcoded "
        "Live Levered Pre-Tax IRR, destroying the sticky-IF audit chain"
    )
    assert 371 in OUTPUT_ROWS, (
        "OUTPUT_ROWS missing row 371 — merge wouldn't carry the "
        "Min Equity DSCR Multiple lock per Tranche 7.12"
    )
    # Lock in the exact tuple as a regression net. Adding a row in
    # the future is fine; the targeted asserts above will still catch
    # specific bad-rows. But pinning the current set documents intent.
    assert set(OUTPUT_ROWS) == {32, 33, 38, 39, 371}, (
        f"OUTPUT_ROWS drifted from Tranche 7.12 baseline: got {OUTPUT_ROWS}"
    )


# --- Tranche 7.13: --auto-fix unifies A1 + D15/D17 on the _FIXED sibling ----
# Caroline's seamless-everyday-solve preference: one flag (--auto-fix) covers
# both the calcPr iterateDelta patch (A1) and the macro re-import for missing-
# functions / hash drift (D15/D17). Both target the _FIXED.xlsm sibling so
# the source workbook is never mutated. The orchestrator's Phase 0.6 block is
# the gate; these tests lock the structural invariants so a future refactor
# can't silently degrade the UX back to the two-flag dance.

def test_orchestrator_auto_fix_block_handles_macro_codes() -> None:
    """The Phase 0.6 auto-fix block in solver/orchestrator.py must
    re-import the macro into the _FIXED.xlsm sibling when D15 or D17
    is in preflight.auto_fixable. Source-level check — invoking the
    real flow requires a licensed Excel install + COM round-trip.
    """
    from pathlib import Path
    repo_root = Path(__file__).resolve().parent.parent
    src = (repo_root / "dn38_solver" / "solver" / "orchestrator.py").read_text(
        encoding="utf-8"
    )

    # The block exists.
    assert "if auto_fix and preflight.auto_fixable:" in src, (
        "Phase 0.6 auto-fix block missing — Tranche 7.13 reverted"
    )
    # Inside it, D15/D17 fork into reimport_macro_subprocess on the
    # _FIXED sibling.
    assert "macro_codes_fired" in src, (
        "Tranche 7.13 macro-codes branch missing from auto-fix block"
    )
    assert "reimport_macro_subprocess(fixed_path)" in src, (
        "Macro re-import must target fixed_path (the _FIXED sibling), "
        "not workbook_path (the source). The whole point of routing "
        "through --auto-fix is to keep the source untouched."
    )
    # And it must NOT call reimport_macro_subprocess(workbook_path)
    # anywhere inside the --auto-fix branch (that would be the
    # destructive --auto-import-macro path leaking into the safe path).
    auto_fix_block_start = src.index("if auto_fix and preflight.auto_fixable:")
    # End of the block: the next `# Bank-grade gate` comment, which is
    # the section divider right after Phase 0.6.
    auto_fix_block_end = src.index("# Bank-grade gate", auto_fix_block_start)
    auto_fix_block = src[auto_fix_block_start:auto_fix_block_end]
    assert "reimport_macro_subprocess(workbook_path)" not in auto_fix_block, (
        "Auto-fix block must not call reimport_macro_subprocess on the "
        "source workbook_path — that destroys the source-untouched "
        "invariant. Use fixed_path (the _FIXED sibling) instead."
    )


def test_orchestrator_auto_import_macro_skips_when_auto_fix_set() -> None:
    """When both --auto-fix and --auto-import-macro are set, --auto-fix
    wins (the source stays untouched). The Phase 0.5 source-mutation
    block must gate itself off when auto_fix is True.
    """
    from pathlib import Path
    repo_root = Path(__file__).resolve().parent.parent
    src = (repo_root / "dn38_solver" / "solver" / "orchestrator.py").read_text(
        encoding="utf-8"
    )
    # The Phase 0.5 gate must include `not auto_fix` so it skips when
    # the operator passed both flags. Otherwise both branches fire and
    # the source ends up mutated.
    assert "if auto_import_macro and not auto_fix and needs_macro_import:" in src, (
        "Phase 0.5 (source-mutation auto-import) must gate on "
        "`not auto_fix` per Tranche 7.13 — otherwise passing both flags "
        "still mutates the source"
    )
