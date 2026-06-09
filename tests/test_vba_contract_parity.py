"""Contract-parity tests: the Python<->VBA boundary must stay in lockstep
with SolveHeadless.bas.

These are pure-Python (no Excel COM): they parse the in-repo
SolveHeadless.bas as text and compare it against the Python-side
declarations. They turn three drift modes that previously failed only at
runtime — after a full multi-minute solve — into fast CI failures:

  * C4 — preflight's REQUIRED_MACRO_FUNCTIONS missing a Sub Python calls
          via Application.Run (so D15/D18 never checked it).
  * C5 — vba_contract.VBASub.args arity diverging from the .bas signature
          (DISP_E_BADPARAMCOUNT at Application.Run time).
  * C13 — the Excel layout constants (PI rows, base col) duplicated in
          Python (config.py) and VBA drifting apart, silently mis-keying
          NPP / Dev Fee / FMV reads/writes.
"""
from __future__ import annotations

import re
from pathlib import Path

from dn38_solver import config
from dn38_solver.com.vba_contract import ALL_PUBLIC_SUBS
from dn38_solver.shadow.preflight import (
    REQUIRED_MACRO_FUNCTIONS,
    _parse_param_counts,
)

_REPO_ROOT = Path(__file__).resolve().parent.parent
_BAS_PATH = _REPO_ROOT / "SolveHeadless.bas"


def _bas_source() -> str:
    # latin-1 matches how preflight reads the .bas (VBA exports are not UTF-8).
    return _BAS_PATH.read_text(encoding="latin-1")


def test_required_macro_functions_covers_all_public_subs() -> None:
    """C4: every Sub Python invokes via Application.Run (the boundary
    contract) must appear in preflight's REQUIRED_MACRO_FUNCTIONS, so the
    D15 presence and D18 signature checks actually cover it. The two that
    were missing — SwitchProjectAndRecalc and the 6-arg StampConvergedValuesHL
    merge stamp — are the regression this locks shut.
    """
    contract_names = {s.name for s in ALL_PUBLIC_SUBS}
    missing = contract_names - set(REQUIRED_MACRO_FUNCTIONS)
    assert not missing, (
        "Application.Run entry point(s) absent from "
        f"REQUIRED_MACRO_FUNCTIONS (preflight will not verify them): {missing}"
    )


def test_contract_arity_matches_bas_signatures() -> None:
    """C5: VBASub.args arity must equal the parameter count of the matching
    Sub declaration in SolveHeadless.bas. A mismatch means Application.Run
    would raise DISP_E_BADPARAMCOUNT (0x8002000E) at runtime — after the
    full solve cost has been paid.
    """
    bas_counts = _parse_param_counts(_bas_source())
    mismatches: list[str] = []
    for sub in ALL_PUBLIC_SUBS:
        if sub.name not in bas_counts:
            mismatches.append(f"{sub.name}: not found in SolveHeadless.bas")
            continue
        if bas_counts[sub.name] != len(sub.args):
            mismatches.append(
                f"{sub.name}: contract={len(sub.args)} arg(s), "
                f"bas={bas_counts[sub.name]} arg(s)"
            )
    assert not mismatches, "Python<->VBA arity drift:\n  " + "\n  ".join(mismatches)


def _bas_long_constants() -> dict[str, int]:
    """Parse `Private Const <NAME> As Long/Integer = <int>` from the .bas."""
    src = _bas_source()
    out: dict[str, int] = {}
    for m in re.finditer(
        r"Private Const (\w+)\s+As (?:Long|Integer)\s*=\s*(\d+)", src
    ):
        out.setdefault(m.group(1), int(m.group(2)))
    return out


def test_bas_layout_constants_match_config() -> None:
    """C13: the Project Inputs layout constants are declared independently
    in config.py (Python) and SolveHeadless.bas (VBA). A one-row template
    insert that updates one side but not the other silently mis-keys
    NPP/DevFee/FMV with no error. Pin them together.
    """
    bas = _bas_long_constants()

    # Scalar layout anchors: exact equality with config.py.
    scalar_pairs = {
        "PI_ROW_NAME": config.PROJECT_NAME_ROW,
        "PI_ROW_TOGGLE": config.PROJECT_TOGGLE_ROW,
        "PI_BASE_COL": config.BASE_COL,
    }
    for name, expected in scalar_pairs.items():
        assert name in bas, f"{name} not found in SolveHeadless.bas"
        assert bas[name] == expected, (
            f"{name}={bas[name]} in .bas but config={expected}"
        )

    # Per-project output rows: each .bas row constant must be a key in
    # config.OUTPUT_ROWS (the canonical row->label map Python reads back).
    output_row_consts = (
        "PI_ROW_APPR_LIVE",   # 31
        "PI_ROW_DEV_FEE",     # 32
        "PI_ROW_FMV",         # 33
        "PI_ROW_IRR_LIVE",    # 37
        "PI_ROW_NPP",         # 38
        "PI_ROW_NPP_TOTAL",   # 39
        "PI_ROW_DSCR_MULT",   # 371
    )
    for name in output_row_consts:
        assert name in bas, f"{name} not found in SolveHeadless.bas"
        assert bas[name] in config.OUTPUT_ROWS, (
            f"{name}={bas[name]} in .bas is not a key in config.OUTPUT_ROWS "
            f"{sorted(config.OUTPUT_ROWS)} — Python read-back would mis-key this row"
        )
