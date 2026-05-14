"""dn38_solver.shadow.preflight — Bank-grade pre-flight validation.

A single workbook check that fails LOUD before paying COM startup +
multi-minute solve cost on a workbook that won't converge.

Three categories of checks, each with a stable error code:

  A. Workbook calc-property failures. These weaken or disable the
     iterative-calc engine that the macro's row 31 self-circular hard-
     stamp pattern relies on. Failures are subtle — convergence may
     still appear to work but with non-deterministic last-wins values.
       A1  iterateDelta missing or > 0.0001
       A2  iterate=False
       A3  calcMode=auto
       A4  fullCalcOnLoad=False
  B. Workbook structure failures (macro crashes mid-run with cryptic
     VBA errors when sheets/cells/protection don't match expectations).
       B5  Required sheets missing
       B7  Required cells missing or wrong type
       B8  Workbook or sheet structurally protected
       B9  No active project flagged in row 7
  C. Critical-path cached-error checks (errors on cells the macro reads
     during convergence — XIRR returns garbage when feeder rows have
     #VALUE!/#DIV/0!/#REF! tokens).
       C10 Errors on Project Inputs F30:F39 master column
       C11 Errors on Appraisal cash flow rows 155-159
       C12 Errors on Appraisal!H161 (the XIRR readout)
  D. Embedded VBA macro version. The repo's SolveHeadless.bas evolves
     (yesterday alone added phase-scoped recalc, hard-stamp post-read pass,
     parallel-runner trust gates). Workbooks where the macro hasn't been
     re-imported run an OUTDATED macro version that's missing critical
     functions; the orchestrator calls them via Application.Run and they
     silently fail (or fall through On Error Resume Next), leaving the
     solve in a half-broken state. Confirmed root cause of the IL TEST
     2026-05-13 regression: that workbook had a 627-line macro vs the
     1220-line current; missing CalcSheetsForAppraisal, ClassifyConver-
     genceHL, StampActiveProjectColumnHL, and 8 others.
       D15 Embedded macro is missing required functions
       D16 Embedded macro has leftover stale modules
  E. Input-range checks against the macro's hardcoded bounds. The chunked
     entry point (SolveOneProjectByColHL) resets pre-solve Dev Fee / NPP
     to seed values when they fall outside [DEV_FEE_MIN..MAX]/[NPP_MIN..
     MAX]. For models where the natural converged Dev Fee is well above
     DEV_FEE_MAX (e.g., utility solar where Dev Fee can be $1.50-$2.50/W),
     this reset destroys the starting state and GoalSeek may not recover.
       E13 Pre-solve Dev Fee outside macro bounds
       E14 Pre-solve NPP outside macro bounds

Findings carry a remediation string and an auto_fixable flag. Today
only A1 (iterateDelta) is auto-fixable; A2/E13/E14 could be added but
mutating model inputs / iterate flag without operator audit is risky.

Bank-grade defaults: errors always abort the run. Warnings are visible
in the log but proceed unless --strict-preflight is set.
"""
from __future__ import annotations

import logging
import re
import shutil
import zipfile
from pathlib import Path

import msgspec
import openpyxl
from openpyxl.utils import column_index_from_string

from dn38_solver.shadow.validation import (
    EXCEL_ERROR_TOKENS,
    WorkbookValidation,
    scan_workbook_errors,
)

log = logging.getLogger(__name__)

# Tolerance ceiling for the iterative-calc engine. Below this value GoalSeek's
# Appraisal inner loop converges reliably across the IL/SMP/RP-Puma models.
# Above it (Excel's default 0.001) the loop bisects on half-converged H161
# and slides Dev Fee to floor — observed on 38DN-IL_US Solar 2026-05-13.
ITERATE_DELTA_CEILING = 0.0001

# Sheets the macro requires. Missing any of these crashes SolveHeadless or
# the chunked entry points with a "Subscript out of range" — surface that
# upfront with a named sheet rather than letting the macro die mid-run.
REQUIRED_SHEETS = (
    "Project Inputs",
    "Appraisal",
    "NPP Calc",
    "Operations",
    "PT Returns",
    "Tax Equity",
    "Perm Debt",
    "CL",
    "Capex",
    "Rate Curves",
    "Global",
)

# Master-column cells the macro reads/writes during the solve. F2 is the
# active-project index, F30/F36 are targets (WACC / Equity), F31/F37 are
# live readouts the GoalSeek operates on, F32 is the changing Dev Fee.
REQUIRED_PI_CELLS = {
    "F2": "active project index",
    "F30": "FMV WACC target",
    "F31": "Live Appraisal IRR (GoalSeek target var)",
    "F32": "Dev Fee (GoalSeek changing cell)",
    "F36": "Equity IRR target",
    "F37": "Live Levered Pre-Tax IRR",
}

# Ranges the macro reads or that feed the GoalSeek convergence variable.
# Critical-path errors here corrupt the inner loop irrespective of model
# inputs — distinct from the broad workbook-wide scan in validation.py.
PI_CRITICAL_PATH_ROWS = (30, 31, 32, 33, 36, 37, 38, 39)
APPRAISAL_CASHFLOW_ROWS = (155, 156, 157, 158, 159)

# Hardcoded bounds in SolveHeadless.bas (lines 46-51). Mirror them here so
# pre-flight can warn before the macro's chunked path resets out-of-range
# values to seed values. If the .bas constants change, update these too —
# there's no clean import path because the macro is not Python.
NPP_BOUNDS = (-0.2, 0.8)         # NPP_MIN, NPP_MAX
DEV_FEE_BOUNDS = (0.05, 0.5)     # DEV_FEE_MIN, DEV_FEE_MAX
PI_ROW_NPP = 38
PI_ROW_DEV_FEE = 32
PI_ROW_TOGGLE = 7
PI_PROJECT_COL_RANGE = ("H", "S")  # 8..19, project columns

# Marker function names that must exist in the embedded macro for the
# CURRENT orchestrator to drive the workbook correctly. Each is an entry
# point or core helper added in a specific commit; absence implies the
# workbook hasn't had the latest macro re-imported. We scan the binary
# vbaProject.bin for these as plain ASCII tokens — VBA's compressed-
# storage format leaves identifier strings readable. Cheap (~10ms on the
# 215KB binary) and avoids requiring Excel COM at preflight time.
REQUIRED_MACRO_FUNCTIONS = (
    "SolveHeadless",                # entry point (single-shot)
    "InitSolveEnvHL",               # entry point (chunked init)
    "SolveOneProjectByColHL",       # entry point (chunked per-project)
    "FinalizeSolveEnvHL",           # entry point (chunked finalize)
    "CalcSheetsForAppraisal",       # phase-scoped recalc (added 2026-04-29)
    "CalcSheetsForNPP",             # phase-scoped recalc
    "CalcSheetsForDSCR",            # phase-scoped recalc
    "ClassifyConvergenceHL",        # strict/relaxed/none tier classifier
    "StampActiveProjectColumnHL",   # post-read hard-stamps (added 2026-05-12)
    "ProjectElapsedHL",             # timer wraparound-safe elapsed
    "HardStampNumericHL",           # IsError-guarded numeric stamping
    "SetSkipOutputRecalcHL",        # output-recalc skip flag (added 2026-05-12)
)

# Names of stale module artifacts that linger in workbooks where macros
# were imported then partially replaced. Each is a hint that the workbook
# has been through several import cycles without the cleanup pass that
# removes the prior versions. Not a hard blocker but worth surfacing.
STALE_MACRO_MODULES = (
    "Module2_Optimized",
    "Module3",
    "Module4",
)


class PreflightFinding(msgspec.Struct, frozen=True, kw_only=True):
    """One actionable finding from the pre-flight pass."""
    code: str
    severity: str  # "error" | "warning" | "info"
    location: str
    message: str
    impact: str
    remediation: str
    auto_fixable: bool = False


class PreflightResult(msgspec.Struct, frozen=True, kw_only=True):
    """Aggregate result of a workbook pre-flight pass."""
    workbook_path: str
    findings: tuple[PreflightFinding, ...]
    error_scan: WorkbookValidation

    @property
    def errors(self) -> tuple[PreflightFinding, ...]:
        return tuple(f for f in self.findings if f.severity == "error")

    @property
    def warnings(self) -> tuple[PreflightFinding, ...]:
        return tuple(f for f in self.findings if f.severity == "warning")

    @property
    def auto_fixable(self) -> tuple[PreflightFinding, ...]:
        return tuple(f for f in self.findings if f.auto_fixable)

    @property
    def ok(self) -> bool:
        """No error-severity findings."""
        return len(self.errors) == 0


# ---------------------------------------------------------------------------
# Individual check functions. Each returns a list (possibly empty) of
# PreflightFindings. Kept as pure functions over an open openpyxl workbook
# so they're trivially testable with synthetic in-memory workbooks.
# ---------------------------------------------------------------------------

def check_calc_properties(wb: openpyxl.Workbook) -> list[PreflightFinding]:
    """A1-A4: workbook calculation properties."""
    findings: list[PreflightFinding] = []
    cp = wb.calculation

    # A1: iterateDelta. Empirically observed missing on IL TEST 2026-05-13
    # alongside the actual root cause (D15, outdated macro). Not by itself
    # a confirmed convergence breaker, but a non-standard XML state that
    # weakens iterative-calc convergence on the row 31 self-circular cells
    # the macro depends on. SMP and RP Puma all have this set explicitly
    # to 1E-4; absence is anomalous and worth flagging loud.
    if cp.iterateDelta is None or cp.iterateDelta > ITERATE_DELTA_CEILING:
        findings.append(PreflightFinding(
            code="A1",
            severity="error",
            location="workbook calcPr",
            message=(
                f"iterateDelta is {cp.iterateDelta!r} (must be <= {ITERATE_DELTA_CEILING})"
            ),
            impact=(
                "Iterative-calc engine may exit before circular references "
                "settle, causing the macro's row 31 self-circular per-project "
                "hard-stamps to capture stale values. All baseline pricing "
                "models on disk have this set explicitly to 1E-4; missing "
                "value indicates non-standard save state."
            ),
            remediation=(
                "In Excel: File > Options > Formulas > Maximum Change = 0.0001. "
                "Save. Or rerun with --auto-fix to patch a copy."
            ),
            auto_fixable=True,
        ))

    # A2: iterative calc must be enabled. Row 31 self-circular cells
    # (=IF(H2=$F$2,$F$31,H31)) become #REF! when iterate=False.
    if not cp.iterate:
        findings.append(PreflightFinding(
            code="A2",
            severity="error",
            location="workbook calcPr",
            message="iterative calculation is disabled (iterate=False)",
            impact=(
                "Project Inputs row 31 self-circular formulas error out, "
                "preventing per-column appraisal IRR capture. The post-solve "
                "hard-stamps will read #REF!."
            ),
            remediation=(
                "In Excel: File > Options > Formulas > "
                "Enable iterative calculation (checked). Save."
            ),
            auto_fixable=False,  # Doable but risky to flip without model audit
        ))

    # A3: calcMode auto causes uncontrolled recalc during macro runs.
    if cp.calcMode and cp.calcMode != "manual":
        findings.append(PreflightFinding(
            code="A3",
            severity="warning",
            location="workbook calcPr",
            message=f"calcMode is {cp.calcMode!r} (expected 'manual')",
            impact=(
                "Excel may trigger background recalcs during the macro, racing "
                "with the macro's own targeted .Calculate calls. Convergence "
                "behavior becomes nondeterministic."
            ),
            remediation=(
                "In Excel: Formulas tab > Calculation Options > Manual. Save."
            ),
            auto_fixable=False,
        ))

    # A4: fullCalcOnLoad ensures the cached values seen by openpyxl scans
    # actually reflect the formulas' current state.
    if cp.fullCalcOnLoad is False:  # explicit False, not None
        findings.append(PreflightFinding(
            code="A4",
            severity="warning",
            location="workbook calcPr",
            message="fullCalcOnLoad is False",
            impact=(
                "Excel won't recalculate on open, so cached values may not "
                "reflect current formulas. Pre-flight error scans (#VALUE!, "
                "#DIV/0!, etc.) become unreliable."
            ),
            remediation=(
                "Open in Excel, press F9 to force full recalc, save. "
                "Alternative: have the model owner enable Workbook > "
                "Calculation > 'Recalculate workbook before saving'."
            ),
            auto_fixable=False,
        ))

    return findings


def check_required_sheets(wb: openpyxl.Workbook) -> list[PreflightFinding]:
    """B5: required sheets must be present."""
    present = set(wb.sheetnames)
    missing = [s for s in REQUIRED_SHEETS if s not in present]
    if not missing:
        return []
    return [PreflightFinding(
        code="B5",
        severity="error",
        location="workbook",
        message=f"missing required sheet(s): {', '.join(missing)}",
        impact=(
            "Macro will fail with VBA 'Subscript out of range' when it "
            "calls Sheets(\"<missing>\")."
        ),
        remediation=(
            f"Restore the missing sheet(s) from a baseline pricing model. "
            f"Required: {', '.join(REQUIRED_SHEETS)}."
        ),
        auto_fixable=False,
    )]


def check_required_cells(wb: openpyxl.Workbook) -> list[PreflightFinding]:
    """B7: master-column cells the macro depends on must exist."""
    if "Project Inputs" not in wb.sheetnames:
        return []  # B5 already flagged this
    ws = wb["Project Inputs"]
    findings: list[PreflightFinding] = []
    for cell, role in REQUIRED_PI_CELLS.items():
        v = ws[cell].value
        if v is None:
            findings.append(PreflightFinding(
                code="B7",
                severity="error",
                location=f"Project Inputs!{cell}",
                message=f"required cell is empty (role: {role})",
                impact=(
                    "Macro reads or writes this cell every iteration. "
                    "Empty value will produce wrong results or VBA error."
                ),
                remediation=(
                    f"Restore Project Inputs!{cell} from a baseline model. "
                    f"This cell holds the {role}."
                ),
                auto_fixable=False,
            ))
    return findings


def check_workbook_protection(wb: openpyxl.Workbook) -> list[PreflightFinding]:
    """B8: workbook and critical-sheet protection."""
    findings: list[PreflightFinding] = []
    # Workbook structure protection (prevents adding/renaming sheets — the
    # macro adds __SolverResults).
    try:
        wb_protected = wb.security and (
            getattr(wb.security, "lockStructure", False)
            or getattr(wb.security, "lockWindows", False)
        )
    except Exception:
        wb_protected = False
    if wb_protected:
        findings.append(PreflightFinding(
            code="B8",
            severity="error",
            location="workbook",
            message="workbook structure is protected",
            impact=(
                "Macro will fail when adding the __SolverResults telemetry "
                "sheet at startup. Protection blocks all sheet add/rename."
            ),
            remediation=(
                "In Excel: Review tab > Protect Workbook (toggle off). "
                "Provide password if one was set."
            ),
            auto_fixable=False,
        ))

    # Per-sheet protection on critical sheets the macro writes to.
    for sname in ("Project Inputs", "Appraisal", "NPP Calc"):
        if sname not in wb.sheetnames:
            continue
        ws = wb[sname]
        if ws.protection.sheet:
            findings.append(PreflightFinding(
                code="B8",
                severity="error",
                location=sname,
                message=f"sheet '{sname}' is protected",
                impact=(
                    f"Macro writes to '{sname}' during solve (project index, "
                    f"hard-stamps, calc triggers). Protection blocks writes."
                ),
                remediation=(
                    f"In Excel: select sheet '{sname}' > Review tab > "
                    f"Unprotect Sheet."
                ),
                auto_fixable=False,
            ))
    return findings


def check_active_projects(wb: openpyxl.Workbook) -> list[PreflightFinding]:
    """B9: at least one project must be flagged active in row 7."""
    if "Project Inputs" not in wb.sheetnames:
        return []  # B5 covered this
    ws = wb["Project Inputs"]
    # Project columns are H..S (8..19). Row 7 is the active flag (1=on).
    active_cols = []
    for col in range(column_index_from_string("H"), column_index_from_string("S") + 1):
        v = ws.cell(row=7, column=col).value
        if v == 1:
            active_cols.append(col)
    if not active_cols:
        return [PreflightFinding(
            code="B9",
            severity="error",
            location="Project Inputs!H7:S7",
            message="no active projects flagged (row 7 = 1 in any project column)",
            impact="Solver will exit immediately with 'No active projects found'.",
            remediation=(
                "Set row 7 to 1 for the project columns you want to solve. "
                "Inactive projects must be 0 or blank."
            ),
            auto_fixable=False,
        )]
    return []


def check_critical_path_errors(
    wb_values: openpyxl.Workbook,
    error_scan: WorkbookValidation,
) -> list[PreflightFinding]:
    """C10-C12: cached error tokens on cells the convergence depends on.

    Uses the data_only=True workbook to read cached error tokens. Falls
    back to the full error scan's reported locations to detect criticality
    in O(broken-cell-count) rather than re-scanning the whole workbook.
    """
    findings: list[PreflightFinding] = []

    def is_error(val: object) -> bool:
        return isinstance(val, str) and any(tok in val for tok in EXCEL_ERROR_TOKENS)

    # C10: PI master column (F30:F39) and active project columns (H..S)
    if "Project Inputs" in wb_values.sheetnames:
        ws = wb_values["Project Inputs"]
        bad_cells: list[str] = []
        for row in PI_CRITICAL_PATH_ROWS:
            for col in range(6, column_index_from_string("S") + 1):  # F..S
                v = ws.cell(row=row, column=col).value
                if is_error(v):
                    bad_cells.append(f"Project Inputs!{ws.cell(row=row, column=col).coordinate}={v}")
                    if len(bad_cells) >= 10:
                        break
            if len(bad_cells) >= 10:
                break
        if bad_cells:
            findings.append(PreflightFinding(
                code="C10",
                severity="error",
                location="Project Inputs F30:S39",
                message=(
                    f"{len(bad_cells)} cached error(s) on convergence-critical cells "
                    f"(showing first {min(10, len(bad_cells))}: {'; '.join(bad_cells[:5])}"
                    f"{'; ...' if len(bad_cells) > 5 else ''})"
                ),
                impact=(
                    "Macro reads these rows for the GoalSeek target/changing "
                    "cells and per-project hard-stamps. Errors here corrupt "
                    "the entire convergence loop."
                ),
                remediation=(
                    "Open the workbook in Excel, navigate to the listed cells, "
                    "and trace the precedent chain to find the source of the "
                    "error. Common causes: missing rate curves, broken named "
                    "ranges, deleted source rows."
                ),
                auto_fixable=False,
            ))

    # C11/C12: Appraisal cash flow rows + H161 readout
    if "Appraisal" in wb_values.sheetnames:
        ws = wb_values["Appraisal"]
        bad_cf: list[str] = []
        # Active cash-flow region only — the OFFSET picks columns to the
        # right of where the CT comes online, so scan J..UC bounded.
        max_col = min(ws.max_column, column_index_from_string("UC"))
        for row in APPRAISAL_CASHFLOW_ROWS:
            for col in range(column_index_from_string("J"), max_col + 1):
                v = ws.cell(row=row, column=col).value
                if is_error(v):
                    bad_cf.append(f"Appraisal!{ws.cell(row=row, column=col).coordinate}={v}")
                    if len(bad_cf) >= 10:
                        break
            if len(bad_cf) >= 10:
                break
        if bad_cf:
            findings.append(PreflightFinding(
                code="C11",
                severity="error",
                location="Appraisal rows 155-159",
                message=(
                    f"{len(bad_cf)} cached error(s) on Appraisal cash flow rows "
                    f"({'; '.join(bad_cf[:5])}{'; ...' if len(bad_cf) > 5 else ''})"
                ),
                impact=(
                    "Cash flow row 159 = SUM(155:158) feeds the XIRR readout "
                    "at H161. Errors here propagate to H161 and corrupt the "
                    "GoalSeek target value."
                ),
                remediation=(
                    "Trace precedents on the listed cells. Often caused by "
                    "missing inputs in Project Inputs row 32 (Dev Fee), row 11 "
                    "(System Size), or rate curve gaps in Rate Curves."
                ),
                auto_fixable=False,
            ))

        h161 = ws["H161"].value
        if is_error(h161):
            findings.append(PreflightFinding(
                code="C12",
                severity="error",
                location="Appraisal!H161",
                message=f"H161 (Live Appraisal IRR XIRR readout) is {h161}",
                impact=(
                    "F31 = Appraisal!H161 is the GoalSeek target variable. "
                    "An error here means GoalSeek has no valid value to drive "
                    "and will fail immediately."
                ),
                remediation=(
                    "Fix the upstream errors first (likely C11 findings) and "
                    "force a full recalc (F9) before re-running."
                ),
                auto_fixable=False,
            ))

    return findings


def check_macro_version(workbook_path: Path) -> list[PreflightFinding]:
    """D15/D16: embedded macro must have all required functions and no
    stale leftover modules.

    Scans xl/vbaProject.bin for ASCII function-name tokens. VBA's
    compressed-storage format leaves identifier strings readable as
    plain ASCII; a substring match for `\\x00<name>` (the typical
    surrounding bytes) is more precise than a bare substring search,
    but a bare search is good enough for our marker functions which
    don't appear elsewhere in the binary.

    Validated against the IL TEST 2026-05-13 regression: that workbook's
    627-line embedded macro is missing 11 of the 12 marker functions
    listed in REQUIRED_MACRO_FUNCTIONS.
    """
    findings: list[PreflightFinding] = []
    # .xlsx is macro-free by spec; skip D-tier entirely. Tests use .xlsx
    # synthetic workbooks, and the solver targets .xlsm only — checking
    # for a missing macro on .xlsx would be noise.
    if workbook_path.suffix.lower() != ".xlsm":
        return findings

    try:
        with zipfile.ZipFile(workbook_path, "r") as z:
            try:
                vba_bin = z.read("xl/vbaProject.bin")
            except KeyError:
                # No VBA project at all in an .xlsm. Solver can't drive a
                # macro-less workbook; flag as a structural error.
                return [PreflightFinding(
                    code="D15",
                    severity="error",
                    location="xl/vbaProject.bin",
                    message="workbook contains no VBA project",
                    impact=(
                        "Solver requires the SolveHeadless macro to be "
                        "embedded. Without it, no entry point exists for "
                        "Application.Run."
                    ),
                    remediation=(
                        "Run: python import_vba_module.py "
                        f'"{workbook_path}"'
                    ),
                    auto_fixable=False,
                )]
    except (zipfile.BadZipFile, OSError) as exc:
        # Defer to scan_workbook_errors's X0 handling; don't double-report.
        log.debug("D-tier scan: cannot open zip (%s): %s", workbook_path, exc)
        return findings

    # Scan as latin-1 so every byte maps to a 1-char string (identifier
    # tokens are ASCII either way; latin-1 is just the safe decode).
    text = vba_bin.decode("latin-1", errors="replace")

    missing = [name for name in REQUIRED_MACRO_FUNCTIONS if name not in text]
    if missing:
        findings.append(PreflightFinding(
            code="D15",
            severity="error",
            location="xl/vbaProject.bin",
            message=(
                f"embedded macro is missing {len(missing)} of "
                f"{len(REQUIRED_MACRO_FUNCTIONS)} required function(s): "
                f"{', '.join(missing)}"
            ),
            impact=(
                "The orchestrator calls these functions via Application.Run "
                "during the chunked solve path. Missing functions either "
                "fail silently (On Error Resume Next) or raise a VBA error, "
                "leaving the solve in a half-broken state. ROOT CAUSE of "
                "the IL TEST 2026-05-13 regression."
            ),
            remediation=(
                f'Re-import the latest macro: python import_vba_module.py '
                f'"{workbook_path}"'
            ),
            auto_fixable=False,  # Requires Excel COM; not safe to auto-do
        ))

    stale_present = [m for m in STALE_MACRO_MODULES if m in text]
    if stale_present:
        findings.append(PreflightFinding(
            code="D16",
            severity="warning",
            location="xl/vbaProject.bin",
            message=(
                f"workbook contains {len(stale_present)} stale macro "
                f"module(s): {', '.join(stale_present)}"
            ),
            impact=(
                "Indicates the workbook has been through several macro-"
                "import cycles without cleanup. Stale modules don't break "
                "the solve directly but may shadow current function names "
                "or contain references to obsolete cells. Worth removing."
            ),
            remediation=(
                "In Excel: Alt+F11 to open VBA editor; right-click each "
                "stale module under VBAProject and choose 'Remove'. "
                "Decline the export prompt unless you want a backup."
            ),
            auto_fixable=False,
        ))

    return findings


def check_input_bounds(wb: openpyxl.Workbook) -> list[PreflightFinding]:
    """E13/E14: pre-solve Dev Fee / NPP per project must be within the
    chunked entry point's hardcoded bounds, else the macro resets them
    to seed at iteration 0 and GoalSeek bisects from a destroyed state.
    """
    findings: list[PreflightFinding] = []
    if "Project Inputs" not in wb.sheetnames:
        return findings
    ws = wb["Project Inputs"]

    npp_lo, npp_hi = NPP_BOUNDS
    df_lo, df_hi = DEV_FEE_BOUNDS

    out_of_band_dev_fee: list[tuple[str, float]] = []
    out_of_band_npp: list[tuple[str, float]] = []

    start_col = column_index_from_string(PI_PROJECT_COL_RANGE[0])
    end_col = column_index_from_string(PI_PROJECT_COL_RANGE[1])
    for col in range(start_col, end_col + 1):
        toggle = ws.cell(row=PI_ROW_TOGGLE, column=col).value
        if toggle != 1:
            continue  # only check active project columns
        coord_letter = ws.cell(row=PI_ROW_DEV_FEE, column=col).coordinate

        df = ws.cell(row=PI_ROW_DEV_FEE, column=col).value
        if isinstance(df, (int, float)) and not (df_lo <= df <= df_hi):
            out_of_band_dev_fee.append((coord_letter, float(df)))

        npp_coord = ws.cell(row=PI_ROW_NPP, column=col).coordinate
        npp = ws.cell(row=PI_ROW_NPP, column=col).value
        if isinstance(npp, (int, float)) and not (npp_lo <= npp <= npp_hi):
            out_of_band_npp.append((npp_coord, float(npp)))

    if out_of_band_dev_fee:
        cells = ", ".join(f"{c}={v:.2f}" for c, v in out_of_band_dev_fee)
        findings.append(PreflightFinding(
            code="E13",
            severity="warning",  # SMP empirically converges with E13 firing.
            location=f"Project Inputs row {PI_ROW_DEV_FEE}",
            message=(
                f"pre-solve Dev Fee outside macro bounds [{df_lo}..{df_hi}] "
                f"on {len(out_of_band_dev_fee)} active project(s): {cells}"
            ),
            impact=(
                "The chunked solve path resets these to DEV_FEE_SEED ($0.20) "
                "at iteration 0 and after every inner GoalSeek call. Some "
                "models (e.g. SMP Greenlee) recover and converge anyway; "
                "others (e.g. IL TEST 2026-05-13) get trapped at the seed "
                "and the run reports '0 iter / NOT CONVERGED'. Worth "
                "investigating if convergence fails — but not always a "
                "blocker on its own."
            ),
            remediation=(
                "If the run fails to converge, two options: (1) Manually "
                "set the Project Inputs Dev Fee cells to a value within "
                f"[${df_lo}..${df_hi}/W] and re-run (risky if Dev Fee is a "
                "deal-side input). (2) Ask the model owner to raise "
                "DEV_FEE_MAX in SolveHeadless.bas to encompass the natural "
                "pricing range, re-import the macro, and re-run."
            ),
            auto_fixable=False,
        ))

    if out_of_band_npp:
        cells = ", ".join(f"{c}={v:.3f}" for c, v in out_of_band_npp)
        findings.append(PreflightFinding(
            code="E14",
            severity="warning",
            location=f"Project Inputs row {PI_ROW_NPP}",
            message=(
                f"pre-solve NPP outside macro bounds [{npp_lo}..{npp_hi}] "
                f"on {len(out_of_band_npp)} active project(s): {cells}"
            ),
            impact=(
                "Macro resets these to NPP_SEED ($0.20) at iteration 0. NPP "
                "is solved within the inner loop (driven by Equity IRR "
                "GoalSeek), so a bad seed is usually recoverable — but "
                "convergence may take more iterations than usual."
            ),
            remediation=(
                "If recurring, raise NPP_MIN/NPP_MAX in SolveHeadless.bas "
                "to match the deal pipeline's natural NPP range."
            ),
            auto_fixable=False,
        ))

    return findings


def run_preflight(workbook_path: Path | str) -> PreflightResult:
    """Run the full pre-flight pass against `workbook_path`.

    Pure function. Never mutates the input file. Returns a PreflightResult
    with categorized findings; caller decides how to act on them.
    """
    p = Path(workbook_path)
    findings: list[PreflightFinding] = []

    # Cell-level error scan (the existing validation module). Reused as
    # both a standalone diagnostic and as the input to the C-tier checks.
    error_scan = scan_workbook_errors(p)

    if error_scan.status == "scan_failed":
        # Can't open the file. Return a single error finding and bail —
        # the per-check passes below all need an open workbook.
        findings.append(PreflightFinding(
            code="X0",
            severity="error",
            location="workbook",
            message=f"unable to open workbook: {error_scan.error}",
            impact="Pre-flight cannot proceed.",
            remediation=(
                "Verify the file exists, is not currently open in another "
                "Excel instance, and is a valid .xlsm/.xlsx."
            ),
            auto_fixable=False,
        ))
        return PreflightResult(
            workbook_path=str(p),
            findings=tuple(findings),
            error_scan=error_scan,
        )

    # D-tier runs first (cheap zip-level read of vbaProject.bin) before
    # paying the openpyxl load cost. Catches outdated-macro workbooks
    # that would also tend to trip downstream checks confusingly.
    findings.extend(check_macro_version(p))

    # Two passes: formulas (calc props live here) + cached values (for C-checks)
    wb_f = openpyxl.load_workbook(str(p), data_only=False, keep_vba=True, read_only=False)
    try:
        findings.extend(check_calc_properties(wb_f))
        findings.extend(check_required_sheets(wb_f))
        findings.extend(check_required_cells(wb_f))
        findings.extend(check_workbook_protection(wb_f))
        findings.extend(check_active_projects(wb_f))
        findings.extend(check_input_bounds(wb_f))
    finally:
        wb_f.close()

    wb_v = openpyxl.load_workbook(str(p), data_only=True, read_only=True)
    try:
        findings.extend(check_critical_path_errors(wb_v, error_scan))
    finally:
        wb_v.close()

    return PreflightResult(
        workbook_path=str(p),
        findings=tuple(findings),
        error_scan=error_scan,
    )


# ---------------------------------------------------------------------------
# Auto-fix
# ---------------------------------------------------------------------------

_CALCPR_RE = re.compile(r"<calcPr[^>]*/?>")


def apply_auto_fixes(
    src: Path | str,
    dst: Path | str,
    findings: tuple[PreflightFinding, ...],
) -> tuple[Path, list[str]]:
    """Patch a copy of `src` to address auto-fixable findings.

    Returns (dst_path, applied_fix_codes). Findings without auto_fixable=True
    are skipped silently — caller is responsible for surfacing them as still-
    blocking errors after the fix pass.

    Bank-grade: never mutates the source file. Always writes a new file.
    Editing happens at the .xlsx/.zip layer rather than via openpyxl save
    to avoid round-tripping every formatted cell on a 13MB workbook (which
    risks unintended cosmetic changes to a heavily styled deliverable).
    """
    src_p = Path(src)
    dst_p = Path(dst)
    if dst_p.exists():
        dst_p.unlink()
    shutil.copy2(src_p, dst_p)

    applied: list[str] = []
    fixable_codes = {f.code for f in findings if f.auto_fixable}

    if "A1" in fixable_codes:
        _patch_iterate_delta_xml(dst_p, ITERATE_DELTA_CEILING)
        applied.append("A1")

    return dst_p, applied


def _patch_iterate_delta_xml(path: Path, value: float) -> None:
    """Add/overwrite iterateDelta="<value>" on the <calcPr/> tag in
    xl/workbook.xml of the .xlsm at `path`. In-place rewrite of the zip.
    """
    # Render value as Excel does: 0.0001 -> "1E-4". Match SMP convention.
    if value == 0.0001:
        rendered = "1E-4"
    elif value == 0.001:
        rendered = "1E-3"
    elif value == 0.00001:
        rendered = "1E-5"
    else:
        rendered = repr(value)

    tmp = path.with_suffix(path.suffix + ".tmp")
    with zipfile.ZipFile(path, "r") as zin:
        with zipfile.ZipFile(tmp, "w", compression=zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename == "xl/workbook.xml":
                    text = data.decode("utf-8")

                    def add_or_replace(m: re.Match[str]) -> str:
                        tag = m.group()
                        # Strip any existing iterateDelta="..."
                        tag = re.sub(r' iterateDelta="[^"]*"', "", tag)
                        # Insert before the closing /> or >
                        if tag.endswith("/>"):
                            return tag[:-2] + f' iterateDelta="{rendered}"/>'
                        return tag[:-1] + f' iterateDelta="{rendered}">'

                    text = _CALCPR_RE.sub(add_or_replace, text, count=1)
                    data = text.encode("utf-8")
                zout.writestr(item, data)
    tmp.replace(path)


# ---------------------------------------------------------------------------
# Reporting
# ---------------------------------------------------------------------------

def format_preflight_report(result: PreflightResult) -> str:
    """One-screen summary of a PreflightResult, suitable for log output."""
    lines: list[str] = []
    err_n = len(result.errors)
    warn_n = len(result.warnings)
    fix_n = len(result.auto_fixable)

    if result.ok and warn_n == 0:
        lines.append(f"  Pre-flight: PASS (0 findings)")
    else:
        status = "FAIL" if err_n else "WARN"
        lines.append(
            f"  Pre-flight: {status} - {err_n} error(s), {warn_n} warning(s), "
            f"{fix_n} auto-fixable"
        )

    for f in result.findings:
        bullet = "[ERROR]" if f.severity == "error" else "[WARN] "
        fix_tag = " (auto-fixable)" if f.auto_fixable else ""
        lines.append(f"  {bullet} {f.code} @ {f.location}{fix_tag}")
        lines.append(f"          {f.message}")
        lines.append(f"          impact: {f.impact}")
        lines.append(f"          fix:    {f.remediation}")

    return "\n".join(lines)
