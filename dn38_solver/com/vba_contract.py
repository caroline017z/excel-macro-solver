"""dn38_solver.com.vba_contract — The Python<->VBA boundary, declared once.

Every Public Sub in `SolveHeadless.bas` that Python invokes via
`Application.Run` is registered here as a `VBASub` constant. Callers
reference these constants instead of hardcoding string literals like
`f"'{wb.Name}'!SolveOneProjectByColHL"` — so a future contributor who
renames the .bas Sub or reorders its arguments breaks an import in
this module rather than failing silently inside a swallowed COM
exception at runtime.

Why this exists
---------------
Architectural review (round 3) flagged the Python<->VBA boundary as
the only real bus-factor risk in the system. Six Public Subs are called
by name + position; the merge fallback path catches `Application.Run`
exceptions inside a `with contextlib.suppress` style, so a parameter
drift in the .bas turns into a per-project warning that's easy to
miss. Centralizing the names + arities + arg specs here turns "did
someone rename the Sub?" into a Python-level NameError instead of an
Excel-level COM error.

Calling convention
------------------
`vba_call_str(wb_name, sub)` builds the `'{wb_name}'!{sub.name}` string
that `Application.Run` accepts. Use it instead of inline f-strings.

When you change SolveHeadless.bas
---------------------------------
1. If you rename a Sub, update `name` here and grep for usages — the
   `VBASub` constants live in this module, so import-only callers
   discover the change at import time.
2. If you change a Sub's signature, update the `args` field. The args
   field is currently advisory (Python doesn't enforce it across
   `Application.Run`), but it serves as documentation and could be
   wired into a runtime check via VBProject inspection.
3. The contract is tested implicitly every solve: the SolveHeadless
   Public Sub list is what the parallel runner verifies via
   `cm.Find("StampConvergedValuesHL", ...)` before the merge fallback.
   Add new entries here AND update import_vba_module.py's verifier if
   you add a new must-exist Sub.
"""
from __future__ import annotations

from typing import NamedTuple


class VBASub(NamedTuple):
    """A Public Sub in SolveHeadless.bas that Python calls via COM.

    `name`: the exact Public Sub identifier as declared in the .bas.
    `args`: list of (arg_name, type_hint) tuples. Advisory — used for
        documentation and for future runtime-introspection checks. The
        type_hint is the VBA type ("Integer", "Double", "Boolean")
        because that's what the Sub signature in the .bas declares;
        Python callers convert before passing.
    """
    name: str
    args: tuple[tuple[str, str], ...] = ()


# ---- Setup / lifecycle entry points -------------------------------------

# Disable output-sheet recalc inside the macro. Python sets this once
# at run start to skip the Dashboard / Portfolio / Waterfall recalc on
# big workbooks where those sheets aren't load-bearing.
SET_SKIP_OUTPUT_RECALC = VBASub(
    name="SetSkipOutputRecalcHL",
    args=(("bSkip", "Boolean"),),
)

# Switch the active project pointer (Project Inputs!F2) to a different
# project offset and run the targeted recalc ladder. Used by Python
# during the post-solve cell-read pass to pull per-project values
# without triggering a full-workbook recalc.
SWITCH_PROJECT_AND_RECALC = VBASub(
    name="SwitchProjectAndRecalc",
    args=(("projOffset", "Integer"),),
)

# Snapshot the active project's per-column convergence cells (rows 31,
# 32, 33, 37, 38, 39 on Project Inputs) as hard constants. Called by
# Python from the post-solve read pass, AFTER SwitchProjectAndRecalc has
# pinned F2 and refreshed the workbook. See the .bas Sub docstring for
# why the stamp lives here rather than inside SolveOneProjectByColHL.
STAMP_ACTIVE_PROJECT_COLUMN = VBASub(
    name="StampActiveProjectColumnHL",
    args=(
        ("colIdx",      "Integer"),
        # Per-project DSCR (meta["dscr"] from __SolverResults!C). Required
        # — PT!F129 is a single live cell that GoalSeek overwrites per
        # project; without this restore the post-read CalculateFull
        # propagates the last-solved-project's DSCR through the equity
        # IRR chain and stamps the wrong IRR into row 37. Pass 0.0 only
        # if no DSCR was captured (placeholder / fast-skip projects).
        ("dscrRestore", "Double"),
    ),
)


# ---- Single-shot solve --------------------------------------------------

# The legacy single-shot entry point. Solves every project in one COM
# call. Risks RPC timeout on large portfolios — use the chunked entry
# points below for cold portfolios > ~10 projects.
SOLVE_HEADLESS = VBASub(name="SolveHeadless")


# ---- Chunked solve (one COM call per project) ---------------------------

# Per-project chunked path. Init opens the macro environment once,
# then SolveOneProjectByColHL is called per project (each its own COM
# round-trip so no single Application.Run can exceed the ~900s RPC
# timeout), then Finalize teardown.
INIT_SOLVE_ENV = VBASub(name="InitSolveEnvHL")

SOLVE_ONE_PROJECT_BY_COL = VBASub(
    name="SolveOneProjectByColHL",
    args=(
        ("colIdx",     "Integer"),
        ("projName",   "String"),
        ("resultsRow", "Integer"),
    ),
)

FINALIZE_SOLVE_ENV = VBASub(
    name="FinalizeSolveEnvHL",
    # Finalize takes the F2-restore offset as its only positional arg.
    args=(("originalProjOffset", "Integer"),),
)


# ---- Parallel-mode merge helper -----------------------------------------

# Called by the VBA-helper merge fallback in `dn38_solver.merge`. Stamps
# converged values from a peer worker's _SOLVED.xlsm into the master
# via Excel COM (used when openpyxl can't round-trip the macro project).
# Order of args MUST match the .bas Sub signature exactly — a swap here
# would silently merge dev_fee values into the FMV row.
STAMP_CONVERGED_VALUES = VBASub(
    name="StampConvergedValuesHL",
    args=(
        ("colIdx",    "Integer"),
        ("npp",       "Double"),
        ("devFee",    "Double"),
        ("fmv",       "Double"),
        ("liveIRR",   "Double"),
        ("apprLive",  "Double"),
        ("nppTotal",  "Double"),
    ),
)


# ---- Convenience helpers ------------------------------------------------

def vba_call_str(wb_name: str, sub: VBASub) -> str:
    """Build the `'{wb_name}'!{sub_name}` string Application.Run accepts.

    Single-quotes around wb_name are required when the workbook name
    contains spaces, dots, or hyphens (which most do). Excel parses the
    quoting; we just need to format it consistently.
    """
    return f"'{wb_name}'!{sub.name}"


# Tuple of every Sub Python relies on, used by parallel_runner's
# fallback-path precondition check (currently inlined as
# `cm.Find("StampConvergedValuesHL", ...)`). When extending the contract,
# add to this tuple AND wire the new Sub into the must-exist verification
# in `dn38_solver.merge.merge_via_vba_fallback`.
ALL_PUBLIC_SUBS: tuple[VBASub, ...] = (
    SET_SKIP_OUTPUT_RECALC,
    SWITCH_PROJECT_AND_RECALC,
    STAMP_ACTIVE_PROJECT_COLUMN,
    SOLVE_HEADLESS,
    INIT_SOLVE_ENV,
    SOLVE_ONE_PROJECT_BY_COL,
    FINALIZE_SOLVE_ENV,
    STAMP_CONVERGED_VALUES,
)
