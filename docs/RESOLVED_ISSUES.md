# Resolved Issues — Historical Record

This file preserves the issue archaeology that used to live in the main
TROUBLESHOOTING.md. The fixes are all in code; this exists for context
when investigating regressions or onboarding to the macro's quirks.

For day-to-day troubleshooting, see `TROUBLESHOOTING.md`.

---

## Issue #1: `Unable to set the Calculation property`
- **Cause:** `DispatchEx` creates a separate process; `Application.Calculation` conflicts with other Excel instances.
- **Fix:** VBA macro sets `xlCalculationManual` from inside its own process. Python no longer attempts to set it.

## Issue #2-3: Extreme slowness / GoalSeek freezes
- **Cause:** Without `xlCalculationManual`, each `Sheet.Calculate()` triggers full 885K formula recalc.
- **Fix:** SolveHeadless sets `xlCalculationManual` in-process, uses targeted `CalcModelCore` (13 sheets, ~650K formulas).

## Issue #4: VBA MsgBox freezes headless execution
- **Cause:** Original macro had 3 MsgBox calls (confirm, summary, error) that freeze in headless COM.
- **Fix:** `SolveHeadless.bas` — complete headless copy of the macro with all dialogs removed.

## Issue #5: Post-solve result reading triggers full-workbook recalcs
- **Cause:** Three compounding recalc bombs after macro: (1) setting calc back to Automatic, (2) CalculationState wait loop, (3) per-project `Sheets.Calculate()` for F2 switching.
- **Fix:** SolveHeadless leaves calc in Manual. `SwitchProjectAndRecalc` does targeted 13-sheet recalc. Calc restored to Automatic only at workbook close.

## Issue #6: GoalSeek over-iteration on pre-solved workbooks
- **Cause:** `MaxIterations=1000` and `MAX_GS_RETRY=6` were overkill for near-optimal starting values.
- **Fix:** Tuned to `MaxIterations=200`, `MAX_GS_RETRY=3` in warm mode. Macro time dropped from 69s to 27s per project on pre-solved inputs. Cold-start mode preserves the higher values.

## Issue #7: CalcModelCore correctness vs performance tradeoff

The primary blocker for production use on unsolved workbooks. The VBA
`CalcModelCoreHL()` sub determines both convergence reliability and
execution speed. Three approaches were tested:

| Implementation                           | Speed (per project) | Cold-start convergence    | Outcome                |
|------------------------------------------|---------------------|---------------------------|------------------------|
| `Application.Calculate` (dirty-cell)     | ~27–36s             | **Fails** — 7/8 wrong     | Fast but unreliable    |
| `Sheets("X").Calculate` × 13             | ~125s               | Correct                   | COM RPC timeout ~1000s |
| `Application.CalculateFull`              | Mid                 | Correct                   | Adopted as fallback    |

**Why `Application.Calculate` alone failed:** it recalculates only cells
flagged dirty by Excel's dependency tracker. After GoalSeek modifies a
cell, direct dependents get flagged but indirect dependents through
volatile `OFFSET()` chains (heavily used in Project Inputs col F → tabs)
may not be. Subsequent GoalSeek calls then see stale intermediate values
and converge to wrong solutions.

**Resolution:** deterministic recalc ladder inside `CalcModelCoreHL`:
1. `Application.Calculate` (fast path)
2. Double `Application.Calculate` (catches one extra dependency layer)
3. `Application.CalculateFull` (correctness guardrail)

Plus per-project warm/cold mode: each project starts in warm mode
(`MAX_GS_RETRY=3`, `MaxIterations=200`) and promotes to cold mode
(`MAX_GS_RETRY=6`, `MaxIterations=1000`) only if tolerance checks fail.

Plus chunked execution (`InitSolveEnvHL` / `SolveOneProjectByColHL` /
`FinalizeSolveEnvHL`) so no single COM call exceeds ~120s, well under the
RPC timeout.

Validated cold-start end-to-end on SMP WalkTEST 2026-05-12: 6 projects,
1366s macro, 5/6 strict + 1/6 relaxed convergence, zero Tier-3
escalations, no COM RPC failures.

### Cold-start convergence evidence (Lightstar, `Application.Calculate` only)

| Project    | NPP (seed → solved) | IRR Gap  | Equity % | Converged? |
|------------|---------------------|----------|----------|------------|
| IL VER001  | 0.2 → 0.02          | 0.633    | -7.26%   | No         |
| IL AUX001  | 0.2 → 0.56          | 0.371    | -1.49%   | No         |
| NY LAK001  | 0.2 → -0.21         | 0.014    | 7.58%    | No         |
| NY GVT001  | 0.2 → 0.21          | 0.001    | 9.81%    | Almost     |
| NY GVT002  | 0.2 → -0.04         | 0.016    | 4.90%    | No         |
| NY BLG005  | 0.2 → 0.19          | 0.112    | 0.99%    | No         |
| NY MTG030  | 0.2 → 0.05          | 0.008    | 8.42%    | No         |
| MD FRE001  | 0.2 → 1.82          | <0.001   | 10.34%   | **Yes**    |

NPP values DID change (GoalSeek ran), but the stale intermediate values
caused convergence to wrong solutions. This is the data that drove the
recalc-ladder design.

---

## Issue #8: Per-project DSCR collision in multi-project runs

- **Cause:** `PT Returns!F129` reflects only the active project; reading it after the solve loop returned the last project's DSCR for every row.
- **Fix:** Macro writes DSCR per project into hidden `__SolverResults` immediately after each project converges. Python runner reads from that sheet instead of `F129`.

---

## Performance history

| Version | CalcModelCore             | Workbook state | Projects | Macro time | Per-project | Converged? |
|---------|---------------------------|----------------|----------|------------|-------------|------------|
| v1      | N/A (subprocess)          | Pre-solved     | 1        | 252s       | 252s        | Yes        |
| v3      | `Application.Calculate`   | Pre-solved     | 1        | 72s        | 72s         | Yes        |
| v4b     | `Application.Calculate`   | Pre-solved     | 1        | 27s        | 27s         | Yes        |
| v4b     | `Application.Calculate`   | Pre-solved     | 5        | 108s       | 22s         | Yes        |
| v4      | `Application.Calculate`   | Unsolved       | 8        | 898s       | 112s        | **1/8**    |
| v5      | `Sheets.Calculate` × 13   | Unsolved       | 8        | 1271s      | —           | **RPC fail** |
| v6      | `Sheets.Calculate` × 13   | Unsolved       | 8        | 997s       | —           | **RPC fail** |
| v7      | Recalc ladder + chunked   | Unsolved       | 6        | 1366s      | 228s        | 5 strict + 1 relaxed |
