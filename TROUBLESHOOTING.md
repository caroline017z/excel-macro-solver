# 38DN Hybrid Shadow Solver — Troubleshooting & Engineering Analysis

## Project Overview

Python automation tool that calls VBA macros in Excel pricing models (885K formulas, 23 sheets, 60 project columns) to solve NPP, Dev Fee, and DSCR via iterative GoalSeek. Target: ~30s per project, matching manual Excel button speed.

---

## Architecture

```
CLI (cli.py)
  → Orchestrator (orchestrator.py)
      → Shadow Reader (openpyxl) — pre-reads workbook, extracts active projects
      → Direct COM Runner (direct_runner.py)
          → Opens temp copy of workbook via win32com.client.Dispatch
          → Runs SolveHeadless VBA macro via Application.Run
          → Reads per-project results via SwitchProjectAndRecalc
          → Writes status JSON for Streamlit tracker
      → SQLite persistence + branded .xlsx export
```

**Key VBA module:** `SolveHeadless.bas` — imported into each workbook via `import_vba_module.py`. Contains the full solve loop (identical to the original `SolveMinEquityWithHoldCo`) with all MsgBox dialogs removed, `DisableNonCoreSheets` optimization, and `SwitchProjectAndRecalc` helper for post-solve reads.

---

## Resolved Issues

### Issue #1: `Unable to set the Calculation property` (RESOLVED)
- **Cause:** `DispatchEx` creates a separate process; `Application.Calculation` conflicts with other Excel instances.
- **Fix:** VBA macro sets `xlCalculationManual` from inside its own process. Python no longer attempts to set it.

### Issue #2-3: Extreme slowness / GoalSeek freezes (RESOLVED)
- **Cause:** Without `xlCalculationManual`, each `Sheet.Calculate()` triggers full 885K formula recalc.
- **Fix:** SolveHeadless sets `xlCalculationManual` in-process, uses targeted `CalcModelCore` (13 sheets, ~650K formulas).

### Issue #4: VBA MsgBox freezes headless execution (RESOLVED)
- **Cause:** Original macro has 3 MsgBox calls (confirm, summary, error) that freeze in headless COM.
- **Fix:** `SolveHeadless.bas` — complete headless copy of the macro with all dialogs removed.

### Issue #5: Post-solve result reading triggers full-workbook recalcs (RESOLVED)
- **Cause:** Three compounding recalc bombs after macro: (1) setting calc back to Automatic, (2) CalculationState wait loop, (3) per-project `Sheets.Calculate()` for F2 switching.
- **Fix:** SolveHeadless leaves calc in Manual. `SwitchProjectAndRecalc` does targeted 13-sheet recalc. Calc restored to Automatic only at workbook close.

### Issue #6: GoalSeek over-iteration on pre-solved workbooks (RESOLVED)
- **Cause:** `MaxIterations=1000` and `MAX_GS_RETRY=6` were overkill for near-optimal starting values.
- **Fix:** Tuned to `MaxIterations=200`, `MAX_GS_RETRY=3`. Macro time dropped from 69s to 27s per project.
- **Caveat:** These tuned values are insufficient for cold-start solves (see Issue #7).

---

## Open Issue #7: CalcModelCore Correctness vs Performance Tradeoff

This is the **primary unresolved blocker** for production use on unsolved workbooks.

### The Problem

The VBA `CalcModelCoreHL()` sub determines both convergence reliability and execution speed. Three implementations have been tested:

| CalcModelCore Implementation | Speed (per project) | Cold-Start Convergence | Outcome |
|---|---|---|---|
| `Application.Calculate` (dirty-cell only) | ~27-36s | **FAILS** — 7/8 projects wrong | Fast but unreliable |
| `Sheets("X").Calculate` × 13 (full sheet recalc) | ~125s | Correct (assumed) | COM RPC timeout at ~1000s |
| `Application.CalculateFull` | **Untested** | **Untested** | Most promising approach |

### Test Data

**Pre-solved workbook** (38DN-IL, 1 project, values near optimal):
- `Application.Calculate`: 27s, converged, NPP=$0.3111 ✓
- Consistent across 9 runs

**Unsolved workbook** (Lightstar, 8 projects, seed NPP=0.2, DevFee=1.0):
- `Application.Calculate`: 898s, **7/8 failed convergence**, IRR gaps up to 63%
- `Sheets("X").Calculate` × 13: 997-1270s, **COM RPC failure** (`-2147023170: The remote procedure call failed`)

### Root Cause Analysis

`Application.Calculate` recalculates only cells flagged as "dirty" by Excel's dependency tracker. After GoalSeek modifies a cell:
1. The changed cell and its **direct** dependents are flagged dirty
2. `Application.Calculate` recalculates those cells
3. **Indirect dependents** (especially through volatile `OFFSET()` chains) may NOT be flagged
4. Subsequent GoalSeek calls see **stale intermediate values** → wrong targets → no convergence

The original VBA macro uses `Sheets("X").Calculate` which forces ALL formulas on each sheet to recalculate — brute-force but correct. However, doing this 13 times per CalcModelCore call × ~128 calls per project × 8 projects = ~13,000 sheet-level recalculations, taking 1000+ seconds and exceeding COM's RPC timeout.

### Convergence Failure Evidence

From the Lightstar test with `Application.Calculate`:

| Project | NPP (seed→solved) | IRR Gap | Equity % | Converged? |
|---|---|---|---|---|
| IL VER001 | 0.2 → 0.02 | 0.633 | -7.26% | No |
| IL AUX001 | 0.2 → 0.56 | 0.371 | -1.49% | No |
| NY LAK001 | 0.2 → -0.21 | 0.014 | 7.58% | No |
| NY GVT001 | 0.2 → 0.21 | 0.001 | 9.81% | Almost |
| NY GVT002 | 0.2 → -0.04 | 0.016 | 4.90% | No |
| NY BLG005 | 0.2 → 0.19 | 0.112 | 0.99% | No |
| NY MTG030 | 0.2 → 0.05 | 0.008 | 8.42% | No |
| MD FRE001 | 0.2 → 1.82 | <0.001 | 10.34% | **Yes** |

NPP values DID change (GoalSeek ran), but the stale intermediate values caused it to converge to wrong solutions.

### Proposed Solutions (Priority Order)

**1. `Application.CalculateFull`** (single-line change, untested)
```vba
Private Sub CalcModelCoreHL()
    Application.CalculateFull
End Sub
```
- Forces full recalculation of ALL formulas in enabled sheets, including those not flagged dirty
- Single API call — Excel can parallelize internally
- Expected speed: between `Application.Calculate` (27s) and 13× sheet calcs (125s)
- Non-core sheets still disabled via `EnableCalculation=False`

**2. Hybrid first-call / subsequent-call approach** (moderate complexity)
```vba
Private m_bFirstCalc As Boolean  ' Module-level flag

Private Sub CalcModelCoreHL()
    If m_bFirstCalc Then
        Application.CalculateFull  ' First call per project: force full recalc
        m_bFirstCalc = False
    Else
        Application.Calculate      ' Subsequent calls: dirty-cell only (fast)
    End If
End Sub
```
- First CalcModelCore per project builds the full dependency state
- Subsequent calls within the same project only need dirty-cell tracking
- Resets `m_bFirstCalc = True` when F2 changes (new project)

**3. `Range.Dirty` on OFFSET-dependent ranges** (targeted, complex)
```vba
Private Sub CalcModelCoreHL()
    ' Force OFFSET-dependent ranges into the recalc queue
    Sheets("Project Inputs").Range("F30:F40").Dirty
    Sheets("PT Returns").Range("C128:C134").Dirty
    Sheets("PT Returns").Range("F128:F130").Dirty
    Application.Calculate
End Sub
```
- Manually flags known OFFSET-dependent cells as dirty before recalc
- Preserves the fast dirty-cell approach for everything else
- Requires knowing which ranges are OFFSET-dependent (fragile if model changes)

**4. Double-calculate** (simplest, slowest of the fast options)
```vba
Private Sub CalcModelCoreHL()
    Application.Calculate  ' Pass 1: recalc dirty cells, flag their dependents
    Application.Calculate  ' Pass 2: recalc newly-flagged dependents
End Sub
```
- Two passes catch indirect dependents missed by the first pass
- ~2x slower than single `Application.Calculate` but still faster than 13 sheet calcs
- May not catch all chains if dependency depth > 2

### Additional Constraints

- **COM RPC timeout:** Macro execution exceeding ~900-1000s causes `(-2147023170, 'The remote procedure call failed.')`. Any solution must keep total macro time under this.
- **GoalSeek parameters:** Cold-start solves need `MAX_GS_RETRY=6` and `MaxIterations=1000` for convergence. Pre-solved workbooks can use `MAX_GS_RETRY=3` and `MaxIterations=200`.
- **Per-project DSCR:** F129 is a GoalSeek value cell (not formula). After solving all projects, F129 reflects only the last project. Fix: write per-project DSCR to a scratch range in VBA before moving to next project.

---

## Performance History

| Version | CalcModelCore | Workbook State | Projects | Macro Time | Per-Project | Converged? |
|---|---|---|---|---|---|---|
| v1 (subprocess) | N/A (DispatchEx) | Pre-solved | 1 | 252s | 252s | Yes |
| v3 (Application.Calculate) | `Application.Calculate` | Pre-solved | 1 | 72s | 72s | Yes |
| v4b (tuned) | `Application.Calculate` | Pre-solved | 1 | 27s | 27s | Yes |
| v4b (tuned) | `Application.Calculate` | Pre-solved | 5 | 108s | 22s | Yes |
| v4 (Application.Calculate) | `Application.Calculate` | **Unsolved** | 8 | 898s | 112s | **1/8** |
| v5 (sheet-level) | `Sheets.Calculate` × 13 | **Unsolved** | 8 | 1271s | — | **RPC fail** |
| v6 (sheet-level) | `Sheets.Calculate` × 13 | **Unsolved** | 8 | 997s | — | **RPC fail** |

---

## File Inventory

| File | Purpose | Status |
|---|---|---|
| `SolveHeadless.bas` | VBA macro (headless, all optimizations) | Active — CalcModelCore needs fix |
| `import_vba_module.py` | Injects VBA into any workbook | Working |
| `dn38_solver/com/direct_runner.py` | Direct COM execution, status tracking | Working |
| `dn38_solver/com/launcher.py` | Subprocess launcher (legacy, replaced) | Deprecated |
| `com_worker.py` | Subprocess COM worker (legacy, replaced) | Deprecated |
| `dn38_solver/solver/orchestrator.py` | Main solve loop, result parsing | Working |
| `dn38_solver/solver/sequence.py` | GoalSeek task builder | Working |
| `dn38_solver/shadow/reader.py` | openpyxl workbook reader | Working |
| `dn38_solver/config.py` | Cell addresses, constants | Working |
| `dn38_solver/types.py` | msgspec Structs | Working |
| `dn38_solver/storage/database.py` | SQLite persistence | Working |
| `dn38_solver/reporting/export_xlsx.py` | Branded summary .xlsx | Working |
| `dn38_solver/dashboard/tracker.py` | Streamlit progress tracker | Framework only |
| `dn38_solver/cli.py` | CLI entry point | Working |
| `run_portfolio_test.py` | 8-project integration test | Working |
| `extract_vba.py` | Extracts VBA source from workbooks | Working |
| `vba_source/` | Extracted VBA from original workbook | Reference |
