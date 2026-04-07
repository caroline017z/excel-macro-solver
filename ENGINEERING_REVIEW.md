# Excel Macro Solver — Full Engineering Review (Speed + Reliability)

## Review Charter

This review consolidates recommendations from a virtual cross-functional team (VBA, Python/COM, model architecture, QA, and operations) to improve solve correctness and execution speed for both warm and cold-start workbooks.

---

## Implementation Status (Current)

### Completed
- Deterministic recalc ladder in VBA (`Calculate` → double `Calculate` → `CalculateFull`).
- Per-project warm/cold GoalSeek escalation.
- Hidden `__SolverResults` sheet with per-project telemetry + DSCR capture.
- Python runner ingestion of `__SolverResults`, DSCR override safety, and heartbeat surfacing.

### In Progress / Outstanding
- Chunked macro execution in Python runner (not yet implemented).
- True in-flight timeout cancellation of COM macro call (current timeout is post-run guard only).
- Regression harness with warm/cold/stress fixtures in a Windows+Excel test environment.

---

## 1) VBA / Calculation Engine Review

### Current Strengths
- `SolveHeadless.bas` now has a deterministic recalc ladder and warm/cold GoalSeek escalation.
- Non-core sheets are already disabled, which is the biggest controllable Excel-level optimization.

### High-Impact Next Changes
1. **Replace volatile dependency chains in model formulas** (especially `OFFSET`, `INDIRECT`, and deep volatile cascades) with non-volatile alternatives (`INDEX`, structured references, helper lookup tables).  
   - Why: volatility is the largest hidden multiplier behind recalc unpredictability and slowdowns.
   - Expected outcome: fewer forced full recalculations and faster Tier 1/Tier 2 solve loops.

2. **Split `CalcModelCoreHL` into phase-specific calc scopes**:
   - `CalcForDSCR`
   - `CalcForIRR`
   - `CalcForAppraisal`
   - Why: not every GoalSeek step needs all 13 core sheets.
   - Expected outcome: lower average time per retry without sacrificing deterministic behavior.

3. **Add strict convergence gates per solve stage** (DSCR gate, IRR gate, Appraisal gate) and capture which stage fails per project.
   - Why: improves targeted tuning and avoids over-escalating all stages for one failing stage.

---

## 2) Python COM Runner Review

### Current Strengths
- Single-process COM flow avoids subprocess overhead.
- Per-project reads via `SwitchProjectAndRecalc` reduce post-solve recalc bombs.

### High-Impact Next Changes
1. **Add chunked workbook execution mode** (N projects per macro invocation). **Status: pending**
   - Process:
     - Set toggle row for a subset of projects.
     - Run `SolveHeadless`.
     - Read/write results.
     - Continue with next subset.
   - Why: prevents long macro calls from approaching COM RPC failure windows.

2. **Add heartbeat + stall detection**: **Status: partial**
   - VBA writes `project_index`, `iteration`, `timestamp` to a known status cell/sheet.
   - Python polls periodically and can classify `stalled` vs `running`.
   - Why: better failure recovery and operator visibility during long portfolios.

3. **Add adaptive retry policy in Python**: **Status: pending**
   - On macro failure: reopen Excel and retry the failed chunk once in conservative mode (cold settings and Tier 3 enabled at start).
   - Why: salvages long runs without restarting entire portfolio.

---

## 3) Model Architecture Review

### High-Impact Structural Changes
1. **Pre-calculate immutable sheets once per workbook session** and cache assumptions where possible.
2. **Introduce a “project state fingerprint”** (seed values + known stable inputs) to classify projects as warm/cold before solve begins.
3. **Add a result integrity table inside workbook** (one row per project with NPP, Dev Fee, DSCR, IRR gap, Appraisal gap, equity %) to eliminate ambiguity from mutable single cells like `F129`.

---

## 4) Quality Engineering Review

### Recommended Test Matrix
- **Warm fixture** (pre-solved): target speed regression guard.
- **Cold fixture** (unsolved): target correctness guard.
- **Stress fixture** (mixed projects): target RPC resilience.

### Suggested Release Gates
1. 100% convergence for required projects under tolerance rules.
2. No invalid outputs (e.g., disallowed negative equity).
3. P95 per-project runtime threshold for warm and cold classes.
4. No COM RPC failures in 3 consecutive stress runs.

---

## 5) Operations / Observability Review

### Metrics to Log Per Project
- Recalc tier usage counts (Tier 1 / 2 / 3).
- GoalSeek retries consumed by stage.
- Time spent in each stage (DSCR, IRR, Appraisal).
- Final tolerance gaps and convergence status code.

### Dashboards / Alerts
- Alert when Tier 3 usage exceeds baseline.
- Alert when average retries rise above trailing 7-run baseline.
- Alert on missing heartbeat for >N seconds.

---

## Prioritized Implementation Roadmap

### Sprint 1 (Stability)
1. ✅ Add per-project status heartbeat in VBA and polling metadata in Python.
2. ✅ Add project result integrity table (including DSCR) and read it from Python.
3. ⏳ Add cold/warm/stress regression harness and publish baseline metrics.

### Sprint 2 (Throughput)
1. ⏳ Implement chunked macro execution in runner.
2. ⏳ Implement phase-specific recalculation scopes.
3. ⏳ Tune escalation thresholds from real telemetry.

### Sprint 3 (Structural Speedups)
1. ⏳ Refactor volatile formula hotspots in workbook model.
2. ⏳ Add project fingerprinting and predictive mode selection.
3. ⏳ Harden retry/recovery workflow for unattended batch runs.

---

## Codebase Audit Findings (Structure + Efficiency)

The following concrete improvements were identified from a full-pass code review of orchestrator, CLI, reader, runner, and VBA integration points:

1. **Correct numeric truthiness bugs in summaries/parsing**  
   - Use explicit `is not None` checks for metrics (NPP/Dev Fee/FM V/DSCR/equity%) so valid zero values are not treated as missing.  
   - This was already applied in `orchestrator.py` for `equity_pct` and summary formatting.

2. **Keep interface language aligned with architecture**  
   - CLI timeout help text should refer to solver macro timeout threshold (not legacy subprocess wording).  
   - This was already updated in `cli.py`.

3. **Normalize telemetry contract across layers**  
   - Standardize required vs optional fields for `__SolverResults` parsing and propagate typed status values to DB/reporting.

4. **Reduce coupling by introducing a solver result adapter**  
   - Add one translation layer from raw COM/VBA output -> stable internal model.  
   - Prevents future VBA telemetry shape changes from rippling through orchestrator/reporting/storage modules.

5. **Add deterministic Windows integration test harness**  
   - Split tests into:
     - pure-Python unit tests (cross-platform)
     - Windows+Excel integration fixtures (gated/optional in CI)
   - This closes current validation gaps where Linux environments cannot execute end-to-end solve flows.
