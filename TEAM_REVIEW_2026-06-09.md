# dn38_solver — Comprehensive Team Review Synthesis

Generated 2026-06-09 by a Fable multi-agent team (5 dimension reviewers → adversarial verification → synthesis), grounded against a live environment: Python 3.11.9, 94 tests passing, Excel 16.0 reachable via COM, and a live end-to-end solve of the Azimuth sample model.

Every finding tagged **CONFIRMED** was independently re-verified against the source. Claims the verifier refuted (positional zip-pairing misattribution; a "guaranteed BADPARAMCOUNT crash" escalation) were excluded.

---

## 1. How it connects & works

dn38_solver is a Python harness that drives the DevEngine pricing model's VBA convergence solver headlessly, without you touching Excel. It is a **single-engine multiplexer that mirrors the model's own design**: just as the workbook runs one live project at a time via `$F$2`, the solver opens one fresh Excel session, cycles `F2` through each toggled-on project, and snapshots results per column.

Flow for `python -m dn38_solver.cli "<model>.xlsm" --chunked`:

1. **Preflight (no Excel).** `solver/orchestrator.py:solve_all` runs `shadow/preflight.py` — cheap zip-level reads checking the embedded macro against the repo's `SolveHeadless.bas` (D15 names, D16 stale modules, D17 SHA hash, D18 function arities) plus workbook calc settings (A1: `iterateDelta` ≤ 1E-4 so the sticky-IF circulars in rows 31/37 settle). Errors block; `--auto-fix` writes a non-destructive `_FIXED.xlsm` sibling with the current macro re-imported and `iterateDelta` patched, source untouched.
2. **Shadow read (openpyxl).** `shadow/reader.py` enumerates active projects from Project Inputs row-7 toggles (cols H onward), names from row 4 — read-only, no COM.
3. **Solve (COM).** `com/direct_runner.py` copies the workbook to `%TEMP%`, launches a fresh isolated Excel (`DispatchEx`), sets macro security low, opens the temp copy, calls VBA via `Application.Run`. All eight entry-point names/arities declared once in `com/vba_contract.py`. Two modes: single-shot (`SolveHeadless`) or chunked (`InitSolveEnvHL` → `SolveOneProjectByColHL` ×N → `FinalizeSolveEnvHL`). A watchdog thread resolves Excel's window handle to a PID and hard-kills it on per-call timeout.
4. **VBA engine.** Per project `SolveHeadless.bas` sets `F2`, runs Min Equity GoalSeek (PT Returns DSCR), HoldCo toggle, then the alternating NPP (row 38 vs F37/F36 IRR gap) / Appraisal Dev Fee (row 32 vs F31/F30) inner loop, with a 3-tier recalc ladder and warm→cold escalation. Each project writes a telemetry row to hidden `__SolverResults` (cols A–T).
5. **Read-back & stamp.** Python bulk-reads `__SolverResults`, then per project switches `F2` and runs `StampActiveProjectColumnHL` — hard-stamps rows 32/33/38/39, locks the converged DSCR onto row 371 (replacing the live PT!F129 link so projects can't bleed), leaving rows 31/37 as sticky-IF settled by a full recalc.
6. **Output.** SaveAs `<name>_SOLVED.xlsm`, error scan, Excel quit, temp cleanup. Results + in-flight checkpoints land in SQLite (`results.db`). Parallel mode (`--workers N`) round-robins across worker Excel instances and merges stamped columns back into one master, with a post-merge verifier.

```
 CLI (cli.py)
   |
   v
 solve_all (solver/orchestrator.py)
   |-- Phase 0  run_preflight (shadow/preflight.py)        [zip reads, no Excel]
   |              D15/D16/D17/D18 macro drift + A1 calcPr
   |              --auto-fix -> _FIXED.xlsm sibling (macro re-import via COM subprocess)
   |-- Phase 1  WorkbookReader (shadow/reader.py, openpyxl) [row-7 toggles -> projects]
   |-- Phase 2  run_direct (com/direct_runner.py)
   |              copy -> %TEMP% ; DispatchEx fresh Excel ; EnableEvents=False ; Open
   |              Application.Run            (contract: com/vba_contract.py)
   |                single-shot: SolveHeadless          (whole portfolio, 1 call)
   |                chunked:     InitSolveEnvHL -> SolveOneProjectByColHL xN -> Finalize
   |              [watchdog: Hwnd->PID, hard-kill Excel on per-call timeout]
   |              VBA per project: F2=k -> MinEquity GS -> HoldCo -> NPP<->Appraisal
   |                               telemetry -> __SolverResults (A-T)
   |              read-back: bulk A2:T + SwitchProjectAndRecalc + StampActiveProjectColumnHL
   |                         (stamps 32/33/38/39 + row-371 DSCR lock; 31/37 stay sticky-IF)
   |              SaveAs *_SOLVED.xlsm -> Quit -> cleanup
   |-- Phase 3  (workers>1) parallel_runner -> merge/ -> verify_merged_file
   '-- Phase 4  SQLite results.db + checkpoints + status JSON (storage/database.py)
```

---

## 2. Does it actually solve the model?

**Yes — but not the Azimuth sample as-is.**

- Environment fully live; 94-test suite passes.
- Preflight `--diagnose` on Azimuth fails with 3 errors + 1 warning, all rooted in one fact: the macro embedded in that workbook is an **older generation than the repo's `SolveHeadless.bas`**:
  - **D18 (error):** `ClassifyConvergenceHL` is 3 args in the workbook vs 4 in repo. The 4th arg makes convergence classification use the workbook's actual min-equity target instead of a hardcoded 10% — the 3-arg version would mis-tier convergence on non-10% models.
  - **D17 (error):** embedded macro SHA ≠ current `SolveHeadless.bas` SHA.
  - **A1 (error):** saved `iterateDelta` unset, so iterative calc could exit before rows-31/37 sticky-IF circulars settle, capturing stale IRRs.
  - **D16 (warn):** 3 stale modules (`Module2_Optimized`, `Module3`, `Module4`).
- **With `--auto-fix`, the pipeline works as designed:** writes `_FIXED.xlsm`, re-imports the macro, patches `iterateDelta`, re-verifies, solves. The live Azimuth run confirmed exactly this on an 18-project portfolio. The gate is doing its job — catching drift in ~10ms of zip reads instead of after a 30+ minute solve.

**Caveat (C6):** the 4-arg `ClassifyConvergenceHL` that preflight treats as canonical exists only as an **uncommitted local edit** to `SolveHeadless.bas`. The drift verdict on Azimuth is judged against a file no other checkout can reproduce. The system is correct; the reference is unversioned.

---

## 3. Confirmed bugs & risks

All **CONFIRMED** by adversarial verification.

### HIGH — can produce silent wrong numbers or lose a whole run

**C1. Auto-recovery resume reports pre-failure projects as converged with stale pre-solve values.** `com/direct_runner.py:674-708, 793-805, 898-911`. On a recoverable mid-portfolio COM error the workbook is closed `SaveChanges=False`, discarding the in-memory converged state of every project solved before the failure; the read pass then stamps **pre-solve values** into those columns, labels them "converged", ships a green `_SOLVED.xlsm`, and deletes the checkpoints holding the only correct numbers. **Silent wrong NPP on a green run.** *Fix:* `wb.Save()` the temp copy before the recovery close; or mark pre-failure projects `recovered_stale` and never clear checkpoints on recovered runs.

**C2. One stamp failure in the read pass throws away the entire run's results.** `com/direct_runner.py:898-911` vs catch-all `:1089-1096`. The per-project `StampActiveProjectColumnHL` has no local handler; the first exception propagates to the catch-all, which returns empty `project_results` and skips SaveAs — vaporizing every already-converged result in the open workbook. Matches the documented SolarStone 2026-06-04 incident and contradicts the module's own "always salvage partial results" note at `:765-771`. *Fix:* per-project try/except — record `stamp_failed`, keep the loop, keep the SaveAs.

**C3. Default (non-chunked) CLI run hard-kills Excel on any portfolio over ~4 projects.** `com/direct_runner.py:55, 1148-1157`. The single `SolveHeadless` call solving the whole portfolio is governed by the 600s **per-call** watchdog sized for one project. `--chunked` is opt-in, so the default path on a 5+ project book gets Excel killed mid-solve; every project lands `not_attempted`, no `_SOLVED.xlsm`. *Fix:* make chunked the default for >1 project, or scale the single-shot cap by project count.

**C4. Preflight is blind to two Subs Python actually calls.** `shadow/preflight.py:171-184`. `REQUIRED_MACRO_FUNCTIONS` omits `SwitchProjectAndRecalc` and `StampConvergedValuesHL` (the 6-arg merge stamp whose docstring warns a drift "would silently merge dev_fee values into the FMV row"). An arity-drifted merge stamp passes preflight and detonates at merge time after all workers paid full solve cost. *Fix:* derive the list from `vba_contract.ALL_PUBLIC_SUBS`; add a coverage test.

**C5. The Python↔VBA contract's `args` field is advisory — nothing enforces it.** `com/vba_contract.py:32-40`. No test compares `VBASub.args` arity to the .bas, so contract drift fails only at runtime as `DISP_E_BADPARAMCOUNT`. The parsing infra (`preflight._parse_param_counts`) already exists; only 2 of 8 entry points have signature tests. *Fix:* one unit test parsing the .bas and asserting arity for all `ALL_PUBLIC_SUBS` — turns contract drift into a CI failure.

**C6. `SolveHeadless.bas` — the canonical reference for all drift checks — is uncommitted.** `git status`: `M SolveHeadless.bas` plus untracked logs and `vba_source_solarstone/`. The 4-arg `ClassifyConvergenceHL`, the D17 hash, and stamps planted into newly imported workbooks all exist only in this working tree. A `git stash` silently changes what preflight considers "correct". *Fix:* commit the .bas (and `export_xlsx.py`) now; add a dirty-tree guard to `import_vba_module.py`.

**C7. `test_multiproject.py` violates the repo's own #1 safety rule against the live production workbook.** `test_multiproject.py:30-44`: openpyxl-load → mutate row-7 toggles → **openpyxl-save the .xlsm in place** → `solve_all` — the verified root cause of the RP Puma 0x800a9c68 incident (TROUBLESHOOTING.md Safety Rule #1), executed against `DEFAULT_WORKBOOK` (the live Box model). *Fix:* route the toggle edit through `com_edit.edit_xlsm` or operate on a temp copy.

### MEDIUM — wrong-but-plausible outputs, lost runs, blind spots

- **C8.** Watchdog (600s) undercuts VBA's own per-project budget (1200s) — `direct_runner.py:55` vs `SolveHeadless.bas:40`. A legitimately slow cold project VBA would finish gets Excel hard-killed, taking the session's read pass down. *Fix:* Python cap ≥ VBA cap + margin.
- **C9.** Watchdog never raises `TimeoutError`; with no proc handle it never times out — `direct_runner.py:183-226`. Kills surface as generic "RPC server unavailable" pointing at the wrong knob; `except TimeoutError` branches are dead code. *Fix:* shared `timed_out` event re-raised as `TimeoutError`.
- **C10.** Watchdog kill race: a call completing at the deadline gets its Excel killed post-success — `:146-180`, no `done.is_set()` re-check. *Fix:* re-check before kill/taskkill.
- **C11.** `Workbooks.Open` is unguarded and runs before the kill handle exists — `:536-557`. *Fix:* capture `excel_proc` right after `DispatchEx`; wrap Open at ~120–180s.
- **C12.** `--dry-run` is blocked by macro-drift errors irrelevant to it — `orchestrator.py:369-383` vs `:444`. Can't even list projects without `--auto-fix`. *Fix:* hoist the dry-run branch above the gate.
- **C13.** Layout constants and the A–T schema are quad-maintained (config.py, preflight.py, export_xlsx.py, .bas) with no cross-check. A one-row template insert silently mis-keys NPP/DevFee. *Fix:* CI test regex-parsing the .bas `Private Const` block vs `config.py`; schema-version cell written by `InitSolveEnvHL`.
- **C14.** Row 371 (the DSCR lock preventing cross-project bleed) is hardcoded in the .bas and never validated. *Fix:* preflight anchor check + pre-stamp assertion the target references `PT Returns!$F$129`.
- **C15.** No guard for Hybrid/Transfer TE structures with multiple projects toggled (your model map §5 flags H-anchored references). Converged-but-wrong, undetectable at runtime. *Fix:* preflight check reading the TE dropdown (PI F767/F768) vs toggle count.
- **C16.** Parallel merged file carries stale sticky-IF rows 31/37 (Live IRRs) for non-master columns, unflagged. An operator pasting Live IRR into an IC summary without cycling F2 gets pre-solve numbers. *Fix:* post-merge log naming stale columns + extend `verify_merged_file`.
- **C17.** Merge-audit telemetry rows don't match the A–T schema they mirror — `SolveHeadless.bas:731-749`. `analyze_solver_results.py` ingests them as phantom projects. *Fix:* mirror the real schema or make col A non-parseable.
- **C18.** Legacy single-shot VBA path lacks the chunked path's hardening — one broken project aborts the whole portfolio, and it's the **default** CLI behavior (C3). *Fix:* backport hardening or deprecate single-shot as a Python entry.
- **C19.** The `_SOLVED.xlsm` output ships with `MaxChange=0.001` — 10× above the ceiling preflight A1 enforces on inputs. Re-running the tool on its own output trips A1. *Fix:* finalize with `MaxChange=0.0001`.
- **C20.** Primary parallel merge path is an openpyxl save on .xlsm — the pattern the codebase elsewhere bans. *Fix:* prefer COM SaveAs as primary or launder through one Excel open/SaveAs.
- **C21.** `solve_all`'s own `timeout_sec=600` default hard-kills parallel workers at ~11 min for any programmatic caller. *Fix:* default 3600; split report-threshold from kill-deadline.
- **C22.** Macro-drift trust chain soft spots: (a) D17 hashes the repo .bas, never `vbaProject.bin` — a VBE edit keeps a "valid" stamp; (b) D18 silently skips when VBA extraction fails — "verified" and "unverifiable" indistinguishable; (c) D15 is a raw substring scan that can false-pass on names in the VBA compile cache. *Fix:* content-bound hash; warning on extraction failure; check against extracted live sources.
- **C23.** `--auto-import-macro` mutates the canonical source workbook in place, no backup, on Box-synced paths. *Fix:* `.pre-import.bak` before SaveAs; deprecate in favor of the `_FIXED` sibling.
- **C24.** mypy strict is configured but not installed — the typed boundary contract never runs. *Fix:* `pip install -e .[dev]`, triage, gate in CI.
- **C25.** The two real-workbook incident-regression tests silently skip — fixtures (hardcoded Desktop paths) gone. *Fix:* `DN38_FIXTURES` env var + committed synthetic .xlsm fixtures.
- **C26.** `keep_vba=True` on read-only loads leaks openpyxl's in-memory zip — root cause of the test-suite `ZipFile.__del__` warning. `post_merge.py:155`, `merge/__init__.py:140`, `preflight.py:1631` pay a full archive copy for nothing. *Fix:* drop `keep_vba` from never-saved loads; then set `filterwarnings = ["error::pytest.PytestUnraisableExceptionWarning"]`.

---

## 4. Comprehensive enhancements (prioritized, deduped)

### Correctness — P0
1. **Fix C1 + C2 in `direct_runner.py`** (save-before-recovery-close; per-project stamp try/except). Only paths to silent wrong NPP on a green run.
2. **Make `--chunked` the default for >1 project** and align the timeout family: Python per-call ≥ VBA `PROJECT_TIMEOUT_SECONDS` (1200) + margin; `solve_all` default 3600; split "report threshold" from "kill deadline" (C3, C8, C21).
3. **Runtime contract enforcement:** derive `REQUIRED_MACRO_FUNCTIONS` from `ALL_PUBLIC_SUBS`; unit test .bas signatures vs `vba_contract.args` for all 8 subs; unit test .bas `Private Const` block vs `config.py`; `__SolverResults` schema-version cell hard-checked on read (C4, C5, C13).
4. **Commit `SolveHeadless.bas`; dirty-tree guard on hash stamping** (C6). Cheapest highest-blast-radius fix.

### Correctness — P1
5. Preflight additions: row-371 DSCR anchor check + pre-stamp F129 assertion (C14); Hybrid/Transfer-TE × multi-project guard (C15); D18 "unverifiable" finding on extraction failure; D15 live-source check; D17b content-bound hash (C22).
6. Finalize with `MaxChange=0.0001` so `_SOLVED` files don't ship A1-violating (C19).
7. Fix the merge-audit row schema (C17); fix the `vba_contract` docstring for `StampActiveProjectColumnHL` (rows 31/37 are not stamped) and the orchestrator's ±1pp relaxed-band docstring (.bas says ±0.5pp).
8. Harden or deprecate the single-shot VBA path (C18).

### Reliability / Recovery — P1
9. Watchdog: `timed_out` event → real `TimeoutError` (C9); `done.is_set()` re-check before kill (C10); capture `excel_proc` after `DispatchEx` and watchdog-wrap `Workbooks.Open` (C11).
10. Backup-before-SaveAs (or Box-path refusal without `--force`) on `--auto-import-macro`; plan deprecation (C23).
11. Stale-module cleanup in `import_vba_module.py` (share `STALE_MACRO_MODULES` with preflight) so D16 stops firing forever.
12. Defensive `psutil` name() filter in `cleanup.py`; structured HRESULT decoding from the live com_error object instead of regexing `str()`.
13. Re-verify `iterateDelta` survived the COM SaveAs round-trip when A1 and D-tier auto-fixes combine.

### Performance — P2
14. Phase-specific recalc is **done** (3-tier ladder) — refresh `ENGINEERING_REVIEW.md`'s stale status block (still marks chunked execution "not yet implemented").
15. Drop `keep_vba` from read-only loads (C26) — removes a full ~13MB in-memory archive copy per preflight/merge/verify load.
16. Evaluate `USE_TIGHT_NPP_SCOPE=True` against the cold fixture once the regression harness exists.

### Testing / CI — P0/P1
17. **P0: CI pipeline.** ~30 lines of GitHub Actions: ubuntu+windows, `pip install -e .[dev]`, `pytest -W error::pytest.PytestUnraisableExceptionWarning`, `mypy dn38_solver`, `ruff check`. The suite was built for CI that was never wired up (C24).
18. **P1: warm/cold/stress Excel regression harness** (ENGINEERING_REVIEW Sprint 1's last open item): nightly Windows+Excel tier — warm P95 speed gate, cold 100%-strict-convergence gate, stress (8+ projects, `--workers 2-4`, 3 clean runs, `validate_parallel` diff == 0). Codify as `docs/RELEASE_GATES.md` + `run_release_gate.ps1`, triggered by any commit touching `SolveHeadless.bas` or `vba_contract.py`.
19. **P1:** `DN38_FIXTURES` + committed synthetic fixtures so incident-regression tests run again (C25); rewrite `test_multiproject.py` through `com_edit.edit_xlsm` (C7).
20. **P2:** unit tests for untested pure-Python surface — `storage/database.py` checkpoint round-trip (zero tests today), merge logic, `reader.extract_active_projects`, `compare_runs.pair_runs`. Convert source-text-assertion tests into tests of extracted pure functions.

### UX / Operability — P1/P2
21. **P1:** hoist `--dry-run` above the preflight gate (C12) — one-line reorder.
22. **P1:** post-merge stale-31/37 warning naming columns to cycle F2 through (C16).
23. **P2:** persist `__SolverResults` per-phase timings + convergence tier into SQLite; `--trend` report vs trailing-N baseline; heartbeat-staleness flag.
24. **P2:** fix `_show_history`'s `0.0`-as-missing truthiness; document `DN38_PER_CALL_TIMEOUT_SEC` in CLI help; move `results.db` to `%LOCALAPPDATA%` (WAL + OneDrive sync is a known "database is locked" source); env-var override for `DEFAULT_WORKBOOK`.

### Model-architecture — P2
25. Toggle-parse parity warning (flag `__SolverResults` offsets Python never enumerated); checkpoint keying by column offset not project name; line-order test pinning the `EnableEvents=False`-before-`Workbooks.Open` security invariant.

---

## 5. Top 5 highest-leverage next actions

1. **Patch the two silent-wrong-results bugs in `com/direct_runner.py`** (C1 + C2). Until these land, a green run is not proof of correct numbers — disqualifying for a bid-pricing tool.
2. **Commit `SolveHeadless.bas` and stand up CI** (pytest with unraisable-warnings-as-errors, mypy, ruff) **plus the three contract-parity unit tests**. One afternoon; converts drift-management from "judged against an unversioned file, enforced by memory" to mechanically gated.
3. **Fix the timeout family:** chunked by default for >1 project, per-call cap ≥ 1300s, `solve_all` default 3600, watchdog `done` re-check + real `TimeoutError`. Eliminates the default-path whole-portfolio kill (C3), the 600-vs-1200 squeeze (C8), and the misleading diagnostics (C9/C10) in one pass.
4. **Close the preflight blind spots and the dry-run gate:** add the two missing Subs, the D18-unverifiable finding, the Hybrid/TE multi-project guard, the row-371 anchor check; hoist `--dry-run` above the error gate. All zip-level, no COM, all protective of the converged-but-wrong failure class.
5. **Build the warm/cold/stress Excel regression harness with committed synthetic fixtures.** The only thing that exercises the riskiest 60% of the codebase (chunked loop, watchdog, recovery, parallel merge) automatically, and the release gate that makes shipping a new `SolveHeadless.bas` a checklist instead of a judgment call.
