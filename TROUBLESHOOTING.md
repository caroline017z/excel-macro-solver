# 38DN Hybrid Shadow Solver — Troubleshooting

Python automation that drives the SolveHeadless VBA macro in 38DN pricing
models (.xlsm) to converge NPP, Dev Fee, and DSCR via iterative GoalSeek.

The fastest path through this doc is to find your symptom in §1, then jump
to the linked section. Background sits at the bottom in §10.

---

## 1. Symptom index — start here

| What you're seeing                                                       | Jump to                                  |
|--------------------------------------------------------------------------|------------------------------------------|
| Generic COM `(-2147352567, 'Exception occurred.', ...)` mid-solve        | §2 Auto-recovery & COM decoding          |
| Pre-flight `D17` — macro hash mismatch                                   | §3 Pre-flight: macro version & drift     |
| Pre-flight `D15` — embedded macro missing functions                      | §3 Pre-flight: macro version & drift     |
| Pre-flight `A1`–`A4` — calc property issues                              | §4 Pre-flight: calc properties           |
| Pre-flight `B5`–`B9` — structure / cells / protection                    | §5 Pre-flight: structure                 |
| Pre-flight `C10`–`C12` — cached formula errors on critical cells         | §6 Pre-flight: critical-path errors      |
| Pre-flight `E13`/`E14` — Dev Fee or NPP out of bounds                    | §7 Pre-flight: input bounds              |
| Pre-flight `E15a` — RC sub-block has `Custom` toggle + empty Rate Name   | §7.1 Pre-flight: RC config audit         |
| Pre-flight `E15b` — RC source mode mismatched across active projects     | §7.1 Pre-flight: RC config audit         |
| Solver converged but Dev Fees look ~3-5× too high; DSCRs ~1.02x          | §7.1 Pre-flight: RC config audit         |
| `(-2147023174, 'RPC server unavailable')` — Excel died                   | §8 Excel process failures                |
| `(-2147417848, 'object disconnected')` — Excel exited mid-call           | §8 Excel process failures                |
| Solver runs to completion but projects show `not_converged`              | §9 Convergence quality                   |
| `Ship-ready: 0/N` even though projects show converged values             | §9 Convergence quality                   |
| Parallel run says `Merge path: copy_master` (ERROR)                      | §9 Convergence quality                   |
| Workbook works in Excel but solver hangs at "Opening workbook..."        | §8 Excel process failures                |
| You just want to vet a workbook before solving                           | `python -m dn38_solver.cli --diagnose`   |

---

## 2. Auto-recovery & COM error decoding

The solver decodes COM exceptions automatically and surfaces a human-readable
hint. When an HRESULT is flagged auto-recoverable, the runner closes the
workbook, re-imports the macro via Excel COM (`import_vba_module.py` in a
fresh subprocess), reopens, and retries the failed chunk once.

**Recoverable HRESULTs** (one-shot retry, no operator action needed):
- `0x800a9c68` (`-2146788248`) — generic VBA exception inside
  `SolveOneProjectByColHL`. Almost always stale workbook state caused by
  an upstream openpyxl save corrupting parts the macro depends on. Verified
  root cause of the 2026-05-13 RP Puma incident; Excel COM SaveAs in the
  re-import pass fully rewrites the file and clears the corruption.
- Other "secondary HRESULT in decoder table" entries logged with the
  `Auto-recovery: ELIGIBLE` tag.

**Non-recoverable HRESULTs** (operator must intervene):
- `0x800706BA` — RPC server unavailable. Excel process died (often the
  ~900s RPC timeout firing on a long single COM call). Retry with
  `--chunked` or `--workers N`.
- `0x80010108` — object disconnected. Excel exited mid-call. Close all
  Excel; rerun with fewer workers or `--no-output-recalc`.

**If recovery fails:**
```
auto-recovery did not converge — original: <hresult> | retry: <hresult>
```
The retry HRESULT is the meaningful one. If it's the same recoverable
HRESULT, the issue is in the source workbook itself rather than transient
state — start with `--diagnose` to surface what preflight finds.

---

## 3. Pre-flight: macro version & drift (D-tier)

Stronger to weaker:

**D17 (error) — `embedded macro hash {old}... does not match current SolveHeadless.bas {new}...`**
The repo's `SolveHeadless.bas` has changed since this workbook was last
re-imported. Fix:
```powershell
python import_vba_module.py "C:\path\to\workbook.xlsm"
```
The hash gets stamped into the workbook's custom doc property
`DN38_BAS_SHA256` on every successful import. Subsequent preflight runs
compare and pass cleanly.

**D17 (warning) — `workbook has no DN38_BAS_SHA256 stamp`**
Workbook predates the stamp convention. Run the same `import_vba_module.py`
command once — it both updates the macro and plants the stamp.

**D15 (error) — `embedded macro is missing N of M required function(s)`**
Confirmed root cause of the 2026-05-13 IL TEST regression: that workbook
carried a 627-line outdated macro vs the current 1220-line version. Same
fix as D17:
```powershell
python import_vba_module.py "C:\path\to\workbook.xlsm"
```

**D16 (warning) — `workbook contains stale macro module(s)`**
Manual cleanup: Excel → Alt+F11 → right-click each stale module under
VBAProject → Remove. Decline the export prompt. Stale modules don't break
the solve directly but shadow current function names.

Since 2026-06-04 the check parses the live module list from the VBA
PROJECT stream instead of substring-scanning the whole binary. Removed
modules leave residual name strings in the identifier table / SRP compile
cache until a full VBA recompile, and the old scan false-positived on
them (observed on SolarStone after cleanup). The substring scan remains
only as a fallback for binaries with no parseable PROJECT stream.

**D18 (error) — `embedded macro signature drift on N required function(s)`**
The embedded macro has every required function NAME (so D15 passes) but at
least one parameter count differs from the repo's `SolveHeadless.bas`. On
workbooks predating the `DN38_BAS_SHA256` stamp, D17 can't see it either.
At runtime the orchestrator's `Application.Run` call fails with
`com_error (-2147352567, ..., -2147352562)` — `DISP_E_BADPARAMCOUNT`
(0x8002000E) — *after* the full solve cost has been paid.

Confirmed root cause of the 2026-06-04 SolarStone failure: a 1-arg
embedded `StampActiveProjectColumnHL` met the current 2-arg call in the
read pass; all 37 projects solved, then all four parallel workers errored
without saving. Same fix as D15/D17:
```powershell
python import_vba_module.py "C:\path\to\workbook.xlsm"
```
`--auto-fix` also recovers it (re-import routed into the `_FIXED.xlsm`
sibling). Detection is best-effort: it requires the VBA source to be
extractable (olefile + MS-OVBA decompression); when it isn't, D18 stays
silent and D15/D17 remain the guards.

---

## 4. Pre-flight: calc properties (A-tier)

| Code | Severity | Issue                                | Fix                                                                                    |
|------|----------|--------------------------------------|----------------------------------------------------------------------------------------|
| A1   | error    | `iterateDelta` missing or > 0.0001   | `--auto-fix` patches `xl/workbook.xml` into `<workbook>_FIXED.xlsm`. Or in Excel: File → Options → Formulas → Maximum Change = 0.0001 |
| A2   | error    | `iterate=False`                       | Excel: File → Options → Formulas → enable iterative calculation, save                  |
| A3   | warning  | `calcMode != "manual"`                | Excel: Formulas → Calculation Options → Manual, save                                   |
| A4   | warning  | `fullCalcOnLoad=False`                | Open in Excel, F9, save                                                                |

Iterative calc must be on with delta ≤ 0.0001 because the workbook's
self-circular sticky-IF formulas (Project Inputs row 31 et al.) rely on the
engine to settle before the macro reads converged values. Excel's default
delta of 0.001 leaves Appraisal IRR half-converged and slides Dev Fee to
floor — observed on the IL US Solar 2026-05-13 file.

---

## 5. Pre-flight: structure (B-tier)

| Code | Severity | Issue                                                | Fix                                                                       |
|------|----------|------------------------------------------------------|---------------------------------------------------------------------------|
| B5   | error    | Required sheet missing                                | Restore from baseline (Project Inputs / Appraisal / NPP Calc / Operations / PT Returns / Tax Equity / Perm Debt / CL / Capex / Rate Curves / Global) |
| B7   | error    | Required Project Inputs cell missing or empty         | Restore the cell. Required: F2, F30, F31, F32, F36, F37                  |
| B8   | error    | Workbook or critical sheet protected                  | Excel → Review → Unprotect (Workbook + Project Inputs / Appraisal / NPP Calc) |
| B9   | error    | No active project flagged in row 7                    | Set `Project Inputs!H7:S7` to 1 for the columns to solve                  |

---

## 6. Pre-flight: critical-path cached errors (C-tier)

| Code | Severity | Issue                                                          | Fix                                                                      |
|------|----------|----------------------------------------------------------------|--------------------------------------------------------------------------|
| C10  | error    | Cached error on `Project Inputs!F30:S39`                       | Trace precedents. Common: missing rate curves, broken named ranges       |
| C11  | error    | Cached error on `Appraisal!155:159` (cash flow rows)           | Often missing PI inputs (row 32 Dev Fee, row 11 size) or rate gaps       |
| C12  | error    | Cached error on `Appraisal!H161` (Live Appraisal IRR readout)  | Fix upstream errors (usually C11), then F9 to recalc                     |

The GoalSeek target is `F31 = Appraisal!H161`. An error here means GoalSeek
has nothing valid to drive — every iteration sees the same broken target.

---

## 7. Pre-flight: input bounds (E-tier)

| Code | Severity | Issue                                                            | Fix                                                                         |
|------|----------|------------------------------------------------------------------|-----------------------------------------------------------------------------|
| E13  | warning  | Pre-solve Dev Fee outside `[$0.05..$0.50/W]` on active columns  | Often recoverable. If convergence fails, raise `DEV_FEE_MAX` in `SolveHeadless.bas` line 50 and re-import |
| E14  | warning  | Pre-solve NPP outside `[-$0.20..$0.80/W]`                       | Macro resets to seed ($0.20); usually recoverable                            |

Utility-scale solar deals can carry natural Dev Fees of $1.50–$2.50/W, which
trip E13. The chunked path resets out-of-range values to seed at iteration
0; some models recover (SMP), others get trapped (IL TEST 2026-05-13).

---

## 7.1 Pre-flight: RC config audit (E15)

Added 2026-05-15 after Queen City MD shipped $4.30-$5.34/W Dev Fees on 6 of
13 projects (vs $1.49-$2.02/W on the other 7). Same workbook, same EPC, same
IX — the only material input difference was that cols O-T had Rate Component
5 with `Toggle="Custom"` and an empty Rate Name, while cols H-N had RC5
`Toggle="Generic"` with a populated 5.5% / 2.5% escalator / 35yr term
merchant rate.

The macro converged correctly given the inputs. The Appraisal IRR = WACC
GoalSeek can land on any Dev Fee that satisfies the equation; with one
revenue stream zeroed out, it inflated Dev Fee to compensate. The output
was mathematically valid but economically nonsensical.

| Code  | Severity | What it means |
|-------|----------|---------------|
| E15a  | warning  | One or more RC sub-blocks have `Toggle="Custom"` but the Rate Name cell is empty. Almost always means the Custom rate rows in the Rate Curves tab are also empty — that revenue component contributes zero, and the model compensates elsewhere. |
| E15b  | warning  | The same RC slot (e.g. RC5) has different `Toggle` values across active projects in one workbook (some Generic, some Custom). CAN be intentional (one project on a bespoke tariff) but warrants explicit operator confirmation. |

**Remediation when E15a fires:** for each flagged RC, either flip the Toggle
to `Generic` and populate the Generic rate row, OR keep `Custom` and populate
both the Rate Name AND the per-project rate vector in the Rate Curves tab.
Re-run the full RC1-RC6 audit (active state + term length across equity /
debt / appraisal) before solving — Caroline's [revenue-component-audit]
memory has the full checklist.

**Do NOT `--allow-relaxed` past E15 warnings without confirming the inputs.**
The macro will converge — that's the trap. Convergence is not validation.

---

## 8. Excel process failures

**`(-2147023174, 'RPC server unavailable')`** — Excel.exe exited mid-call.
Most common cause: a single COM call exceeded Excel's ~900s RPC timeout.
The chunked path (`--chunked` or `--workers N`) keeps each per-project
call well under that ceiling.

**`(-2147417848, 'object disconnected from clients')`** — Excel process
died, often from heap exhaustion on the Dashboard/Waterfall recalc.
Workarounds:
- Close all Excel processes before rerunning
- `--no-output-recalc` skips the heavy non-core sheets (Portfolio, AT
  Returns, Corp Model Output, Cust Prop, Waterfall Sensitivity);
  Dashboard and Table still recalc
- Reduce `--workers N`

**Solver hangs at "Opening workbook..."** — usually a zombie EXCEL.EXE
still holding the file. The fix is to kill leftover Excel processes:
```powershell
Get-Process EXCEL -ErrorAction SilentlyContinue | Stop-Process -Force
```
Then rerun. The orchestrator uses `DispatchEx` so it never attaches to an
existing Excel session, but a stuck process can still hold the workbook's
file lock.

---

## 9. Convergence quality

The end-of-run summary reports:

- **`Ship-ready: X/Y projects`** — strict-tier convergence count (plus
  relaxed when `--allow-relaxed`). `X == Y` with `Status: CONVERGED` means
  the merged `_SOLVED.xlsm` is safe to circulate.
- **`Convergence: A strict / B relaxed / C none / D not_attempted`** —
  tier breakdown. `not_attempted` ≠ 0 means a worker crashed before
  reaching some projects; those rows are stale in the merged file.
- **`Merge path:`** (parallel mode only)
  - `openpyxl` — INFO, standard path
  - `vba_fallback` — WARNING, used the Excel-COM stamp helper
  - `copy_master` — **ERROR, do not ship** — both merge paths failed;
    the per-worker `_SOLVED.xlsm` files in the preserved `parent_tmp`
    directory are authoritative
- **`Post-merge verification`** — re-opens the merged file and diffs the
  hard-stamped convergence cells against each worker's reported values.
  Any mismatch flips the run to ERROR and preserves `parent_tmp`.

**Inspect a failed mid-run:**
```powershell
python -m dn38_solver.cli --show-checkpoints <batch_id>
```
Per-project checkpoints survive crashes — successful runs clear them, so
a non-empty result indicates an incomplete prior run.

**`not_converged` projects despite a clean run:** check the project's row
in the workbook's hidden `__SolverResults` sheet. The `heartbeat` column
carries the last phase the macro reached for that project (`solving NPP`,
`solving Appraisal`, `ERROR`, etc.); `irr_gap` and `appr_gap` show how
far GoalSeek got. Common causes:
- Cold-start workbook hitting the per-project iteration cap — rerun once
  more (the previous run's near-converged state seeds a faster second
  pass)
- Out-of-bound E13/E14 inputs not flagged because preflight runs in
  warning mode (use `--strict-preflight` for unattended runs)
- Rate curve gap that produces zero revenue for a year — surface via the
  cached-error scan but it can also slip past as a valid 0

---

## 10. Quick reference

### Run the solver
```powershell
# Required once per workbook (after pulling new SolveHeadless.bas)
python import_vba_module.py "C:\path\to\workbook.xlsm"

# Standard run
.\run_desktop.ps1 "C:\path\to\workbook.xlsm"

# Parallel (cold portfolios ≥ 6 projects)
.\run_desktop.ps1 "C:\path\to\workbook.xlsm" -Workers 2

# No-COM diagnosis only (fast)
python -m dn38_solver.cli "C:\path\to\workbook.xlsm" --diagnose

# Correctness validation: sequential vs parallel diff
python validate_parallel.py "C:\path\to\workbook.xlsm" --workers 2
```

### Useful flags
| Flag                     | Effect                                                                       |
|--------------------------|------------------------------------------------------------------------------|
| `--diagnose`             | Run preflight + macro hash check and exit. No COM startup, no solve.          |
| `--chunked`              | Per-project COM calls (avoid the ~900s RPC timeout on single-shot)            |
| `--workers N`            | N parallel Excel instances. Forces `--chunked`. Cap 8.                        |
| `--strict-preflight`     | Treat preflight warnings (A3/A4/E13/E14/D17-warning) as errors                |
| `--auto-fix`             | Patch A1 (iterateDelta) into `<workbook>_FIXED.xlsm` and proceed              |
| `--allow-relaxed`        | Count relaxed-tier convergence as ship-ready                                   |
| `--no-output-recalc`     | Skip Portfolio / AT Returns / Corp Model Output / Cust Prop / Waterfall recalc |
| `--strip-sheets a,b,c`   | Delete non-critical sheets from temp copy before solving                       |
| `--show-checkpoints ID`  | Print per-project checkpoints for a crashed batch                              |
| `--history`              | List recent runs                                                                |

### Architecture (one paragraph)

CLI parses args → orchestrator runs Phase 0 preflight (A/B/C/D/E checks,
zip-level XML reads, no COM) → if pre-flight passes, direct_runner copies
the workbook to a temp dir, opens via `DispatchEx`, runs `SolveHeadless`
or per-project chunked entry points (`InitSolveEnvHL` / `SolveOneProjectByColHL`
/ `FinalizeSolveEnvHL`), reads results from the hidden `__SolverResults`
sheet, SaveAs the converged file. Auto-recovery wraps the macro call: a
recoverable COM exception triggers a one-shot close → re-import → reopen
→ retry cycle. Parallel runner forks N workers each running the same
single-process pipeline against a round-robin slice of projects, then
merges per-project convergence cells into a single `_SOLVED.xlsm`.

### Safety rules

These are load-bearing and have caused incidents when violated:

1. **Never openpyxl-save a .xlsm between a macro re-import and a solver
   run.** openpyxl's `keep_vba=True` preserves the VBA blob but strips or
   rewrites data validation extensions, conditional formatting state, and
   calculation chain caches. The macro tolerates direct cell writes but
   throws the generic VBA error (DISP_E_EXCEPTION with secondary 0x800a9c68) inside the first GoalSeek. Mid-workflow .xlsm
   mutations must go through `dn38_solver.com.com_edit.edit_xlsm`, which
   uses Excel COM SaveAs and is verified safe.
2. **Pre-flight A2 (`iterate=False`) is never auto-fixed.** Flipping
   iterative calc on without an audit can mask real circular-reference
   bugs in the model. Only A1 (iterateDelta = 0.0001) is auto-fixable.
3. **Critical sheets are never stripped.** `--strip-sheets` silently
   refuses Dashboard / Table / PT Returns / NPP Calc / Appraisal / Perm
   Debt / Tax Equity / CL / Project Inputs even if requested.

### File inventory

| File                                            | Purpose                                                                       |
|-------------------------------------------------|-------------------------------------------------------------------------------|
| `SolveHeadless.bas`                              | VBA macro (headless, deterministic recalc, per-project telemetry)              |
| `import_vba_module.py`                           | Imports .bas into a workbook; stamps `DN38_BAS_SHA256` for drift detection     |
| `dn38_solver/com/direct_runner.py`               | Direct COM execution, telemetry ingestion, auto-recovery wiring                |
| `dn38_solver/com/auto_recovery.py`               | Close → re-import → reopen → retry helper                                      |
| `dn38_solver/com/hresult.py`                     | HRESULT decoder with recovery hints                                            |
| `dn38_solver/com/com_edit.py`                    | Canonical Excel COM helper for safe .xlsm mutation                             |
| `dn38_solver/shadow/preflight.py`                | Bank-grade preflight (A/B/C/D/E checks)                                        |
| `dn38_solver/solver/orchestrator.py`             | Phase 0 (preflight) + Phase 1 (solve) + Phase 2 (verify) loop                  |
| `dn38_solver/cli.py`                             | CLI entry point with `--diagnose`, `--chunked`, `--workers`, etc.              |
| `dn38_solver/storage/database.py`                | SQLite run + checkpoint persistence                                            |
| `dn38_solver/reporting/export_xlsx.py`           | Branded summary .xlsx                                                          |
| `tests/test_preflight.py`                        | Preflight unit + integration tests                                             |
| `validate_parallel.py`                           | Sequential-vs-parallel correctness gate                                        |
| `run_portfolio_test.py`                          | 8-project integration test                                                     |

### Resolved-issue archive

See `docs/RESOLVED_ISSUES.md` for the historical record of issues #1–#7
(MsgBox freezes, calc-mode races, GoalSeek over-iteration, CalcModelCore
correctness-vs-performance, RPC timeouts) and the per-version performance
table. The fixes are in the macro and runner; the doc exists for
historical context when investigating regressions.
