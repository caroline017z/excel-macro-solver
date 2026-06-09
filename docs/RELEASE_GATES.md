# Release Gates — warm / cold / stress Excel regression tier

The fast unit suite (run in CI on ubuntu + windows) covers the pure-Python
surface — preflight, the Python↔VBA contract, the watchdog logic, the gate
rules. It cannot cover the part of the system that only exists when Excel is
driving: the chunked solve loop, the in-flight watchdog kill, auto-recovery,
and the parallel merge. That's ~60% of the runtime risk, and it's exactly
the surface a new `SolveHeadless.bas` propagates into every workbook.

This tier closes that gap. It runs real solves against real workbooks on a
Windows + Excel box and judges them with the rules in
`dn38_solver/validation/release_gates.py` (themselves unit-tested in
`tests/test_release_gates.py`). It is **opt-in** — it never runs in CI.

## When to run it

Before shipping any change that touches:

- `SolveHeadless.bas` (every re-import propagates it to every workbook)
- `dn38_solver/com/vba_contract.py` or the layout constants in `config.py`
- the chunked loop, watchdog, auto-recovery, or parallel merge code

## The three gates

| Gate | Fixture | Passes when |
|---|---|---|
| **warm** | a **pre-solved** workbook | run status is CONVERGED and total wall time ≤ `DN38_WARM_MAX_SEC` (default 600s); optional P95-per-project budget via `DN38_WARM_P95_SEC`. Guards **speed** regressions. |
| **cold** | an **unsolved** workbook | every scored project reaches **strict** convergence (±0.25pp). Relaxed is shippable for a bid but **not** acceptable for this gate. Guards **correctness** regressions. |
| **stress** | a **multi-project** workbook (8+) | `DN38_STRESS_RUNS` (default 3) consecutive parallel runs (`DN38_STRESS_WORKERS`, default 2) all converge with per-project NPP stable across runs (≤ `DN38_STRESS_NPP_TOL`, default $0.005/W). Guards **RPC resilience + merge** correctness. |

Deliberate placeholder skips and worker-crash `not_attempted` rows are
excluded from the cold gate's scoring — they are not convergence outcomes.

## How to run

```powershell
# All three gates:
.\run_release_gate.ps1 `
  -Warm   "C:\path\to\pre-solved.xlsm" `
  -Cold   "C:\path\to\unsolved.xlsm" `
  -Stress "C:\path\to\portfolio.xlsm"

# A single gate while iterating (omit the others):
.\run_release_gate.ps1 -Cold "C:\path\to\unsolved.xlsm"
```

Or drive pytest directly:

```powershell
$env:DN38_EXCEL_TESTS = "1"
$env:DN38_COLD_FIXTURE = "C:\path\to\unsolved.xlsm"
python -m pytest tests/test_excel_regression.py -v -m excel_integration
```

Each gate **skips** (does not fail) when its fixture env var is unset or the
file is missing, so a partial run is fine. With `DN38_EXCEL_TESTS` unset the
entire tier skips — which is why it is inert in CI.

The solves run with `auto_fix=True`, so a fixture whose embedded macro has
drifted is patched into a `_FIXED.xlsm` sibling first (the original is never
modified). Budget time accordingly: a cold multi-project solve is minutes
per project.

## Environment contract

| Var | Meaning |
|---|---|
| `DN38_EXCEL_TESTS=1` | master switch (required) |
| `DN38_WARM_FIXTURE` | path to the pre-solved workbook |
| `DN38_COLD_FIXTURE` | path to the unsolved workbook |
| `DN38_STRESS_FIXTURE` | path to the multi-project workbook |
| `DN38_WARM_MAX_SEC` | warm total wall-time budget (default 600) |
| `DN38_WARM_P95_SEC` | warm P95-per-project budget (optional) |
| `DN38_STRESS_WORKERS` | parallel workers for the stress run (default 2) |
| `DN38_STRESS_RUNS` | consecutive clean runs required (default 3) |
| `DN38_STRESS_NPP_TOL` | max cross-run NPP drift treated as clean (default 0.005) |
