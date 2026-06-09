# Release gate: warm / cold / stress Excel regression tier.
#
# Run on a Windows box with Excel before shipping a change to
# SolveHeadless.bas or vba_contract.py — it exercises the chunked loop,
# watchdog, recovery, and parallel merge against real workbooks, which the
# cross-platform CI cannot. The pass/fail rules are unit-tested in
# tests/test_release_gates.py; this script drives the real solves.
#
# Usage:
#   .\run_release_gate.ps1 -Warm <wb.xlsm> -Cold <wb.xlsm> -Stress <wb.xlsm>
#
# Any fixture you omit is skipped, so you can run a single gate while
# iterating. Fixtures can also come from the DN38_*_FIXTURE env vars.
param(
    [string]$Warm          = $env:DN38_WARM_FIXTURE,
    [string]$Cold          = $env:DN38_COLD_FIXTURE,
    [string]$Stress        = $env:DN38_STRESS_FIXTURE,
    [int]$WarmMaxSec       = 600,
    [int]$StressWorkers    = 2,
    [int]$StressRuns       = 3
)

$ErrorActionPreference = "Stop"
Set-Location $PSScriptRoot

if (-not ($Warm -or $Cold -or $Stress)) {
    Write-Host "No fixtures provided." -ForegroundColor Yellow
    Write-Host "  Pass at least one of -Warm / -Cold / -Stress (paths to .xlsm models),"
    Write-Host "  or set DN38_WARM_FIXTURE / DN38_COLD_FIXTURE / DN38_STRESS_FIXTURE."
    exit 2
}

$env:DN38_EXCEL_TESTS = "1"
if ($Warm)   { $env:DN38_WARM_FIXTURE   = $Warm }
if ($Cold)   { $env:DN38_COLD_FIXTURE   = $Cold }
if ($Stress) { $env:DN38_STRESS_FIXTURE = $Stress }
$env:DN38_WARM_MAX_SEC   = "$WarmMaxSec"
$env:DN38_STRESS_WORKERS = "$StressWorkers"
$env:DN38_STRESS_RUNS    = "$StressRuns"

Write-Host "Running release gate (warm/cold/stress)..." -ForegroundColor Cyan
Write-Host "  Warm:   $(if ($Warm)   { $Warm }   else { '(skipped)' })"
Write-Host "  Cold:   $(if ($Cold)   { $Cold }   else { '(skipped)' })"
Write-Host "  Stress: $(if ($Stress) { $Stress } else { '(skipped)' })  (workers=$StressWorkers, runs=$StressRuns)"

python -m pytest tests/test_excel_regression.py -v -m excel_integration
$code = $LASTEXITCODE

if ($code -eq 0) {
    Write-Host "RELEASE GATE PASSED" -ForegroundColor Green
} else {
    Write-Host "RELEASE GATE FAILED (exit $code)" -ForegroundColor Red
}
exit $code
