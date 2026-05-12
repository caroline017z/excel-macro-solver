# Accept the workbook as the first arg; fall back to a single-project IL
# default for ad-hoc one-off solves. Portfolio runs should pass the path
# explicitly, e.g.:
#   .\run_desktop.ps1 "C:\Users\CarolineZepecki\OneDrive - 38 Degrees North\38DN-MD_CI Renewables_Project Blue Crab_Pricing Model_2026.04.21.xlsm"
param(
    [string]$Workbook = "$env:USERPROFILE\Box\2. Deal Flow\Novel Energy Solutions\Pricing Model\38DN-IL_Novel Energy Solutions_lease financing_PricingModel_100% Commercial_2026.04.15.xlsm",
    [int]$TimeoutSec = 3600,
    [switch]$NoChunked,
    [switch]$StrictOnly
)

$ErrorActionPreference = "Stop"
$repo = $PSScriptRoot

if (-not (Test-Path $Workbook)) {
    Write-Host "Workbook not found at: $Workbook" -ForegroundColor Red
    Read-Host "Press Enter to close"
    exit 1
}

Set-Location $repo

# Chunked + allow-relaxed are the right defaults for portfolio runs:
#   --chunked       avoids the ~900s COM RPC ceiling on long cold solves
#   --allow-relaxed counts +/-0.5pp equity hits as run-level converged
# Pass -NoChunked / -StrictOnly to opt out.
$cliArgs = @($Workbook, "--timeout", $TimeoutSec)
if (-not $NoChunked)  { $cliArgs += "--chunked" }
if (-not $StrictOnly) { $cliArgs += "--allow-relaxed" }

python -u -m dn38_solver.cli @cliArgs
Read-Host "Done. Press Enter to close"
