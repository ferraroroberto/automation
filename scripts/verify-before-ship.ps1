# Pre-ship verification gate.
#
# Runs the pipeline documented in CLAUDE.md's "Verification" section as one
# pass/fail: byte-compile, then the two existing unittest suites.
#
# Usage:
#   C:/Windows/System32/WindowsPowerShell/v1.0/powershell.exe -File scripts/verify-before-ship.ps1
#   (never bare pwsh -- it's a 0-byte WindowsApps reparse stub on this machine that fails non-interactively)
#
# Exits non-zero on the first failure with the offending output left visible.

$ErrorActionPreference = "Stop"
$repoRoot = Split-Path -Parent $PSScriptRoot
$python = Join-Path $repoRoot ".venv\Scripts\python.exe"
$sw = [System.Diagnostics.Stopwatch]::StartNew()

function Fail($message) {
    Write-Host ""
    Write-Host "[X] $message" -ForegroundColor Red
    Write-Host ("Failed after {0:n1}s." -f $sw.Elapsed.TotalSeconds) -ForegroundColor Red
    exit 1
}

if (-not (Test-Path $python)) {
    Fail ".venv missing -- create it and install requirements.txt first."
}

Push-Location $repoRoot
try {
    Write-Host "==> compileall (repo-wide, excluding .venv)..." -ForegroundColor Cyan
    & $python -m compileall -q -x '\.venv' .
    if ($LASTEXITCODE -ne 0) { Fail "byte-compile failed." }

    Write-Host "==> unittest (system/foldersearcher)..." -ForegroundColor Cyan
    & $python -m unittest discover -s system/foldersearcher -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "foldersearcher unittest suite failed." }

    Write-Host "==> unittest (system.test_local_config_hygiene)..." -ForegroundColor Cyan
    & $python -m unittest system.test_local_config_hygiene
    if ($LASTEXITCODE -ne 0) { Fail "local config hygiene unittest suite failed." }
}
finally {
    Pop-Location
}

$sw.Stop()
Write-Host ""
Write-Host ("[OK] Ready to ship -- all checks passed in {0:n1}s." -f $sw.Elapsed.TotalSeconds) -ForegroundColor Green
exit 0
