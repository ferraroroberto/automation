# Pre-ship verification gate.
#
# Runs the pipeline documented in CLAUDE.md's "Verification" section as one
# pass/fail: byte-compile, the structure, instruction and doc refs ratchets, then the unittest suites.
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

    # Opt-in via .fleet.toml [structure_ratchet]. Exit 1 is a new or grown offender or a
    # root folder outside the declared shape; exit 3 is "not confirmed" (git unusable,
    # unparseable file, unknown shape). Both fail: 3 is never a pass.
    Write-Host "==> structure ratchet..." -ForegroundColor Cyan
    & $python "scripts/structure_ratchet.py"
    if ($LASTEXITCODE -eq 3) { Fail "structure ratchet NOT CONFIRMED (exit 3) -- it could not measure the tree." }
    if ($LASTEXITCODE -ne 0) { Fail "structure ratchet failed (see findings above)." }

    # Opt-in via .fleet.toml [instruction_ratchet]. Exit 1 is a violation-cap hit count above
    # its baseline or an instruction file past its size cap; exit 3 is "not confirmed" (git
    # unusable, unreadable file). Both fail: 3 is never a pass.
    Write-Host "==> instruction ratchet..." -ForegroundColor Cyan
    & $python "scripts/instruction_ratchet.py"
    if ($LASTEXITCODE -eq 3) { Fail "instruction ratchet NOT CONFIRMED (exit 3) -- it could not measure the instruction files." }
    if ($LASTEXITCODE -ne 0) { Fail "instruction ratchet failed (see findings above)." }

    # Opt-in via .fleet.toml [doc_refs_ratchet]. Exit 1 is a doc reference to a repo path, module
    # or relative link that does not exist and is not baselined; exit 3 is "not confirmed" (git
    # unusable, a non-UTF-8 doc, an unclosed fence). Both fail: 3 is never a pass.
    Write-Host "==> doc refs ratchet..." -ForegroundColor Cyan
    & $python "scripts/doc_refs_ratchet.py"
    if ($LASTEXITCODE -eq 3) { Fail "doc refs ratchet NOT CONFIRMED (exit 3) -- it could not read the docs or the tree." }
    if ($LASTEXITCODE -ne 0) { Fail "doc refs ratchet failed (see findings above)." }

    Write-Host "==> unittest (tests/system/foldersearcher)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/system/foldersearcher -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "foldersearcher unittest suite failed." }

    Write-Host "==> unittest (tests/system)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/system -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "system unittest suites (config hygiene, outlook web mail, word to markdown) failed." }

    Write-Host "==> unittest (tests/google)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/google -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "photos weekly album unittest suite failed." }

    Write-Host "==> unittest (tests/notion/articles_sync)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/notion/articles_sync -p "test*.py"
    if ($LASTEXITCODE -ne 0) { Fail "notion articles sync unittest suite failed." }

    Write-Host "==> unittest (tests/system/wifi)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/system/wifi -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "wifi netsh decoding unittest suite failed." }

    Write-Host "==> unittest (tests/audio)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/audio -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "audio unittest suite failed." }

    Write-Host "==> unittest (tests/excel/accounting_stripe)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/excel/accounting_stripe -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "accounting_stripe unittest suite failed." }

    Write-Host "==> unittest (tests/linkedin devtools)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/linkedin/profiles_data_extractor/devtools -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "linkedin devtools unittest suite failed." }

    Write-Host "==> unittest (tests/linkedin common)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/linkedin/profiles_data_extractor/common -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "linkedin common unittest suite failed." }

    Write-Host "==> unittest (tests/linkedin check_ip)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/linkedin/check_ip -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "linkedin check_ip unittest suite failed." }

    Write-Host "==> unittest (tests/system/apple)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/system/apple -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "apple vcard unittest suite failed." }

    Write-Host "==> unittest (tests/system/quickdeck)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/system/quickdeck -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "quickdeck unittest suite failed." }

    Write-Host "==> unittest (tests/image)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/image -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "image unittest suite failed." }

    Write-Host "==> unittest (tests/video)..." -ForegroundColor Cyan
    & $python -m unittest discover -s tests/video -p "test_*.py"
    if ($LASTEXITCODE -ne 0) { Fail "gpu_recovery console decoding unittest suite failed." }

    Write-Host "==> pytest (tests/parking)..." -ForegroundColor Cyan
    & $python -m pytest -q tests/parking
    if ($LASTEXITCODE -ne 0) { Fail "parking pytest suite failed." }
}
finally {
    Pop-Location
}

$sw.Stop()
Write-Host ""
Write-Host ("[OK] Ready to ship -- all checks passed in {0:n1}s." -f $sw.Elapsed.TotalSeconds) -ForegroundColor Green
exit 0
