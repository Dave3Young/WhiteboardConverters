#requires -Version 5.1
<#
.SYNOPSIS
Runs the Python/Playwright visual checks using the tests\.venv made by Setup-VisualTests.ps1.
Any extra arguments are passed through, e.g.:
    .\tests\Invoke-VisualTests.ps1 --sample AssumptionGrid ImageBoard --version v2_solid
Opens tests\out\visual\report.html when done (skip with -NoOpen).
#>
param([switch] $NoOpen, [Parameter(ValueFromRemainingArguments)] [string[]] $PassThru)
$venvPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $venvPython)) { throw 'Visual tests are not set up yet: run .\tests\Setup-VisualTests.ps1 first.' }
& $venvPython (Join-Path $PSScriptRoot 'visual\run_visual_checks.py') @PassThru
$code = $LASTEXITCODE
$report = Join-Path $PSScriptRoot 'out\visual\report.html'
if (-not $NoOpen -and (Test-Path $report)) { Invoke-Item $report }
exit $code
