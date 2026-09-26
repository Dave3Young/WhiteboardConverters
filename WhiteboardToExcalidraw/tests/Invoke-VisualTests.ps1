#requires -Version 5.1
<#
.SYNOPSIS
Runs the Python/Playwright geometry checks using the tests\.venv made by Setup-VisualTests.ps1.
Any extra arguments are passed through, e.g.:
    .\tests\Invoke-VisualTests.ps1 --sample AssumptionGrid ImageBoard --version solid
Results are written to tests\out\visual\results.json.
#>
param([Parameter(ValueFromRemainingArguments)] [string[]] $PassThru)
$venvPython = Join-Path $PSScriptRoot '.venv\Scripts\python.exe'
if (-not (Test-Path $venvPython)) { throw 'Visual tests are not set up yet: run .\tests\Setup-VisualTests.ps1 first.' }
& $venvPython (Join-Path $PSScriptRoot 'visual\run_visual_checks.py') @PassThru
exit $LASTEXITCODE
