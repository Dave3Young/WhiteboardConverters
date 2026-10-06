#requires -Version 5.1
<#
.SYNOPSIS
One-time setup for the Python/Playwright visual checks (tests\visual\run_visual_checks.py).

.DESCRIPTION
1. Finds Python 3.9+ (the "py" launcher first, then python.exe -- ignoring the Microsoft Store
   "python.exe" alias that only opens the Store). With -InstallPython, installs Python 3.12
   for the current user via winget if none is found.
2. Creates a virtual environment in tests\.venv (git-ignored).
3. Installs tests\visual\requirements.txt (Playwright) into it.
4. Downloads Playwright's Chromium build (into %LOCALAPPDATA%\ms-playwright).
5. Installs @excalidraw/utils into tests\visual\node_modules (git-ignored) with npm, so the
   report can show each converted scene as Excalidraw draws it. Skipped, with a warning, when
   Node.js isn't installed.

Safe to re-run: existing pieces are reused.

.EXAMPLE
.\tests\Setup-VisualTests.ps1

.EXAMPLE
.\tests\Setup-VisualTests.ps1 -InstallPython
#>
[CmdletBinding()]
param([switch] $InstallPython)

$ErrorActionPreference = 'Stop'
$venv = Join-Path $PSScriptRoot '.venv'
$requirements = Join-Path $PSScriptRoot 'visual\requirements.txt'

function Test-PythonCommand {
    param([string[]] $Command)
    try {
        $exe = $Command[0]; $rest = @($Command | Select-Object -Skip 1)
        $out = & $exe @rest -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($LASTEXITCODE -eq 0 -and "$out" -match '^3\.(\d+)$' -and [int]$Matches[1] -ge 9) { return "$out".Trim() }
    } catch { }
    return $null
}

function Find-Python {
    foreach ($candidate in @(@('py', '-3'), @('python'), @('python3'))) {
        if (-not (Get-Command $candidate[0] -ErrorAction SilentlyContinue)) { continue }
        $ver = Test-PythonCommand $candidate
        if ($ver) { return [pscustomobject]@{ Command = $candidate; Version = $ver } }
    }
    return $null
}

Write-Host '1/5 Looking for Python 3.9+ ...'
$python = Find-Python
if (-not $python) {
    if (-not $InstallPython) {
        throw ("Python 3.9+ was not found. Re-run with -InstallPython to install it with winget, " +
               "or install it from https://www.python.org/downloads/ and re-run.")
    }
    if (-not (Get-Command winget -ErrorAction SilentlyContinue)) { throw 'winget is not available; install Python manually and re-run.' }
    Write-Host '    Installing Python 3.12 for the current user with winget ...'
    winget install --exact --id Python.Python.3.12 --scope user --accept-package-agreements --accept-source-agreements --silent
    # Pick up the new PATH entries without restarting the shell.
    $env:Path = [Environment]::GetEnvironmentVariable('Path', 'User') + ';' + [Environment]::GetEnvironmentVariable('Path', 'Machine')
    $python = Find-Python
    if (-not $python) {
        $direct = Join-Path $env:LOCALAPPDATA 'Programs\Python\Python312\python.exe'
        if (Test-Path $direct) { $python = [pscustomobject]@{ Command = @($direct); Version = (Test-PythonCommand @($direct)) } }
    }
    if (-not $python) { throw 'Python was installed but could not be found; open a new PowerShell window and re-run.' }
}
Write-Host ("    Using Python {0} ({1})" -f $python.Version, ($python.Command -join ' '))

Write-Host "2/5 Virtual environment: $venv"
$venvPython = Join-Path $venv 'Scripts\python.exe'
if (-not (Test-Path $venvPython)) {
    $exe = $python.Command[0]; $rest = @($python.Command | Select-Object -Skip 1)
    & $exe @rest -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw 'Creating the virtual environment failed.' }
} else { Write-Host '    (already exists)' }

Write-Host '3/5 Installing Python packages ...'
& $venvPython -m pip install --disable-pip-version-check --upgrade pip | Out-Null
& $venvPython -m pip install --disable-pip-version-check -r $requirements
if ($LASTEXITCODE -ne 0) { throw 'pip install failed.' }

Write-Host '4/5 Installing Playwright Chromium ...'
& $venvPython -m playwright install chromium
if ($LASTEXITCODE -ne 0) { throw 'Playwright browser install failed.' }

Write-Host '5/5 Installing @excalidraw/utils (draws the Excalidraw pictures in the report) ...'
if (Get-Command npm -ErrorAction SilentlyContinue) {
    Push-Location (Join-Path $PSScriptRoot 'visual')
    try { npm ci --no-audit --no-fund } finally { Pop-Location }
    if ($LASTEXITCODE -ne 0) { throw 'npm ci failed.' }
} else {
    Write-Warning 'npm (Node.js) was not found: the checks still run, but the report will have no Excalidraw pictures.'
}

Write-Host ''
Write-Host 'Done. Run the visual checks with:'
Write-Host '    .\tests\Invoke-VisualTests.ps1'
