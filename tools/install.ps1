<#
.SYNOPSIS
    One-step installer for xlOil (Esmaeelpour/xloil fork).

.DESCRIPTION
    Finds a 64-bit Python (python.org, conda or venv), downloads the matching
    xlOil wheel from the GitHub release, installs it with pip and runs the
    guided `xloil setup`, which installs the Excel add-in and checks everything.

    Run from a normal PowerShell window (in an activated conda env or venv if
    you want xlOil to use that environment):

        powershell -ExecutionPolicy Bypass -c "irm https://github.com/Esmaeelpour/xloil/releases/latest/download/install.ps1 | iex"

    Or download it and run with options:

        .\install.ps1 -Python C:\envs\risk\python.exe -Features pandas,charts,qt -Yes

.PARAMETER Python
    The python.exe to use. Default: the active conda env / venv, then `py -3`, then `python`.
.PARAMETER Features
    Comma-separated features for `xloil setup` (pandas,polars,charts,qt,intellisense,cache).
    If omitted, setup asks.
.PARAMETER Yes
    Accept setup's defaults without asking.
.PARAMETER Version
    Release tag to install, e.g. v0.22.1.post2. Default: the latest release.
#>
param(
    [string]$Python = "",
    [string]$Features = "",
    [switch]$Yes,
    [string]$Version = "latest"
)

$ErrorActionPreference = "Stop"
$Repo = "Esmaeelpour/xloil"

function Fail($msg) {
    Write-Host "ERROR: $msg" -ForegroundColor Red
    exit 1
}

# 1. Find Python
if (-not $Python) {
    if ($env:CONDA_PREFIX -and (Test-Path "$env:CONDA_PREFIX\python.exe")) { $Python = "$env:CONDA_PREFIX\python.exe" }
    elseif ($env:VIRTUAL_ENV -and (Test-Path "$env:VIRTUAL_ENV\Scripts\python.exe")) { $Python = "$env:VIRTUAL_ENV\Scripts\python.exe" }
    elseif (Get-Command py -ErrorAction SilentlyContinue) { $Python = (& py -3 -c "import sys; print(sys.executable)") }
    elseif (Get-Command python -ErrorAction SilentlyContinue) { $Python = (& python -c "import sys; print(sys.executable)") }
}
if (-not $Python -or -not (Test-Path $Python)) {
    Fail "No Python found. Install 64-bit Python 3.9-3.14 from https://www.python.org or Miniconda, then run this again (or pass -Python <path to python.exe>)."
}

$info = & $Python -c "import sys, struct; print(str(sys.version_info[0]) + str(sys.version_info[1]), struct.calcsize('P') * 8)"
$pyTag, $bits = $info.Trim().Split(" ")
Write-Host "Using Python: $Python (cp$pyTag, $bits-bit)"
if ($bits -ne "64") { Fail "xlOil needs a 64-bit Python (and 64-bit Excel)." }

# 2. Find the matching wheel in the release
$api = if ($Version -eq "latest") { "https://api.github.com/repos/$Repo/releases/latest" } else { "https://api.github.com/repos/$Repo/releases/tags/$Version" }
try { $release = Invoke-RestMethod -Uri $api -Headers @{ "User-Agent" = "xloil-installer" } }
catch { Fail "Could not read release '$Version' from GitHub: $($_.Exception.Message)" }

$wheel = $release.assets | Where-Object { $_.name -match "^xloil-.*-cp$pyTag-cp$pyTag-win_amd64\.whl$" } | Select-Object -First 1
if (-not $wheel) {
    $tags = ($release.assets | Where-Object { $_.name -match "-cp(\d+)-" } | ForEach-Object { $Matches[1] }) -join ", "
    Fail "Release $($release.tag_name) has no wheel for Python cp$pyTag. Available: $tags"
}

# 3. Download and install it
$dest = Join-Path $env:TEMP $wheel.name
Write-Host "Downloading $($wheel.name) from release $($release.tag_name)"
Invoke-WebRequest -Uri $wheel.browser_download_url -OutFile $dest -UseBasicParsing
Unblock-File $dest
& $Python -m pip install --upgrade $dest
if ($LASTEXITCODE -ne 0) { Fail "pip install failed" }

# 4. Guided setup: features, add-in install, checks
$setupArgs = @("-m", "xloil.command_line", "setup")
if ($Features) { $setupArgs += @("--features", $Features) }
if ($Yes) { $setupArgs += "--yes" }
& $Python @setupArgs
exit $LASTEXITCODE
