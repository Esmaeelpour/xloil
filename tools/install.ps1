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
.PARAMETER Wheel
    Install this already-downloaded wheel file instead of downloading one
    (for offline or locked-down computers).
#>
param(
    [string]$Python = "",
    [string]$Features = "",
    [switch]$Yes,
    [string]$Version = "latest",
    [string]$Wheel = ""
)

$ErrorActionPreference = "Stop"
$Repo = "Esmaeelpour/xloil"

function Fail($msg) {
    Write-Host "ERROR: $msg" -ForegroundColor Red
    exit 1
}

function Retry([scriptblock]$Action, [string]$What) {
    # Networks drop connections: try a few times before giving up
    for ($i = 1; $i -le 4; $i++) {
        try { return & $Action }
        catch {
            if ($i -eq 4) { throw }
            Write-Host "$What failed ($($_.Exception.Message)), retrying..."
            Start-Sleep -Seconds (3 * $i)
        }
    }
}

# TLS 1.2 for older Windows PowerShell defaults
try { [Net.ServicePointManager]::SecurityProtocol = [Net.ServicePointManager]::SecurityProtocol -bor [Net.SecurityProtocolType]::Tls12 } catch {}

# 1. Find a Python that actually starts. Try the active conda env / venv, then
#    the py launcher, then python on PATH, and report why any candidate failed.
function Probe([string]$Exe, [string[]]$PreArgs) {
    # Windows PowerShell turns a native program's error output into a fatal error
    # when ErrorActionPreference is Stop: relax it so failures can be reported
    $ErrorActionPreference = "Continue"
    $out = & $Exe @PreArgs -c "import sys; print(sys.executable)" 2>&1 | ForEach-Object { "$_" }
    if ($LASTEXITCODE -eq 0 -and $out) { return @{ ok = $true; exe = ($out | Select-Object -Last 1).Trim() } }
    return @{ ok = $false; why = (($out | Where-Object { $_ -match 'Error|error|not found|No module' } | Select-Object -First 2) -join " / ") }
}

$failures = @()
if ($Python) {
    $r = Probe $Python @()
    if (-not $r.ok) { $failures += "$Python : $($r.why)" } else { $Python = $r.exe }
} else {
    $candidates = @()
    if ($env:CONDA_PREFIX -and (Test-Path "$env:CONDA_PREFIX\python.exe")) { $candidates += ,@("$env:CONDA_PREFIX\python.exe", @()) }
    if ($env:VIRTUAL_ENV -and (Test-Path "$env:VIRTUAL_ENV\Scripts\python.exe")) { $candidates += ,@("$env:VIRTUAL_ENV\Scripts\python.exe", @()) }
    if (Get-Command py -ErrorAction SilentlyContinue) { $candidates += ,@("py", @("-3")) }
    if (Get-Command python -ErrorAction SilentlyContinue) { $candidates += ,@("python", @()) }
    foreach ($c in $candidates) {
        $r = Probe $c[0] $c[1]
        if ($r.ok) { $Python = $r.exe; break }
        $failures += "$($c[0]) $($c[1] -join ' '): $($r.why)"
    }
}
if (-not $Python -or -not (Test-Path $Python)) {
    $msg = "No working Python found."
    if ($failures) { $msg += "`n  Tried:`n    " + ($failures -join "`n    ") }
    if ($env:PYTHONHOME -or $env:PYTHONPATH) {
        $msg += "`n  PYTHONHOME/PYTHONPATH are set ('$env:PYTHONHOME' / '$env:PYTHONPATH'): these often stop Python " +
                "starting ('No module named encodings'). Clear them for this window with:`n" +
                "    Remove-Item Env:PYTHONHOME, Env:PYTHONPATH -ErrorAction SilentlyContinue"
    }
    $msg += "`n  Or open an Anaconda Prompt / activate your venv first, or pass -Python <path to python.exe>," +
            "`n  or install 64-bit Python 3.9-3.14 from https://www.python.org (no admin needed)."
    Fail $msg
}

$info = & $Python -c "import sys, struct; print(str(sys.version_info[0]) + str(sys.version_info[1]), struct.calcsize('P') * 8)"
$pyTag, $bits = $info.Trim().Split(" ")
Write-Host "Using Python: $Python (cp$pyTag, $bits-bit)"
if ($bits -ne "64") { Fail "xlOil needs a 64-bit Python (and 64-bit Excel)." }

if ($Wheel) {
    if (-not (Test-Path $Wheel)) { Fail "Wheel file not found: $Wheel" }
    if ((Split-Path $Wheel -Leaf) -notmatch "-cp$pyTag-cp$pyTag-win_amd64\.whl$") {
        Fail "$(Split-Path $Wheel -Leaf) is not built for this Python (cp$pyTag, 64-bit)"
    }
    $dest = (Resolve-Path $Wheel).Path
} else {

# 2. Find the matching wheel in the release
$api = if ($Version -eq "latest") { "https://api.github.com/repos/$Repo/releases/latest" } else { "https://api.github.com/repos/$Repo/releases/tags/$Version" }
try { $release = Retry { Invoke-RestMethod -Uri $api -Headers @{ "User-Agent" = "xloil-installer" } } "Reading the release" }
catch { Fail "Could not read release '$Version' from GitHub: $($_.Exception.Message)" }

$asset = $release.assets | Where-Object { $_.name -match "^xloil-.*-cp$pyTag-cp$pyTag-win_amd64\.whl$" } | Select-Object -First 1
if (-not $asset) {
    $tags = ($release.assets | Where-Object { $_.name -match "-cp(\d+)-" } | ForEach-Object { $Matches[1] }) -join ", "
    Fail "Release $($release.tag_name) has no wheel for Python cp$pyTag. Available: $tags"
}

# 3. Download and install it
$dest = Join-Path $env:TEMP $asset.name
Write-Host "Downloading $($asset.name) from release $($release.tag_name)"
Retry { Invoke-WebRequest -Uri $asset.browser_download_url -OutFile $dest -UseBasicParsing } "Download" | Out-Null
}
Unblock-File $dest

# Excel locks xlOil's files while it runs, which makes the install fail part-way
while (Get-Process EXCEL -ErrorAction SilentlyContinue) {
    Write-Host "Excel is running. Close every Excel window (check Task Manager for EXCEL.EXE)," -ForegroundColor Yellow
    if ($Yes) { Fail "Excel is running: close it and run the installer again." }
    Read-Host "then press Enter to continue" | Out-Null
}

& $Python -m pip install --upgrade $dest
if ($LASTEXITCODE -ne 0) { Fail "pip install failed" }

# 4. Guided setup: features, add-in install, checks
$setupArgs = @("-m", "xloil.command_line", "setup")
if ($Features) { $setupArgs += @("--features", $Features) }
if ($Yes) { $setupArgs += "--yes" }
& $Python @setupArgs
exit $LASTEXITCODE
