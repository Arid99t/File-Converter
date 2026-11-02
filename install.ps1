# FileConverter installer - per-user, no admin rights needed.
$ErrorActionPreference = "Continue"   # native tools write warnings to stderr; we check exit codes instead
$here    = Split-Path -Parent $MyInvocation.MyCommand.Path
$dataDir = Join-Path $env:LOCALAPPDATA "FileConverter"
$venv    = Join-Path $dataDir "venv"

function Step($m) { Write-Host "`n==> $m" -ForegroundColor Cyan }

function Test-Python($exe, $pre) {
    try {
        $v = & $exe @pre -c "import sys; print('%d.%d' % sys.version_info[:2])" 2>$null
        if ($LASTEXITCODE -eq 0 -and $v -match '^3\.(\d+)$' -and [int]$Matches[1] -ge 9) { return $true }
    } catch {}
    return $false
}

function Find-Python {
    if (Get-Command py -ErrorAction SilentlyContinue) {
        if (Test-Python "py" @("-3")) { return @{exe="py"; pre=@("-3")} }
    }
    $cmd = Get-Command python -ErrorAction SilentlyContinue
    if ($cmd -and $cmd.Source -notlike "*WindowsApps*") {
        if (Test-Python $cmd.Source @()) { return @{exe=$cmd.Source; pre=@()} }
    }
    foreach ($d in Get-ChildItem "$env:LOCALAPPDATA\Programs\Python" -Directory -ErrorAction SilentlyContinue | Sort-Object Name -Descending) {
        $p = Join-Path $d.FullName "python.exe"
        if ((Test-Path $p) -and (Test-Python $p @())) { return @{exe=$p; pre=@()} }
    }
    return $null
}

try {
    Step "Looking for Python 3.9+"
    $py = Find-Python
    if (-not $py) {
        Write-Host "Python not found - installing Python 3.12 for your user account via winget..."
        winget install -e --id Python.Python.3.12 --scope user --silent --accept-package-agreements --accept-source-agreements
        $py = Find-Python
        if (-not $py) { throw "Python install failed. Install Python 3.12 from python.org, then run install.bat again." }
    }
    Write-Host "Using: $($py.exe) $($py.pre -join ' ')"

    Step "Creating private environment in $venv"
    New-Item -ItemType Directory -Force -Path $dataDir | Out-Null
    if (-not (Test-Path (Join-Path $venv "Scripts\python.exe"))) {
        & $py.exe @($py.pre) -m venv $venv
        if ($LASTEXITCODE -ne 0) { throw "Could not create virtual environment." }
    }
    $vpy = Join-Path $venv "Scripts\python.exe"

    Step "Installing conversion libraries (first run takes a minute or two)"
    & $vpy -m pip install --upgrade pip --quiet --disable-pip-version-check
    & $vpy -m pip install -r (Join-Path $here "requirements.txt") --quiet --disable-pip-version-check
    if ($LASTEXITCODE -ne 0) { throw "pip install failed - see messages above." }

    Step "Registering the right-click menu"
    & $vpy (Join-Path $here "convert.py") --install
    if ($LASTEXITCODE -ne 0) { throw "Menu registration failed." }

    Step "Checking conversion engines"
    & $vpy (Join-Path $here "convert.py") --check

    Write-Host "`nDONE. Right-click a file -> Convert to -> PDF / JPEG image / Word document." -ForegroundColor Green
    Write-Host "On Windows 11 it's under 'Show more options' (or Shift + right-click)."
}
catch {
    Write-Host "`nINSTALL FAILED: $_" -ForegroundColor Red
}
Write-Host ""
Read-Host "Press Enter to close"
