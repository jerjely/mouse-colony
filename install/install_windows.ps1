# Mouse Colony - Windows installer
# Copies the program to %LOCALAPPDATA%\MouseColony, creates a private Python environment there,
# installs openpyxl, puts the workbook in Documents\Mouse colony (only if it is not there yet)
# and adds Desktop + Start-menu shortcuts. Run via install_windows.bat (double-click).

$ErrorActionPreference = 'Stop'
$src  = Split-Path -Parent $PSScriptRoot          # the unzipped MouseColony folder
$root = Join-Path $env:LOCALAPPDATA 'MouseColony'
$app  = Join-Path $root 'app'
$venv = Join-Path $root 'venv'
$docs = Join-Path ([Environment]::GetFolderPath('MyDocuments')) 'Mouse colony'

Write-Host ''
Write-Host '=== Mouse Colony - installer ===' -ForegroundColor Cyan
Write-Host "Program folder: $root"
Write-Host ''

function Test-Python($cmd) {
    # Python 3.8+ that also has Tkinter (the window toolkit)
    $exe = $cmd[0]
    $rest = @()
    if ($cmd.Count -gt 1) { $rest = $cmd[1..($cmd.Count - 1)] }
    try {
        $v = & $exe @rest -c "import sys, tkinter; print('%d.%d' % sys.version_info[:2]); sys.exit(0 if sys.version_info >= (3, 8) else 1)" 2>$null
        if ($LASTEXITCODE -eq 0) { return $v }
    } catch {}
    return $null
}

# 1. find Python (python.org launcher first, then PATH, then Anaconda/Miniconda)
$candidates = @()
if (Get-Command py -ErrorAction SilentlyContinue)     { $candidates += ,@('py', '-3') }
if (Get-Command python -ErrorAction SilentlyContinue) { $candidates += ,@('python') }
foreach ($p in @("$env:USERPROFILE\anaconda3\python.exe", "$env:USERPROFILE\miniconda3\python.exe",
                 "$env:ProgramData\anaconda3\python.exe", "$env:ProgramData\miniconda3\python.exe",
                 "$env:LOCALAPPDATA\anaconda3\python.exe", "$env:LOCALAPPDATA\miniconda3\python.exe")) {
    if (Test-Path $p) { $candidates += ,@($p) }
}
$py = $null
foreach ($c in $candidates) {
    $ver = Test-Python $c
    if ($ver) { $py = $c; Write-Host "Using Python $ver  ($($c -join ' '))"; break }
}
if (-not $py) {
    Write-Host 'No Python 3.8+ with Tkinter found.' -ForegroundColor Red
    Write-Host 'Install it from https://www.python.org/downloads/ (tick "Add python.exe to PATH";'
    Write-Host 'Tkinter is included by default), then double-click install_windows.bat again.'
    exit 1
}
$pyExe = $py[0]
$pyArgs = @()
if ($py.Count -gt 1) { $pyArgs = $py[1..($py.Count - 1)] }

# 2. copy the program
New-Item -ItemType Directory -Force -Path $app | Out-Null
foreach ($f in @('mouse_colony_tool.py', 'mousecolony.ico', 'mousecolony.png', 'requirements.txt', 'README.md', 'CAGE_CARDS_FORMAT.md', 'GOOGLE_SETUP.md', 'GEMINI_CARD_INSTRUCTIONS.txt')) {
    Copy-Item -Force (Join-Path $src $f) $app
}
Copy-Item -Force (Join-Path $PSScriptRoot 'uninstall_windows.ps1') $root

# 3. private environment + openpyxl
$vpy = Join-Path $venv 'Scripts\python.exe'
if (-not (Test-Path $vpy)) {
    Write-Host 'Creating the Python environment...'
    & $pyExe @pyArgs -m venv $venv
    if ($LASTEXITCODE -ne 0) { throw 'Could not create the Python environment.' }
}
Write-Host 'Installing packages (small, under a minute)...'
& $vpy -m pip install --disable-pip-version-check --upgrade pip | Out-Null
& $vpy -m pip install --disable-pip-version-check -r (Join-Path $app 'requirements.txt')
if ($LASTEXITCODE -ne 0) {
    Write-Host ''
    Write-Host 'Package installation failed.' -ForegroundColor Red
    Write-Host 'If the error mentions SSL and you use Anaconda, install Python from python.org and run'
    Write-Host 'the installer again (delete the folder above first).'
    exit 1
}
& $vpy -c "import tkinter, openpyxl, gspread, google_auth_oauthlib, PIL; print('All packages OK')"
# optional: drag & drop of files onto the window (the app works without it)
& $vpy -m pip install --disable-pip-version-check -q tkinterdnd2 2>$null
if ($LASTEXITCODE -ne 0) { Write-Host 'Drag & drop add-on not installed - use the Import button instead.' }
if ($LASTEXITCODE -ne 0) { throw 'Package check failed.' }

# 4. the workbook goes to Documents - never overwritten by an update
$wbSrc = Join-Path $src 'Volgyi_mice_colony.xlsx'
$wbDst = Join-Path $docs 'Volgyi_mice_colony.xlsx'
if (Test-Path $wbDst) {
    Write-Host "Workbook already in $docs - left unchanged."
} elseif (Test-Path $wbSrc) {
    New-Item -ItemType Directory -Force -Path $docs | Out-Null
    Copy-Item $wbSrc $wbDst
    Write-Host "Workbook copied to $wbDst"
}

# 5. shortcuts
$ws = New-Object -ComObject WScript.Shell
$targets = @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))
foreach ($dir in $targets) {
    $lnk = $ws.CreateShortcut((Join-Path $dir 'Mouse Colony.lnk'))
    $lnk.TargetPath       = Join-Path $venv 'Scripts\pythonw.exe'
    $lnk.Arguments        = '"' + (Join-Path $app 'mouse_colony_tool.py') + '"'
    $lnk.WorkingDirectory = $app
    $lnk.IconLocation     = Join-Path $app 'mousecolony.ico'
    $lnk.Description      = 'Mouse colony - breeder ages, relatedness, pair suggestions'
    $lnk.Save()
}
# a console variant that shows error messages, for troubleshooting
$lnk = $ws.CreateShortcut((Join-Path $targets[1] 'Mouse Colony (console).lnk'))
$lnk.TargetPath       = $vpy
$lnk.Arguments        = '"' + (Join-Path $app 'mouse_colony_tool.py') + '"'
$lnk.WorkingDirectory = $app
$lnk.IconLocation     = Join-Path $app 'mousecolony.ico'
$lnk.Save()

Write-Host ''
Write-Host 'Done. Start "Mouse Colony" from the Desktop or the Start menu.' -ForegroundColor Green
Write-Host 'It opens the workbook in Documents\Mouse colony; use "Open workbook..." to switch to a'
Write-Host 'shared copy (e.g. OneDrive or a network drive) - it remembers the last one opened.'
Write-Host 'To update later: unzip the new version and run install_windows.bat again.'
