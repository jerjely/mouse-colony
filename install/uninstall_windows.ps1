# Removes %LOCALAPPDATA%\MouseColony, the settings and the shortcuts.
# The workbook (Documents\Mouse colony) and its backups are NOT touched.
$root = Join-Path $env:LOCALAPPDATA 'MouseColony'
foreach ($dir in @([Environment]::GetFolderPath('Desktop'), [Environment]::GetFolderPath('Programs'))) {
    foreach ($n in @('Mouse Colony.lnk', 'Mouse Colony (console).lnk')) {
        $p = Join-Path $dir $n
        if (Test-Path $p) { Remove-Item -Force $p }
    }
}
$cfg = Join-Path $env:APPDATA 'MouseColony'
if (Test-Path $cfg) { Remove-Item -Recurse -Force $cfg }
if (Test-Path $root) { Remove-Item -Recurse -Force $root }
Write-Host 'Mouse Colony removed. Your workbook was not touched.'
