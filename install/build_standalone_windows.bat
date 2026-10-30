@echo off
rem Builds a standalone "MouseColony" folder (with MouseColony.exe) that runs on PCs without Python.
rem Run install_windows.bat first. Output: dist\MouseColony  - zip that folder and share it.
set VPY=%LOCALAPPDATA%\MouseColony\venv\Scripts\python.exe
if not exist "%VPY%" (echo Run install_windows.bat first. & pause & exit /b 1)
"%VPY%" -m pip install --disable-pip-version-check pyinstaller
cd /d "%~dp0.."
"%VPY%" -m PyInstaller --noconfirm --distpath dist --workpath build install\mousecolony.spec
if errorlevel 1 (echo Build failed. & pause & exit /b 1)
copy /y Volgyi_mice_colony.xlsx dist\MouseColony\ >nul
echo.
echo Built: %CD%\dist\MouseColony\MouseColony.exe
pause
