@echo off
rem Removes the program, its Python environment and the shortcuts. The workbook is not touched.
powershell -NoProfile -ExecutionPolicy Bypass -File "%~dp0install\uninstall_windows.ps1"
pause
