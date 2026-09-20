@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-EXLeratePerUser.ps1" -KeepFiles -Silent
exit /b 0
