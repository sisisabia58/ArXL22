@echo off
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-EXLeratePerUser.ps1" -SkipCopy -SkipArp -Silent
exit /b %ERRORLEVEL%
