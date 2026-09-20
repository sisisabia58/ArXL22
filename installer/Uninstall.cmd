@echo off
setlocal
cd /d "%~dp0"
echo EXLerate Explorer - per-user uninstall
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Uninstall-EXLeratePerUser.ps1"
if errorlevel 1 (
  echo.
  echo Uninstall failed.
  pause
  exit /b 1
)
echo.
pause
