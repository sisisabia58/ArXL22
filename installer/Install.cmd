@echo off
setlocal
cd /d "%~dp0"
echo EXLerate Explorer - per-user install (no admin)
echo.
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0Install-EXLeratePerUser.ps1" -PayloadDir "%~dp0"
if errorlevel 1 (
  echo.
  echo Install failed.
  pause
  exit /b 1
)
echo.
pause
