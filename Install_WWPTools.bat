@echo off
setlocal

set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Update_WWPTools.ps1"

if %errorlevel% neq 0 (
  echo.
  echo Install failed. Please contact support.
  exit /b %errorlevel%
)

echo.
echo WWPTools install/update complete.
pause

