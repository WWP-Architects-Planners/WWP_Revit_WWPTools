@echo off
setlocal

set SCRIPT_DIR=%~dp0
powershell -NoProfile -ExecutionPolicy Bypass -File "%SCRIPT_DIR%Update_DynamoPackages.ps1"

if %errorlevel% neq 0 (
  echo.
  echo Dynamo packages install failed. Please contact support.
  exit /b %errorlevel%
)

echo.
echo Dynamo packages install/update complete.
pause

