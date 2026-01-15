@echo off
setlocal

pushd "%~dp0"
PowerShell -ExecutionPolicy Bypass -File "Build_WWPTools_MSI.ps1"
set "exit_code=%ERRORLEVEL%"
popd

exit /b %exit_code%
