@echo off
setlocal

set "script_dir=%~dp0"
call "%script_dir%installer\Build_WWPTools_MSI.bat"
exit /b %ERRORLEVEL%
