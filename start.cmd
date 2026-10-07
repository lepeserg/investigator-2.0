@echo off
setlocal
cd /d "%~dp0"
powershell.exe -NoProfile -ExecutionPolicy Bypass -File "%~dp0setup-windows.ps1" %*
set "setup_result=%errorlevel%"
if not "%setup_result%"=="0" echo Setup incomplete or cancelled. See the message above and README.md.
pause
exit /b %setup_result%
