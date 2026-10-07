@echo off
setlocal
set "TASK_PYTHON=%~dp0.venv\Scripts\python.exe"
if not exist "%TASK_PYTHON%" (
  echo Run start.cmd first and install the base and audio profiles. See README.md.
  exit /b 2
)
"%TASK_PYTHON%" "%~dp0run.py" --offline setup_hf_access.py %*
set "TASK_EXIT=%ERRORLEVEL%"
if "%~1"=="" pause
exit /b %TASK_EXIT%
