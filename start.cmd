@echo off
setlocal
cd /d "%~dp0"
py -3.12 bootstrap.py --profile base
if errorlevel 1 (
  echo Installation failed. Python 3.12 and internet access are required.
  pause
  exit /b 1
)
py -3.12 connect_skills.py
if errorlevel 1 (
  echo Skill connection failed. See README.md.
  pause
  exit /b 1
)
echo Base environment and project skills prepared. See README.md for local settings.
pause
