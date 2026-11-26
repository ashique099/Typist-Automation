@echo off
REM Typist Automation - office PC par website chalane ke liye.
REM Pehli baar: Python 3.12 install karein (python.org), "Add to PATH" tick karein.
cd /d %~dp0
if not exist .venv (
  python -m venv .venv
  .venv\Scripts\pip install -r requirements.txt
)
if not exist settings.bat (
  echo set ANTHROPIC_API_KEY=yahan-apni-key> settings.bat
  echo set APP_PASSWORD=team-ka-password>> settings.bat
  echo settings.bat ban gayi hai. Usme API key aur password daal kar dobara chalayein.
  notepad settings.bat
  exit /b
)
call "%~dp0settings.bat"
echo Website chal rahi hai: http://localhost:8000  (doosre computer: http://IS-PC-KA-IP:8000)
.venv\Scripts\python -m uvicorn web.app:app --host 0.0.0.0 --port 8000
