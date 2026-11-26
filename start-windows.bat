@echo off
REM Typist Automation - office PC par website chalane ke liye.
REM Pehli baar: Python 3.12 install karein (python.org), "Add to PATH" tick karein.
cd /d %~dp0
if not exist .venv (
  python -m venv .venv
)
REM har baar: nayi requirements (agar update aaya ho) install ho jayengi
.venv\Scripts\pip install -q -r requirements.txt
if not exist settings.bat (
  echo set ANTHROPIC_API_KEY=yahan-apni-key> settings.bat
  echo settings.bat ban gayi hai. Usme API key daal kar dobara chalayein.
  notepad settings.bat
  exit /b
)
call "%~dp0settings.bat"
echo Website chal rahi hai: http://localhost:8000  (doosre computer: http://IS-PC-KA-IP:8000)
echo Pehli baar khulne par admin account banayein.
.venv\Scripts\python -m uvicorn web.app:app --host 0.0.0.0 --port 8000
