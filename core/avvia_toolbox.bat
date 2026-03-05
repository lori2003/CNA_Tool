@echo off
setlocal

REM Risolve ROOT = cartella padre di core\ (la radice di toolbox)
for %%i in ("%~dp0..") do set "ROOT=%%~fi"

set "TOOLBOX_HOME=%ROOT%"
set "VENV_PY=%ROOT%\.venv\Scripts\python.exe"
set "LOG=%~dp0toolbox.log"
set "SERVER=%ROOT%\server.py"

echo [Toolbox] Avvio FastAPI... > "%LOG%"

REM 0) Check server.py
if not exist "%SERVER%" (
  echo [ERRORE] Non trovo server.py in: %ROOT% >> "%LOG%"
  start notepad "%LOG%"
  exit /b 1
)

REM 1) Create venv if missing
if not exist "%VENV_PY%" (
  echo [Toolbox] Creating venv in "%ROOT%\.venv"... >> "%LOG%"
  where py >nul 2>&1
  if %errorlevel%==0 (
    py -m venv "%ROOT%\.venv" >> "%LOG%" 2>&1
  ) else (
    python -m venv "%ROOT%\.venv" >> "%LOG%" 2>&1
  )
)

REM 2) Install deps
echo [Toolbox] Installing/Updating deps... >> "%LOG%"
"%VENV_PY%" -m pip install --upgrade pip >> "%LOG%" 2>&1

if exist "%ROOT%\requirements_web.txt" (
  "%VENV_PY%" -m pip install -r "%ROOT%\requirements_web.txt" >> "%LOG%" 2>&1
) else (
  "%VENV_PY%" -m pip install fastapi uvicorn[standard] python-multipart openpyxl pandas >> "%LOG%" 2>&1
)

REM 3) Esegui FastAPI server dalla ROOT
echo [Toolbox] Starting FastAPI server... >> "%LOG%"
cd /d "%ROOT%"
"%VENV_PY%" server.py >> "%LOG%" 2>&1

REM Se il server esce apri log
start notepad "%LOG%"
exit /b 1
