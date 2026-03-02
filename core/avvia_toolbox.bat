@echo off
setlocal

REM Risolve ROOT = cartella padre di core\ (la radice di toolbox)
for %%i in ("%~dp0..") do set "ROOT=%%~fi"

REM TOOLBOX_HOME dice ad app.py dove trovare tools/, data/, .streamlit/
set "TOOLBOX_HOME=%ROOT%"

set "VENV_PY=%ROOT%\.venv\Scripts\python.exe"
set "LOG=%~dp0toolbox.log"
set "APP=%~dp0app.py"

echo [Toolbox] Avvio... > "%LOG%"

REM 0) Check app.py
if not exist "%APP%" (
  echo [ERRORE] Non trovo app.py in: %~dp0 >> "%LOG%"
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

if exist "%ROOT%\requirements.txt" (
  "%VENV_PY%" -m pip install -r "%ROOT%\requirements.txt" >> "%LOG%" 2>&1
) else (
  "%VENV_PY%" -m pip install streamlit openpyxl pandas >> "%LOG%" 2>&1
)

REM 3) Esegui Streamlit dalla ROOT cosi trova .streamlit/
echo [Toolbox] Starting Streamlit... >> "%LOG%"
cd /d "%ROOT%"
"%VENV_PY%" -m streamlit run "%APP%" >> "%LOG%" 2>&1

REM Se streamlit esce e' quasi sempre errore: apri log
start notepad "%LOG%"
exit /b 1
