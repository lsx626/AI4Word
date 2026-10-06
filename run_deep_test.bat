@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "RC=0"

echo ============================================================
echo   AI4Word deep_test  (single process, fully real driving)
echo   real window / real keys+clicks / real Word COM / real Atria
echo ============================================================
echo NOTE 1: do NOT touch mouse/keyboard while running (10-40 min).
echo NOTE 2: quit any running AI4Word first (single-instance lock).
echo NOTE 3: needs ATRIA_API_KEY in .env.
echo.
echo args (passed through): --only FILTER / --seed N / --skip A,B / --keep-word
echo.

set "PY="
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if not defined PY if exist ".venv\python.exe" set "PY=.venv\python.exe"
if not defined PY echo [env error] python.exe NOT FOUND under .venv\ or .venv\Scripts\ - create .venv per README (conda create -p .venv python=3.11) and pip install -r requirements.txt
if not defined PY set "RC=2"
if defined PY "%PY%" -u tests\deep_test.py %*
if defined PY set "RC=%ERRORLEVEL%"

echo.
echo ============================================================
echo   finished, exit code %RC%    (0=clean 1=findings 2=env error)
echo ============================================================
echo results in tests\deep_runs\ - latest folder - triage.txt + summary.json
echo.

echo --- press any key to close this window ---
pause >nul
endlocal & exit /b %RC%
