@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal
set "PYTHONUTF8=1"
set "PYTHONIOENCODING=utf-8"
set "RC=0"

echo ============================================================
echo   AI4Word deep_test  -  parallel slots (default 3)
echo   each slot: own Word instance / own log / own settings
echo   global hotkey + autostart registry run only on slot 0
echo ============================================================
echo NOTE 1: do NOT touch mouse/keyboard while running (~10-25 min).
echo          multiple windows will be moved/clicked for real.
echo NOTE 2: quit any running AI4Word first (single-instance lock).
echo NOTE 3: needs ATRIA_API_KEY in .env.
echo NOTE 4: each slot opens its own Word (~200-400MB each).
echo          use --slots 2 if low on memory.
echo.
echo args (passed through): --slots N / --only FILTER / --seed N / --skip A,B / --timeout S
echo.

set "PY="
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"
if not defined PY if exist ".venv\python.exe" set "PY=.venv\python.exe"
if not defined PY echo [env error] python.exe NOT FOUND under .venv\ or .venv\Scripts\ - create .venv per README (conda create -p .venv python=3.11) and pip install -r requirements.txt
if not defined PY set "RC=2"
if defined PY "%PY%" -u tests\deep_parallel.py %*
if defined PY set "RC=%ERRORLEVEL%"

echo.
echo ============================================================
echo   finished, exit code %RC%    (0=clean 1=findings 2=env error)
echo ============================================================
echo merged results: tests\deep_runs\ - latest folder - triage.txt + summary.json
echo per-slot logs : that folder's slotN subfolders (debug.log, console.log)
echo.

echo --- press any key to close this window ---
pause >nul
endlocal & exit /b %RC%
