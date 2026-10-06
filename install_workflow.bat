@echo off
chcp 65001 >nul
cd /d "%~dp0"
setlocal
set "RC=0"

echo === install Claude Code workflow: .claude\workflows\deep-test.js ===
echo.

if not exist "claude-workflows\deep-test.js" echo [error] claude-workflows\deep-test.js NOT FOUND
if not exist "claude-workflows\deep-test.js" set "RC=2"
if exist "claude-workflows\deep-test.js" if not exist ".claude\workflows" mkdir ".claude\workflows"
if exist "claude-workflows\deep-test.js" copy /y "claude-workflows\deep-test.js" ".claude\workflows\deep-test.js" >nul
if exist ".claude\workflows\deep-test.js" echo installed: .claude\workflows\deep-test.js
if exist ".claude\workflows\deep-test.js" echo.
if exist ".claude\workflows\deep-test.js" echo next: start Claude Code in this repo (run: claude), then type:
if exist ".claude\workflows\deep-test.js" echo   /deep-test            full loop: test - analyze - fix - retest
if exist ".claude\workflows\deep-test.js" echo   /deep-test rounds=3   at most 3 fix rounds
if exist ".claude\workflows\deep-test.js" echo on first run confirm the workflow (Yes, and don't ask again to skip asks).
if not exist ".claude\workflows\deep-test.js" echo [error] install failed
if not exist ".claude\workflows\deep-test.js" set "RC=2"

echo.
echo --- press any key to close this window ---
pause >nul
endlocal
exit /b %RC%
