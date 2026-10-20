@echo off
REM ============================================================
REM  College Football Model - data updater only (no Excel)
REM  You normally don't need this: "Open CFB Model.bat" / the
REM  desktop shortcut updates automatically before opening.
REM
REM  Runs at most once per day to save API credits. To force
REM  another update the same day:  python update_data.py --force
REM ============================================================

cd /d "%~dp0"

REM Finds Python, or installs a private copy on first run
call "%~dp0setup_python.bat"
if not defined PY (
    pause
    exit /b 1
)

"%PY%" update_data.py %*

echo.
pause
