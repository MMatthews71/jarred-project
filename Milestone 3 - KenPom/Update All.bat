@echo off
REM ============================================================
REM  College Football + Basketball - data updater
REM  Double-click this to refresh BOTH tabs of
REM  "College Sports.xlsx" (football first, then basketball).
REM  It only runs when you click it.
REM  Close the workbook in Excel first.
REM ============================================================

cd /d "%~dp0"

call "%~dp0setup_python.bat"
if not defined PY (
    pause
    exit /b 1
)

echo ===== FOOTBALL =====
"%PY%" update_football.py %*
echo.
echo ===== BASKETBALL =====
"%PY%" update_basketball.py %*

echo.
pause
