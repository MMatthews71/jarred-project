@echo off
REM ============================================================
REM  College Basketball - data updater
REM  Double-click this to pull the latest numbers into
REM  the CBB tab of "College Sports.xlsx".  It only runs when
REM  you click it.
REM  Close the workbook in Excel first.
REM ============================================================

cd /d "%~dp0"

call "%~dp0setup_python.bat"
if not defined PY (
    pause
    exit /b 1
)

"%PY%" update_basketball.py %*

echo.
pause
