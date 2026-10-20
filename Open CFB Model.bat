@echo off
REM ============================================================
REM  College Football Model - open with fresh data
REM  Pulls the latest numbers (at most once per day, to save API
REM  credits), then opens CFB_Betting_Model.xlsx in Excel.
REM  The "CFB Betting Model" desktop shortcut runs this file.
REM ============================================================

cd /d "%~dp0"
title CFB Betting Model - checking for new data

REM Finds Python, or installs a private copy on first run
call "%~dp0setup_python.bat"
if not defined PY (
    echo The workbook will open with the last saved numbers.
    pause
    goto open
)

"%PY%" update_data.py
if errorlevel 2 goto open
if errorlevel 1 (
    echo.
    echo ------------------------------------------------------------
    echo The data update hit a problem ^(see above^).
    echo The workbook will open with the last saved numbers.
    echo ------------------------------------------------------------
    pause
)

:open
start "" "%~dp0CFB_Betting_Model.xlsx"
