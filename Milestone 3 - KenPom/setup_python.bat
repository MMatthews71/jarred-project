@echo off
REM ============================================================
REM  Finds (or installs) a Python that can run update_data.py.
REM  Called by "Open CFB Model.bat" and "UpdateData.bat" - you
REM  don't need to run it yourself.
REM
REM  On a PC with no Python, the first run downloads a private
REM  copy of Python (~10 MB) into the "python" folder next to
REM  this file, plus the packages it needs. Nothing is installed
REM  system-wide and no admin rights are needed. Later runs reuse
REM  that copy and start instantly.
REM
REM  On success, the variable PY holds the python.exe to use.
REM  To start over, delete the "python" folder.
REM ============================================================

set "PYVER=3.12.10"
set "PKGS=requests openpyxl tzdata kenpompy"
set "PYDIR=%~dp0python"
set "PY="

REM 1. Private copy from an earlier run
if exist "%PYDIR%\python.exe" (
    set "PY=%PYDIR%\python.exe"
    goto check_local
)

REM 2. An existing Python that already has the packages
for %%C in ("%USERPROFILE%\anaconda3\python.exe" "%USERPROFILE%\miniconda3\python.exe") do (
    if exist %%C (
        %%C -c "import requests, openpyxl" >nul 2>&1 && (set "PY=%%~C" & goto :eof)
    )
)
python -c "import requests, openpyxl" >nul 2>&1 && (set "PY=python" & goto :eof)

REM 3. Nothing usable - install a private copy
echo.
echo ------------------------------------------------------------
echo  First-time setup: downloading Python %PYVER% and the
echo  packages this model needs. This takes a minute or two and
echo  only happens once. Please keep this window open.
echo ------------------------------------------------------------
echo.

set "ARCH=amd64"
if /i "%PROCESSOR_ARCHITECTURE%"=="ARM64" set "ARCH=arm64"
if /i "%PROCESSOR_ARCHITECTURE%"=="x86" if not defined PROCESSOR_ARCHITEW6432 set "ARCH=win32"

set "ZIP=%TEMP%\cfb-python-%PYVER%.zip"
set "GETPIP=%TEMP%\cfb-get-pip.py"
set "PS=powershell -NoProfile -ExecutionPolicy Bypass -Command"
set "DL=[Net.ServicePointManager]::SecurityProtocol='Tls12'; $ProgressPreference='SilentlyContinue'; Invoke-WebRequest -UseBasicParsing"

echo [1/4] Downloading Python...
%PS% "%DL% -Uri 'https://www.python.org/ftp/python/%PYVER%/python-%PYVER%-embed-%ARCH%.zip' -OutFile '%ZIP%'"
if errorlevel 1 goto fail

echo [2/4] Unpacking...
%PS% "Expand-Archive -Force -LiteralPath '%ZIP%' -DestinationPath '%PYDIR%'"
if errorlevel 1 goto fail
del "%ZIP%" >nul 2>&1
REM The embedded Python ignores site-packages until "import site" is enabled
%PS% "Get-ChildItem -LiteralPath '%PYDIR%' -Filter 'python*._pth' | ForEach-Object { (Get-Content -LiteralPath $_.FullName) -replace '^#\s*import site','import site' | Set-Content -LiteralPath $_.FullName }"
if errorlevel 1 goto fail

echo [3/4] Installing pip...
%PS% "%DL% -Uri 'https://bootstrap.pypa.io/get-pip.py' -OutFile '%GETPIP%'"
if errorlevel 1 goto fail
"%PYDIR%\python.exe" "%GETPIP%" --no-warn-script-location -q
REM get-pip can exit 0 even when its download failed, so check pip directly
"%PYDIR%\python.exe" -m pip --version >nul 2>&1
if errorlevel 1 goto fail
del "%GETPIP%" >nul 2>&1

echo [4/4] Installing packages...
"%PYDIR%\python.exe" -m pip install --no-warn-script-location -q %PKGS%
if errorlevel 1 goto fail

set "PY=%PYDIR%\python.exe"
echo.
echo Setup complete.
echo.
goto :eof

:check_local
REM Repair a half-finished install (e.g. the window was closed mid-setup)
"%PY%" -c "import requests, openpyxl, tzdata" >nul 2>&1 && goto :eof
if not exist "%PYDIR%\Scripts\pip.exe" if not exist "%PYDIR%\Lib\site-packages\pip" (
    rmdir /s /q "%PYDIR%"
    set "PY="
    goto :eof_retry
)
echo Installing missing packages...
"%PY%" -m pip install --no-warn-script-location -q %PKGS%
if errorlevel 1 goto fail
goto :eof

:eof_retry
call "%~f0"
goto :eof

:fail
echo.
echo ------------------------------------------------------------
echo  Automatic Python setup failed ^(see above^). Check the
echo  internet connection and try again. If it keeps failing,
echo  delete the "python" folder next to this file and retry.
echo ------------------------------------------------------------
set "PY="
exit /b 1
