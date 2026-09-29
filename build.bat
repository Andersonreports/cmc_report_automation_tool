@echo off
rem Build the standalone Windows exe locally.
rem Just double-click this file (or run it from a terminal). The window stays
rem open at the end so you can read the result even on failure.
cd /d "%~dp0"

rem Use the project's virtual environment if one exists, otherwise fall back to
rem the Python on PATH. Using "%PY% -m ..." guarantees pip/PyInstaller run under
rem the very same interpreter (not a different one that happens to be on PATH).
set "PY=python"
if exist ".venv\Scripts\python.exe" set "PY=.venv\Scripts\python.exe"

echo Using interpreter: %PY%
"%PY%" --version
if errorlevel 1 goto :error

echo.
echo Installing dependencies (first run can take a minute)...
"%PY%" -m pip install --upgrade pip
"%PY%" -m pip install -r requirements.txt -r requirements-build.txt
if errorlevel 1 goto :error

echo.
echo Building the executable...
"%PY%" -m PyInstaller build.spec --noconfirm --clean
if errorlevel 1 goto :error

echo.
echo ============================================================
echo  Build complete:  dist\CMCReportAutomation.exe
echo  Distribute ONLY that .exe - never share the source with it.
echo ============================================================
echo.
pause
goto :eof

:error
echo.
echo ************************************************************
echo  BUILD FAILED - scroll up to see the cause.
echo ************************************************************
echo.
pause
exit /b 1
