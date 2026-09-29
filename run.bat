@echo off
rem Launch the CMC Report Automation desktop app on Windows.
cd /d "%~dp0"
if exist "%~dp0.venv\Scripts\python.exe" (
    "%~dp0.venv\Scripts\python.exe" -m app.main %*
) else (
    python -m app.main %*
)
if errorlevel 1 pause
