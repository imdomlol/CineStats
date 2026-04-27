@echo off
REM setup.bat — One-time setup for CineStats on Windows.
REM
REM Creates a local virtual environment and installs dependencies.
REM No admin rights required.
REM
REM Usage: double-click this file, or run from PowerShell/Command Prompt.

SET SCRIPT_DIR=%~dp0

echo Creating virtual environment...
python -m venv "%SCRIPT_DIR%.venv"

echo Installing dependencies...
"%SCRIPT_DIR%.venv\Scripts\pip.exe" install --quiet -r "%SCRIPT_DIR%requirements.txt"

echo.
echo Setup complete. Run the app with: run.bat
pause
