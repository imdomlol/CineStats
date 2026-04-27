@echo off
REM run.bat — Launch CineStats on Windows.
REM
REM Usage: double-click this file, or run it from PowerShell/Command Prompt.
REM Run setup.bat first if you haven't already.

SET SCRIPT_DIR=%~dp0
SET VENV=%SCRIPT_DIR%.venv\Scripts\python.exe

IF NOT EXIST "%VENV%" (
    echo Virtual environment not found.
    echo Please run setup.bat first.
    pause
    exit /b 1
)

"%VENV%" "%SCRIPT_DIR%src\main.py"
