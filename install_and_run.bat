@echo off
title MailMerge Studio
cd /d "%~dp0"

where python >nul 2>nul
if errorlevel 1 (
  echo Python is not installed. Please install it from https://python.org and re-run.
  pause
  exit /b 1
)

if not exist ".venv" (
  echo First run: setting up MailMerge Studio...
  python -m venv .venv
  call .venv\Scripts\activate.bat
  pip install -q flask openpyxl
) else (
  call .venv\Scripts\activate.bat
)

start "" http://127.0.0.1:5000
python app.py
pause