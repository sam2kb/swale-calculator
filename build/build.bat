@echo off
setlocal enabledelayedexpansion

for %%I in ("%~dp0..") do set "ROOT_DIR=%%~fI"

set "OUT_DIR=%ROOT_DIR%\build"
set "XLSX_FILE=%OUT_DIR%\swale_calculator.xlsx"
set "VENV_DIR=%ROOT_DIR%\.venv-windows"
set "PY=%VENV_DIR%\Scripts\python.exe"
set "DEPS_SENTINEL=%VENV_DIR%\.deps_installed_windows"

cd /d "%ROOT_DIR%"
if not exist "%OUT_DIR%" mkdir "%OUT_DIR%"

rem Create venv if missing
if not exist "%PY%" (
  where py >nul 2>&1
  if %errorlevel%==0 (
    py -3 -m venv "%VENV_DIR%"
  ) else (
    python -m venv "%VENV_DIR%"
  )
)

rem Install deps if needed (Windows sentinel + import check)
set "NEED_DEPS="
if not exist "%DEPS_SENTINEL%" set "NEED_DEPS=1"
"%PY%" -c "import openpyxl" >nul 2>&1 || set "NEED_DEPS=1"

if defined NEED_DEPS (
  "%PY%" -m pip install --upgrade pip >nul
  "%PY%" -m pip install -r "%ROOT_DIR%\requirements.txt"
  type nul > "%DEPS_SENTINEL%"
)

rem Generate workbook
"%PY%" "%ROOT_DIR%\swale-calculator.py" --out "%XLSX_FILE%"

if %errorlevel% neq 0 exit /b %errorlevel%
