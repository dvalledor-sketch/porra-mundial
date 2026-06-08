@echo off
setlocal
cd /d "%~dp0"
set "PY="
where py >nul 2>&1 && set "PY=py -3"
if "%PY%"=="" where python >nul 2>&1 && set "PY=python"
if "%PY%"=="" (
  echo X Falta Python. Descargalo de https://www.python.org/downloads/
  pause & exit /b 1
)
%PY% -m pip install --quiet --user openpyxl
%PY% scripts\build_matches.py
%PY% scripts\build_templates.py --fase grupos
echo OK Plantilla de grupos regenerada en plantillas\plantilla_porra.xlsx
pause
