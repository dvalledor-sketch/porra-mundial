@echo off
REM Genera plantilla_eliminatorias.xlsx con los emparejamientos reales.
REM Usa este script TRAS LA FASE DE GRUPOS, cuando ya se conocen los cruces.
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
%PY% scripts\build_templates.py --fase elim
echo OK Plantilla de eliminatorias generada en plantillas\plantilla_eliminatorias.xlsx
echo    Distribuye este archivo. Los participantes deben guardar su copia como:
echo    porra_SUNOMBRE_elim.xlsx  en la carpeta pronosticos\
pause
