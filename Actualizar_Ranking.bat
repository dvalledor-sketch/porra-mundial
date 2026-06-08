@echo off
REM ─────────────────────────────────────────────────────────────
REM   Porra del Mundial 2026 — Windows
REM   Doble click para procesar pronosticos y abrir el dashboard.
REM ─────────────────────────────────────────────────────────────
setlocal
cd /d "%~dp0"

set "PY="
where py >nul 2>&1 && set "PY=py -3"
if "%PY%"=="" where python >nul 2>&1 && set "PY=python"
if "%PY%"=="" (
  echo X No se encontro Python. Instalalo desde https://www.python.org/downloads/
  pause & exit /b 1
)

echo - Instalando dependencias (solo la primera vez)...
%PY% -m pip install --quiet --user openpyxl

echo - Generando dataset de partidos...
%PY% scripts\build_matches.py

echo - Calculando puntos...
%PY% scripts\scorer.py

echo - Abriendo dashboard en el navegador...
start "" "dashboard\dashboard.html"

echo OK. Cierra esta ventana cuando termines.
pause
