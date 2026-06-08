#!/usr/bin/env bash
# Genera plantilla_eliminatorias.xlsx con los emparejamientos reales.
# Usa este script TRAS LA FASE DE GRUPOS, cuando ya se conocen los cruces.
set -e
cd "$(dirname "$0")"
PY="$(command -v python3 || command -v python)"
[ -z "$PY" ] && { echo "✘ Falta Python"; read -p "Enter…" _; exit 1; }
"$PY" -m pip install --quiet --user --break-system-packages openpyxl >/dev/null 2>&1 || true
"$PY" scripts/build_templates.py --fase elim
echo "✔ Plantilla de eliminatorias → plantillas/plantilla_eliminatorias.xlsx"
echo "   Distribuye este archivo. Los participantes deben guardar su copia como:"
echo "   porra_SUNOMBRE_elim.xlsx  en la carpeta pronosticos/"
read -p "Pulsa Enter para cerrar… " _
