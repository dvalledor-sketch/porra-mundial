#!/usr/bin/env bash
# Regenera plantilla_porra.xlsx (Fase Grupos + Premios) desde matches.json
# Usa este script ANTES del torneo para distribuir la plantilla inicial.
set -e
cd "$(dirname "$0")"
PY="$(command -v python3 || command -v python)"
[ -z "$PY" ] && { echo "✘ Falta Python"; read -p "Enter…" _; exit 1; }
"$PY" -m pip install --quiet --user --break-system-packages openpyxl >/dev/null 2>&1 || true
"$PY" scripts/build_matches.py
"$PY" scripts/build_templates.py --fase grupos
echo "✔ Plantilla de grupos → plantillas/plantilla_porra.xlsx"
read -p "Pulsa Enter para cerrar… " _
