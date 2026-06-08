#!/usr/bin/env bash
# ─────────────────────────────────────────────────────────────
#  Porra del Mundial 2026 — Mac
#  Doble click para procesar pronósticos y publicar en Netlify.
# ─────────────────────────────────────────────────────────────
set -e
cd "$(dirname "$0")"

# ── CONFIGURACIÓN ─────────────────────────────────────────────
NETLIFY_TOKEN="nfp_kp6cBifMAgrs9KkZdAkargBP4ryJacLm52ff"
NETLIFY_SITE_ID="porra-mundial-sdx8u"
# ──────────────────────────────────────────────────────────────

PY=""
for cand in python3 python; do
  if command -v "$cand" >/dev/null 2>&1; then PY="$cand"; break; fi
done
if [ -z "$PY" ]; then
  echo "✘ No se encontró Python. Instálalo desde https://www.python.org/downloads/"
  read -p "Pulsa Enter para cerrar… " _; exit 1
fi

echo "▸ Instalando dependencias (sólo la primera vez)…"
"$PY" -m pip install --quiet --user openpyxl requests >/dev/null 2>&1 || \
"$PY" -m pip install --quiet --user --break-system-packages openpyxl requests >/dev/null 2>&1 || true

echo "▸ Generando dataset de partidos…"
"$PY" scripts/build_matches.py

echo "▸ Calculando puntos…"
"$PY" scripts/scorer.py

echo "▸ Abriendo dashboard en el navegador…"
open dashboard/index.html

# ── Publicar en Netlify ────────────────────────────────────────
if [ "$NETLIFY_TOKEN" = "PEGA_AQUI_TU_TOKEN" ] || [ -z "$NETLIFY_TOKEN" ]; then
  echo "▸ (Netlify no configurado — salto la publicación)"
else
  echo "▸ Publicando en Netlify…"
  "$PY" - <<PYEOF
import requests, zipfile, io, os, random, string

token   = "$NETLIFY_TOKEN"
site_id = "$NETLIFY_SITE_ID".strip()
headers = {"Authorization": f"Bearer {token}"}

# Empaquetar dashboard.html como index.html
with open("dashboard/dashboard.html", "r", encoding="utf-8") as f:
    html = f.read()
html_bytes = html.encode("utf-8")
buf = io.BytesIO()
with zipfile.ZipFile(buf, "w", zipfile.ZIP_DEFLATED) as zf:
    zf.writestr("index.html", html_bytes)
    zf.writestr("_headers", "/index.html\n  Content-Type: text/html; charset=utf-8\n")
    zf.writestr("netlify.toml",
        '[[headers]]\n  for = "/*"\n  [headers.values]\n    Content-Type = "text/html; charset=utf-8"\n')
buf.seek(0)

# Crear sitio si no hay ID
if not site_id:
    sufijo = "".join(random.choices(string.ascii_lowercase + string.digits, k=5))
    r = requests.post(
        "https://api.netlify.com/api/v1/sites",
        headers={**headers, "Content-Type": "application/json"},
        json={"name": f"porra-mundial-{sufijo}"}
    )
    site = r.json()
    site_id = site["id"]
    print(f"  🆕 Sitio creado: {site['ssl_url']}")
    print(f"  ⚠️  Guarda este Site ID en el .command para la próxima vez:")
    print(f"  NETLIFY_SITE_ID=\"{site_id}\"")

# Deploy
r = requests.post(
    f"https://api.netlify.com/api/v1/sites/{site_id}/deploys",
    headers={**headers, "Content-Type": "application/zip"},
    data=buf.getvalue()
)
if r.status_code in (200, 201):
    url = r.json().get("ssl_url") or r.json().get("url", "")
    print(f"  ✅ Publicado: {url}")
else:
    print(f"  ❌ Error Netlify {r.status_code}: {r.text}")
PYEOF
fi

echo "✔ Listo. Cierra esta ventana cuando termines."
read -p "Pulsa Enter para salir… " _
