# Despliegue en GitHub + Cloudflare Pages

## Resumen del flujo

```
Participantes  →  descargan plantilla desde la web
               →  rellenan y mandan Excel a Daniel por WhatsApp/email

Daniel         →  copia el .xlsx a la carpeta pronosticos/
               →  hace git push

GitHub Actions →  detecta el push → ejecuta scorer.py → actualiza data.json
               →  cada hora: llama a API-Football → actualiza resultados → recalcula

Cloudflare     →  detecta el commit → redespliega la web en segundos
```

---

## Paso 1 — Crear el repositorio en GitHub

1. Ve a https://github.com/new
2. Nombre sugerido: `porra-mundial-2026`
3. Visibilidad: **Private** (para que nadie vea los pronósticos antes de tiempo)
4. Sin README ni .gitignore (ya existen)

Desde la carpeta del proyecto en tu Mac:

```bash
cd "/Users/daniel/Documents/Porra Mundial"
git init
git add .
git commit -m "primer commit"
git branch -M main
git remote add origin https://github.com/TU_USUARIO/porra-mundial-2026.git
git push -u origin main
```

---

## Paso 2 — Obtener la API key gratuita de API-Football

1. Regístrate en https://dashboard.api-football.com/register (gratis)
2. Ve a **My Account → API Keys**
3. Copia tu clave (formato `xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx`)

---

## Paso 3 — Añadir el secret al repositorio

En GitHub: **Settings → Secrets and variables → Actions → New repository secret**

- Nombre: `FOOTBALL_API_KEY`
- Valor: tu clave de API-Football

---

## Paso 4 — Conectar Cloudflare Pages

1. Ve a https://dash.cloudflare.com → **Workers & Pages → Create → Pages**
2. **Connect to Git** → elige tu repositorio `porra-mundial-2026`
3. Configuración de build:
   - **Framework preset**: None
   - **Build command**: *(dejar vacío)*
   - **Build output directory**: `dashboard`
4. Haz clic en **Save and Deploy**

Cloudflare te dará una URL del tipo `porra-mundial-2026.pages.dev`.
Cada vez que GitHub Actions hace un commit, Cloudflare redespliega automáticamente.

---

## Paso 5 — (Opcional) Dominio personalizado

En Cloudflare Pages → tu proyecto → **Custom domains → Add domain**  
Introduce tu dominio (ej. `porra.sonamovil.com`) y sigue los pasos.

---

## Flujo diario durante el torneo

**Cuando recibes el Excel de un participante:**
```bash
cp ~/Downloads/plantilla_Antonio.xlsx "/Users/daniel/Documents/Porra Mundial/pronosticos/"
cd "/Users/daniel/Documents/Porra Mundial"
git add pronosticos/plantilla_Antonio.xlsx
git commit -m "pronóstico Antonio"
git push
```
→ GitHub Actions recalcula en ~30 segundos → web actualizada.

**Los resultados se actualizan solos** cada hora gracias al workflow `update-results.yml`.
También puedes lanzarlo a mano desde GitHub → Actions → "Actualizar resultados automáticamente" → Run workflow.

---

## Estructura de archivos relevante

```
.github/
  workflows/
    update-results.yml   ← cron hourly: fetch API + recalculate
    recalculate.yml      ← on push: recalculate scorer

scripts/
  fetch_results.py       ← llama a API-Football, actualiza resultados_reales.xlsx
  scorer.py              ← calcula puntos y genera dashboard/data.json

data/
  matches.json           ← definición de todos los partidos
  resultados_reales.xlsx ← resultados oficiales (auto-rellenado por fetch_results.py)

pronosticos/
  plantilla_Antonio.xlsx ← pronósticos de cada participante
  plantilla_DaniV.xlsx
  ...

dashboard/               ← lo que sirve Cloudflare Pages
  dashboard.html         ← la web (como index.html)
  data.json              ← datos generados por scorer.py
  downloads/
    plantilla_porra.xlsx ← plantilla para que descarguen los participantes
```

> **Nota**: Cloudflare Pages busca `index.html` en la carpeta `dashboard/`.  
> Renombra `dashboard.html` a `index.html` antes del primer push:
> ```bash
> mv dashboard/dashboard.html dashboard/index.html
> ```
> Y actualiza la línea `open dashboard/dashboard.html` en `Actualizar_Ranking.command`
> por `open dashboard/index.html`.
