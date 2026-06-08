# Porra del Mundial 2026 · Oficina

Sistema **multiplataforma** (Mac + Windows) para gestionar la porra del Mundial 2026 entre empleados, basado en archivos Excel en una **carpeta compartida** y un **dashboard HTML** que se actualiza con doble click.

---

## Estructura de la carpeta

```
porra-mundial/
├── Actualizar_Ranking.command          ← doble click en Mac (calcula puntos + abre dashboard)
├── Actualizar_Ranking.bat              ← doble click en Windows
├── Regenerar_Plantillas.command        ← Mac · regenera plantilla de GRUPOS
├── Regenerar_Plantillas.bat            ← Windows
├── Regenerar_Plantilla_Elim.command    ← Mac · genera plantilla ELIMINATORIAS (tras fase de grupos)
├── Regenerar_Plantilla_Elim.bat        ← Windows
├── plantillas/
│   ├── plantilla_porra.xlsx            ← FASE 1: Grupos + Premios (distribuir antes del torneo)
│   └── plantilla_eliminatorias.xlsx    ← FASE 2: Eliminatorias (distribuir tras fase de grupos)
├── pronosticos/                        ← cada empleado deja su .xlsx aquí
│   ├── porra_NOMBRE.xlsx               ← archivo de grupos (Fase 1)
│   └── porra_NOMBRE_elim.xlsx          ← archivo de eliminatorias (Fase 2, mismo alias)
├── data/
│   ├── matches.json                    ← maestro de los 104 partidos
│   ├── resultados_reales.xlsx          ← ADMIN: resultados de grupos
│   ├── resultados_reales_elim.xlsx     ← ADMIN: resultados de eliminatorias
│   ├── premios_reales.json             ← ADMIN: MVP, Goleador, Portero (al final del torneo)
│   ├── quien_paso.json                 ← ADMIN: equipo clasificado en empates a 90'
│   └── limites.json                    ← horas límite de cada fase
├── dashboard/
│   ├── index.html                      ← fuente del dashboard
│   ├── dashboard.html                  ← standalone generado por scorer.py (abrir con doble click)
│   └── data.json                       ← lo escribe el scorer
└── scripts/
    ├── build_matches.py                ← genera data/matches.json
    ├── build_templates.py              ← regenera plantillas Excel
    └── scorer.py                       ← calcula puntos y exporta dashboard/data.json
```

---

## Para los empleados (participantes)

La porra funciona en **dos fases**. Es obligatorio usar el **mismo alias** en las dos.

**Fase 1 — Grupos + Premios** (antes del torneo):
1. **Copia** `plantillas/plantilla_porra.xlsx` y renómbralo `porra_TUNOMBRE.xlsx`.
2. Escribe tu **alias** en la celda **C2** de las hojas *Fase Grupos* y *Premios*.
3. Rellena los goles en las **columnas amarillas** (D = local, E = visitante) para los 72 partidos. En la hoja *Premios*, escribe el nombre del jugador que crees ganará el MVP, el Goleador y el Portero del torneo.
4. **Guárdalo** en `pronosticos/` como `porra_TUNOMBRE.xlsx` antes del **11 jun 2026 · 17:00 h**.

**Fase 2 — Eliminatorias** (tras la fase de grupos):
1. Cuando el administrador distribuya `plantilla_eliminatorias.xlsx`, cópiala y renómbrala `porra_TUNOMBRE_elim.xlsx` (con el sufijo `_elim`, obligatorio).
2. Usa el **mismo alias** que en la Fase 1.
3. Rellena el marcador a los 90' en las columnas amarillas y el equipo clasificado en la columna verde **G (Quién Pasa)**.
4. **Guárdala** en `pronosticos/` antes de la hora límite de cada ronda.

Si subes el archivo varias veces vale el más reciente antes del límite. Un archivo guardado **después** de la hora límite tiene esa fase **invalidada automáticamente**.

> **Regla de oro para las eliminatorias**: anota el marcador al término de los 90′ + descuento. La prórroga y los penaltis **no cuentan** para el marcador. Rellena también la columna **G (Quién Pasa)** con el equipo que crees que se clasifica (+1 pt).

---

## Para el administrador

**Grupos:** abre `data/resultados_reales.xlsx`, busca el ID del partido (columna C) y escribe los goles en D y E.

**Eliminatorias:** igual pero en `data/resultados_reales_elim.xlsx`. Si el partido acabó en empate a 90' (fue a prórroga/penaltis), anota también el equipo clasificado en `data/quien_paso.json`.

Doble click en **`Actualizar_Ranking.command`** (Mac) o **`Actualizar_Ranking.bat`** (Windows) para recalcular y abrir el dashboard. El script instala `openpyxl` automáticamente la primera vez.

### Puntuación

| Concepto                                             | Puntos  |
| ---------------------------------------------------- | ------- |
| Pleno absoluto (resultado exacto a los 90')          | **3**   |
| Tendencia (mismo ganador o empate a los 90')         | **2**   |
| Fallo                                                | 0       |
| Quién pasa de ronda (solo eliminatorias, +1 extra)   | **+1**  |
| MVP del Torneo (pronosticado antes de grupos)        | **+15** |
| Goleador del Torneo                                  | **+15** |
| Portero del Torneo                                   | **+15** |

### Premios económicos

| Premio                    | % del bote | Ejemplo (20 pax · 200 €) |
| ------------------------- | ---------- | ------------------------- |
| Ganador Fase de Grupos    | **30 %**   | 60 €                      |
| Ganador Clasificación Final | **70 %** | 140 €                     |

### Desempate

En caso de empate en puntos, gana quien tenga más **plenos absolutos** (resultados exactos).

### Archivos extra del administrador

| Archivo                        | Cuándo rellenarlo                                     |
| ------------------------------ | ----------------------------------------------------- |
| `data/premios_reales.json`     | Al finalizar el torneo (MVP, Goleador, Portero)       |
| `data/quien_paso.json`         | Cuando un partido de eliminatorias acaba en empate a 90' (añade el equipo que pasó por prórroga/penaltis) |

---

## Personalización rápida

| Quiero…                                          | Edito…                                          |
| ------------------------------------------------ | ----------------------------------------------- |
| Cambiar un equipo de un grupo                    | `scripts/build_matches.py` → diccionario `GRUPOS`, luego doble click en *Regenerar_Plantillas* |
| Mover una hora límite                            | `data/limites.json`                             |
| Cambiar puntos (3/2/1/15)                        | `scripts/scorer.py` constantes `PUNTOS_*`       |
| Cambiar el aspecto del dashboard                 | `dashboard/index.html`                          |

---

## Dependencias

- **Python 3.9+** — se instala desde [python.org/downloads](https://www.python.org/downloads/) en Mac o Windows si no está.
- **openpyxl** — se instala automáticamente en la primera ejecución de los scripts.

No requiere base de datos, Google Drive ni servidor: todo vive en la carpeta compartida.

---

## Notas técnicas

- Los IDs de partido siguen la convención `G-{Letra}{n}-{Letra}{m}` para grupos (ej. `G-A1-A2` = México–Sudáfrica) y `R32-XX / R16-XX / QF-XX / SF-XX / TER / FIN` para eliminatorias.
- Los avatares pixel-art son **deterministas**: el mismo alias genera siempre el mismo personaje, sin guardar imágenes (algoritmo: FNV-1a 32-bit + mulberry32 PRNG sobre canvas 8×8 simétrico).
- El scorer respeta la regla del reglamento: el `modifiedTime` del archivo es la hora oficial de registro.
