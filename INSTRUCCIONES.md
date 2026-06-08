# Porra Mundial 2026 — Instrucciones de uso

---

## Para los participantes

La porra funciona en **dos fases**. Recibirás una plantilla distinta para cada una. Es imprescindible usar el **mismo alias** en las dos.

---

### FASE 1 — Grupos + Premios especiales

**1. Descarga tu plantilla**
Copia el archivo `plantillas/plantilla_porra.xlsx` y renómbralo `porra_TUNOMBRE.xlsx`.

**2. Escribe tu alias**
Abre el Excel y escribe tu nombre/alias en la celda **C2** de las dos hojas:
- *Fase Grupos*
- *Premios*

**3. Rellena tus pronósticos**

- **Fase Grupos:** escribe los goles previstos en las columnas amarillas (D = local, E = visitante) para los 72 partidos.
- **Premios:** escribe el nombre del jugador que crees ganará el MVP, el Goleador y el Portero del torneo (+15 pts cada acierto).

**4. Entrega el archivo**
Guarda tu archivo como `porra_TUNOMBRE.xlsx` en la carpeta compartida `pronosticos/` antes del:

> **11 de junio de 2026 · 17:00 h (hora España)**

---

### FASE 2 — Eliminatorias

Cuando acabe la fase de grupos, el administrador distribuirá una **segunda plantilla** con los emparejamientos reales.

**1. Descarga la plantilla de eliminatorias**
Copia `plantillas/plantilla_eliminatorias.xlsx` y renómbrala `porra_TUNOMBRE_elim.xlsx`.
⚠️ Fíjate en el sufijo `_elim` — es obligatorio para que el sistema lo reconozca.

**2. Escribe tu alias** — el **mismo** que usaste en la fase de grupos.

**3. Rellena tus pronósticos**

- **Columnas D y E (amarillo):** goles al término de los **90 minutos + descuento** (sin prórroga ni penaltis).
- **Columna G (verde — Quién Pasa):** escribe el equipo que crees que avanza de ronda. Especialmente importante si predices un empate a 90' (+1 pt extra).

**4. Entrega el archivo**
Guarda `porra_TUNOMBRE_elim.xlsx` en la carpeta `pronosticos/` antes de cada fecha límite:

| Fase            | Fecha límite              |
| --------------- | ------------------------- |
| Ronda de 32     | 28 jun 2026 · 17:00 h     |
| Octavos         | 4 jul 2026 · 17:00 h      |
| Cuartos         | 9 jul 2026 · 17:00 h      |
| Semifinales     | 14 jul 2026 · 17:00 h     |
| Final           | 17 jul 2026 · 17:00 h     |

> Si subes el archivo varias veces, vale siempre el más reciente antes del límite.

---

## Para el administrador

### Antes del torneo

1. Doble click en **`Regenerar_Plantillas.command`** (Mac) o **`Regenerar_Plantillas.bat`** (Windows).
   Genera `plantillas/plantilla_porra.xlsx` (grupos + premios). Distribúyela a los participantes.

### Durante la fase de grupos

1. Tras cada partido, abre `data/resultados_reales.xlsx`, busca el ID del partido (columna C) y escribe los goles reales (columnas D y E).
2. Doble click en **`Actualizar_Ranking`** para recalcular y abrir el dashboard.

### Cuando acaba la fase de grupos

1. Actualiza `data/matches.json` con los emparejamientos reales de la Ronda de 32 (equipos locales y visitantes de cada R32-XX).
2. Doble click en **`Regenerar_Plantilla_Elim.command`** (Mac) o **`Regenerar_Plantilla_Elim.bat`** (Windows).
   Genera `plantillas/plantilla_eliminatorias.xlsx`. Distribúyela a los participantes.
3. Los participantes la rellenan y la entregan como `porra_TUNOMBRE_elim.xlsx`.

### Durante las eliminatorias

1. Introduce resultados reales en `data/resultados_reales_elim.xlsx` (mismo formato: ID en C, goles en D y E).
2. Si un partido acaba en empate a 90' (va a prórroga/penaltis), abre `data/quien_paso.json` y añade:
   ```json
   { "R32-01": "España", "QF-02": "Brasil" }
   ```
3. Doble click en **`Actualizar_Ranking`** para actualizar.

### Al finalizar el torneo

1. Abre `data/premios_reales.json` y escribe el ganador exacto de cada premio:
   ```json
   { "mvp": "Kylian Mbappe", "goleador": "Erling Haaland", "portero": "Emiliano Martinez" }
   ```
2. Doble click en **`Actualizar_Ranking`** una última vez.

---

## Resumen de archivos por participante

| Archivo                       | Cuándo se entrega               | Contenido                        |
| ----------------------------- | ------------------------------- | -------------------------------- |
| `porra_TUNOMBRE.xlsx`         | Antes del 11 jun · 17:00 h      | Grupos (72 partidos) + Premios   |
| `porra_TUNOMBRE_elim.xlsx`    | Antes de cada fase eliminatoria | Eliminatorias (32 partidos)      |

**El alias en los dos archivos debe ser idéntico** — así el sistema los asocia automáticamente.

---

## Puntuación rápida

| Concepto                                    | Puntos       |
| ------------------------------------------- | ------------ |
| Resultado exacto (goles)                    | 3            |
| Resultado correcto (ganador o empate a 90') | 2            |
| Quién pasa de ronda (eliminatorias)         | +1           |
| MVP / Goleador / Portero del torneo         | +15 cada uno |
| Fallo                                       | 0            |

**Desempate:** gana quien tenga más resultados exactos (3 pts).

**Premios económicos** (aportación 10 € por persona):
- 30 % del bote → ganador de la fase de grupos
- 70 % del bote → ganador de la clasificación final
