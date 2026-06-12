"""
fetch_results.py — descarga resultados del Mundial 2026 desde API-Football
y actualiza data/resultados_reales.xlsx automáticamente.

Requiere la variable de entorno FOOTBALL_API_KEY con tu clave de api-sports.io
(cuenta gratuita: https://dashboard.api-football.com/register).

Uso manual:
    FOOTBALL_API_KEY=xxx python scripts/fetch_results.py
"""

from __future__ import annotations

import json
import os
import sys
from pathlib import Path

import requests
from openpyxl import load_workbook

# ── rutas ─────────────────────────────────────────────────────────────────────
RAIZ        = Path(__file__).resolve().parent.parent
MATCHES_F   = RAIZ / "data" / "matches.json"
REALES_F    = RAIZ / "data" / "resultados_reales.xlsx"
REALES_ELIM = RAIZ / "data" / "resultados_reales_elim.xlsx"

API_BASE    = "https://v3.football.api-sports.io"
API_KEY     = os.environ.get("FOOTBALL_API_KEY", "")
WC_LEAGUE   = 1      # FIFA World Cup — id fijo según docs oficiales
WC_SEASON   = 2026

# ── mapeo nombre API (inglés) → nombre interno (español) ──────────────────────
NOMBRES: dict[str, str] = {
    "Mexico": "México", "South Africa": "Sudáfrica",
    "South Korea": "Corea del Sur", "Korea Republic": "Corea del Sur",
    "Czechia": "Chequia", "Czech Republic": "Chequia",
    "Canada": "Canadá", "Switzerland": "Suiza",
    "Qatar": "Catar", "Bosnia": "Bosnia y Herzegovina",
    "Bosnia and Herzegovina": "Bosnia y Herzegovina",
    "Brazil": "Brasil", "Morocco": "Marruecos",
    "Scotland": "Escocia", "Haiti": "Haití",
    "USA": "Estados Unidos", "United States": "Estados Unidos",
    "Paraguay": "Paraguay", "Australia": "Australia",
    "Turkey": "Turquía", "Türkiye": "Turquía",
    "Germany": "Alemania", "Curacao": "Curazao",
    "Ivory Coast": "Costa de Marfil", "Côte d'Ivoire": "Costa de Marfil", "Cote d'Ivoire": "Costa de Marfil", "Ecuador": "Ecuador",
    "Netherlands": "Países Bajos", "Japan": "Japón",
    "Tunisia": "Túnez", "Sweden": "Suecia",
    "Belgium": "Bélgica", "Egypt": "Egipto",
    "Iran": "Irán", "New Zealand": "Nueva Zelanda",
    "Spain": "España", "Cape Verde": "Cabo Verde",
    "Saudi Arabia": "Arabia Saudí", "Uruguay": "Uruguay",
    "France": "Francia", "Senegal": "Senegal",
    "Norway": "Noruega", "Iraq": "Irak",
    "Argentina": "Argentina", "Algeria": "Argelia",
    "Austria": "Austria", "Jordan": "Jordania",
    "Portugal": "Portugal", "Colombia": "Colombia",
    "Uzbekistan": "Uzbekistán", "DR Congo": "RD Congo",
    "England": "Inglaterra", "Croatia": "Croacia",
    "Ghana": "Ghana", "Panama": "Panamá", "Panamá": "Panamá",
}


def _es(api_name: str) -> str:
    """Convierte nombre API a nombre español."""
    return NOMBRES.get(api_name, api_name)


def _api_get(endpoint: str, params: dict) -> dict:
    r = requests.get(
        f"{API_BASE}/{endpoint}",
        headers={"x-apisports-key": API_KEY},
        params=params,
        timeout=15,
    )
    r.raise_for_status()
    data = r.json()
    # API-Football devuelve HTTP 200 con un campo "errors" cuando algo falla
    # (clave inválida, plan sin acceso a la temporada, cuota diaria agotada…).
    errors = data.get("errors")
    if errors and (not isinstance(errors, (list, dict)) or len(errors) > 0):
        raise RuntimeError(f"API-Football devolvió errores: {errors}")
    return data


def _find_wc_league() -> int:
    """Devuelve el league_id del Mundial 2026 (siempre 1 según docs oficiales)."""
    print(f"✓ Mundial 2026 — league_id={WC_LEAGUE}, season={WC_SEASON}")
    return WC_LEAGUE


def _build_match_index() -> dict[tuple[str, str], str]:
    """Construye dict (local_es, visitante_es) → match_id a partir de matches.json."""
    with open(MATCHES_F, encoding="utf-8") as f:
        data = json.load(f)
    return {(m["local"], m["visitante"]): m["id"] for m in data["partidos"]}


def _get_all_fixtures(league_id: int) -> list[dict]:
    """Devuelve TODOS los fixtures del Mundial 2026 en una sola llamada.

    Una única petición (en vez de 1 por ronda) para no agotar la cuota
    diaria del plan gratuito de API-Football (100 peticiones/día).
    """
    data = _api_get("fixtures", {"league": league_id, "season": WC_SEASON})
    return data.get("response", [])


FINISHED = {"FT", "AET", "PEN"}


def _get_fixtures(all_fixtures: list[dict]) -> list[dict]:
    """Filtra los partidos ya terminados."""
    return [f for f in all_fixtures
            if f["fixture"]["status"]["short"] in FINISHED]


def _update_excel(archivo: Path, resultados: dict[str, tuple[int, int]]) -> int:
    """
    Actualiza el Excel de resultados reales con los marcadores obtenidos.
    Solo rellena celdas que estén vacías (no sobreescribe correcciones manuales).
    Devuelve el número de celdas actualizadas.
    """
    if not archivo.exists():
        print(f"⚠️  No encontrado: {archivo}")
        return 0

    wb = load_workbook(archivo)
    actualizados = 0

    for sheet_name in wb.sheetnames:
        ws = wb[sheet_name]
        for row in ws.iter_rows(min_row=1):
            cells = [c.value for c in row]
            padded = cells + [None] * (10 - len(cells))
            match_id = padded[2]
            if not isinstance(match_id, str):
                continue
            match_id = match_id.strip()
            if match_id not in resultados:
                continue
            gl, gv = resultados[match_id]
            col_d = row[3] if len(row) > 3 else None
            col_e = row[4] if len(row) > 4 else None
            if col_d is not None and col_d.value is None:
                col_d.value = gl
                actualizados += 1
            if col_e is not None and col_e.value is None:
                col_e.value = gv
                actualizados += 1

    if actualizados:
        wb.save(archivo)
    return actualizados


def main() -> None:
    if not API_KEY:
        print("✘ Variable FOOTBALL_API_KEY no definida. Saliendo.")
        sys.exit(1)

    print("▸ Buscando Mundial 2026 en API-Football…")
    league_id = _find_wc_league()
    print("▸ Descargando fixtures (1 sola llamada)…")
    all_fixtures = _get_all_fixtures(league_id)
    print(f"  {len(all_fixtures)} fixtures totales")
    if not all_fixtures:
        print("✘ La API devolvió 0 fixtures sin reportar error explícito.")
        print("  Revisa en https://dashboard.api-football.com que tu plan")
        print("  cubre la temporada 2026 y que queda cuota diaria.")
        sys.exit(1)
    fixtures = _get_fixtures(all_fixtures)
    print(f"  {len(fixtures)} partidos terminados")

    idx = _build_match_index()
    resultados: dict[str, tuple[int, int]] = {}

    for fix in fixtures:
        home = _es(fix["teams"]["home"]["name"])
        away = _es(fix["teams"]["away"]["name"])
        # goals.home/away es el campo estándar; score.fulltime como fallback
        gl   = fix.get("goals", {}).get("home") or fix["score"]["fulltime"]["home"]
        gv   = fix.get("goals", {}).get("away") or fix["score"]["fulltime"]["away"]
        if gl is None or gv is None:
            continue
        match_id = idx.get((home, away))
        if match_id:
            resultados[match_id] = (int(gl), int(gv))
        else:
            print(f"  ⚠️  Sin mapeo para: {home} vs {away}")

    print(f"  {len(resultados)} partidos mapeados a IDs internos")

    n_grupos = _update_excel(REALES_F, resultados)
    n_elim   = _update_excel(REALES_ELIM, resultados)
    total    = n_grupos + n_elim
    print(f"✓ Excel actualizado — {total} celdas nuevas")

    print("▸ Construyendo cuadro eliminatorio…")
    bracket = _build_bracket(all_fixtures)
    import datetime as _dt
    BRACKET_F.write_text(
        json.dumps({"updated": _dt.datetime.utcnow().isoformat(), "rounds": bracket},
                   ensure_ascii=False, indent=2),
        encoding="utf-8"
    )
    print(f"✓ bracket.json actualizado")


# ── Bracket (eliminatorias) ───────────────────────────────────────────────────
BRACKET_F = RAIZ / "data" / "bracket.json"

ROUND_MAP: dict[str, str] = {
    "ronda_32":      "Round of 32",
    "octavos":       "Round of 16",
    "cuartos":       "Quarter-finals",
    "semifinales":   "Semi-finals",
    "tercer_puesto": "3rd Place Final",
    "final":         "Final",
}

ROUND_ORDER = ["ronda_32", "octavos", "cuartos", "semifinales", "final", "tercer_puesto"]


def _build_bracket(all_fixtures: list[dict]) -> dict[str, list[dict]]:
    """Construye el cuadro eliminatorio a partir de los fixtures ya descargados
    (sin llamadas extra a la API)."""
    bracket: dict[str, list[dict]] = {fase: [] for fase in ROUND_MAP}
    round_to_fase = {api_round: fase for fase, api_round in ROUND_MAP.items()}
    for f in all_fixtures:
        fase = round_to_fase.get(f.get("league", {}).get("round", ""))
        if not fase:
            continue
        bracket[fase].append({
            "fixture_id": f["fixture"]["id"],
            "date":       f["fixture"]["date"],
            "status":     f["fixture"]["status"]["short"],
            "team1":      _es(f["teams"]["home"]["name"]),
            "team2":      _es(f["teams"]["away"]["name"]),
            "score1":     (f.get("goals") or {}).get("home") or f["score"]["fulltime"]["home"],
            "score2":     (f.get("goals") or {}).get("away") or f["score"]["fulltime"]["away"],
        })
    for fase, fixes in bracket.items():
        fixes.sort(key=lambda x: x["date"] or "")
        print(f"  {fase}: {len(fixes)} partidos")
    return bracket


if __name__ == "__main__":
    main()
