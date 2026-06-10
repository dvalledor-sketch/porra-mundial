"""
Genera el archivo maestro data/matches.json con los 104 partidos del Mundial 2026.

Fase de grupos (72 partidos):  IDs G-{Letra}{n}-{Letra}{m}  ej. G-A1-A2
Eliminatorias (32 partidos):
    Ronda de 32 (1/16): R32-01 ... R32-16
    Octavos    (1/8) : R16-01 ... R16-08
    Cuartos    (1/4) : QF-01  ... QF-04
    Semifinales      : SF-01, SF-02
    Tercer puesto    : TER
    Final            : FIN
Total: 72 + 16 + 8 + 4 + 2 + 1 + 1 = 104 ✓

Los datos de equipos provienen del sorteo final del 5 de diciembre de 2025
en el Kennedy Center de Washington D.C. (FIFA).
"""

import json
from pathlib import Path

# 12 grupos × 4 selecciones, en posiciones 1..4
# (Los datos pueden ajustarse si la FIFA actualiza alguna plaza de repesca)
GRUPOS = {
    "A": ["México",        "Sudáfrica",         "Corea del Sur",     "Chequia"],
    "B": ["Canadá",        "Suiza",             "Catar",             "Bosnia y Herzegovina"],
    "C": ["Brasil",        "Marruecos",         "Escocia",           "Haití"],
    "D": ["Estados Unidos","Paraguay",          "Australia",         "Turquía"],
    "E": ["Alemania",      "Curazao",           "Costa de Marfil",        "Ecuador"],
    "F": ["Países Bajos",  "Japón",             "Túnez",             "Suecia"],
    "G": ["Bélgica",       "Egipto",            "Irán",              "Nueva Zelanda"],
    "H": ["España",        "Cabo Verde",        "Arabia Saudí",      "Uruguay"],
    "I": ["Francia",       "Senegal",           "Noruega",           "Irak"],
    "J": ["Argentina",     "Argelia",           "Austria",           "Jordania"],
    "K": ["Portugal",      "Colombia",          "Uzbekistán",        "RD Congo"],
    "L": ["Inglaterra",    "Croacia",           "Ghana",             "Panamá"],
}

# Pares de jornadas de la fase de grupos (orden tradicional FIFA: J1, J2, J3)
ENFRENTAMIENTOS_GRUPO = [
    (1, 2),  # Jornada 1
    (3, 4),  # Jornada 1
    (1, 3),  # Jornada 2
    (2, 4),  # Jornada 2
    (1, 4),  # Jornada 3
    (2, 3),  # Jornada 3
]

# Fechas estimadas por jornada (rango). El Mundial empieza el 11 de junio y la fase
# de grupos termina aprox. el 27 de junio.  El usuario puede afinar al partido.
JORNADAS_FECHAS = {
    1: "2026-06-11 / 2026-06-17",
    2: "2026-06-15 / 2026-06-22",
    3: "2026-06-21 / 2026-06-27",
}


def construir_grupos():
    partidos = []
    for letra, equipos in GRUPOS.items():
        for idx, (i, j) in enumerate(ENFRENTAMIENTOS_GRUPO):
            jornada = idx // 2 + 1
            pid = f"G-{letra}{i}-{letra}{j}"
            partidos.append({
                "id": pid,
                "fase": "grupos",
                "grupo": letra,
                "jornada": jornada,
                "fecha_aprox": JORNADAS_FECHAS[jornada],
                "local": equipos[i - 1],
                "visitante": equipos[j - 1],
            })
    return partidos


def construir_eliminatorias():
    partidos = []
    # Ronda de 32 (1/16):  los emparejamientos exactos dependen del 3.º mejor
    # — el reglamento permite editarlos a posteriori. Dejamos slots con etiqueta
    # genérica que el administrador rellenará tras los grupos.
    for n in range(1, 17):
        partidos.append({
            "id": f"R32-{n:02d}",
            "fase": "ronda_32",
            "fecha_aprox": "2026-06-28 / 2026-07-03",
            "local":     f"Por definir (R32-{n:02d} L)",
            "visitante": f"Por definir (R32-{n:02d} V)",
        })
    for n in range(1, 9):
        partidos.append({
            "id": f"R16-{n:02d}",
            "fase": "octavos",
            "fecha_aprox": "2026-07-04 / 2026-07-07",
            "local":     f"Por definir (R16-{n:02d} L)",
            "visitante": f"Por definir (R16-{n:02d} V)",
        })
    for n in range(1, 5):
        partidos.append({
            "id": f"QF-{n:02d}",
            "fase": "cuartos",
            "fecha_aprox": "2026-07-09 / 2026-07-11",
            "local":     f"Por definir (QF-{n:02d} L)",
            "visitante": f"Por definir (QF-{n:02d} V)",
        })
    for n in range(1, 3):
        partidos.append({
            "id": f"SF-{n:02d}",
            "fase": "semifinales",
            "fecha_aprox": "2026-07-14 / 2026-07-15",
            "local":     f"Por definir (SF-{n:02d} L)",
            "visitante": f"Por definir (SF-{n:02d} V)",
        })
    partidos.append({
        "id": "TER",
        "fase": "tercer_puesto",
        "fecha_aprox": "2026-07-18",
        "local":     "Por definir (TER L)",
        "visitante": "Por definir (TER V)",
    })
    partidos.append({
        "id": "FIN",
        "fase": "final",
        "fecha_aprox": "2026-07-19",
        "local":     "Por definir (FIN L)",
        "visitante": "Por definir (FIN V)",
    })
    return partidos


def main():
    aquí = Path(__file__).resolve().parent.parent
    salida = aquí / "data" / "matches.json"
    salida.parent.mkdir(exist_ok=True, parents=True)

    grupos = construir_grupos()
    elim   = construir_eliminatorias()
    todos  = grupos + elim

    payload = {
        "torneo": "Copa Mundial de la FIFA 2026",
        "sede":   "Estados Unidos, México y Canadá",
        "total_partidos": len(todos),
        "grupos": GRUPOS,
        "partidos": todos,
    }
    salida.write_text(json.dumps(payload, ensure_ascii=False, indent=2), encoding="utf-8")
    print(f"OK · {len(todos)} partidos escritos en {salida}")
    assert len(todos) == 104, f"esperado 104, obtenido {len(todos)}"


if __name__ == "__main__":
    main()
