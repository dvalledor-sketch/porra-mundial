"""
download_avatars.py — Descarga retratos de futbolistas históricos de Wikipedia
y los guarda como avatares 512×512 en la carpeta Avatar/.

Uso:
    python3 scripts/download_avatars.py
"""

import urllib.request, json, io
from pathlib import Path
from PIL import Image

AVATAR_DIR = Path(__file__).resolve().parent.parent / "Avatar"
SIZE = 512

PLAYERS = [
    ("01_rene_higuita",      "René_Higuita"),
    ("02_carlos_valderrama", "Carlos_Valderrama"),
    ("03_roberto_baggio",    "Roberto_Baggio"),
    ("04_ronaldinho",        "Ronaldinho"),
    ("05_romario",           "Romário"),
    ("06_hristo_stoichkov",  "Hristo_Stoichkov"),
    ("07_oliver_kahn",        "Oliver_Kahn"),
    ("08_diego_maradona",    "Diego_Maradona"),
    ("09_ruud_gullit",       "Ruud_Gullit"),
    ("10_claudio_caniggia",  "Claudio_Paul_Caniggia"),
    ("11_eric_cantona",      "Eric_Cantona"),
    ("12_zinedine_zidane",   "Zinedine_Zidane"),
    ("13_david_beckham",     "David_Beckham"),
    ("14_peter_schmeichel",  "Peter_Schmeichel"),
    ("15_marco_van_basten",  "Marco_van_Basten"),
    ("16_paolo_maldini",     "Paolo_Maldini"),
    ("17_guti",              "Guti_(footballer)"),
    ("18_raul_gonzalez",     "Raúl_(footballer)"),
    ("19_roger_milla",       "Roger_Milla"),
    ("20_jurgen_klinsmann",  "Jürgen_Klinsmann"),
]

HEADERS = {"User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36"}


def get_image_url(wiki_title: str):
    api = f"https://en.wikipedia.org/api/rest_v1/page/summary/{urllib.request.quote(wiki_title)}"
    req = urllib.request.Request(api, headers=HEADERS)
    with urllib.request.urlopen(req, timeout=12) as r:
        data = json.loads(r.read())
    return data.get("originalimage", {}).get("source") or data.get("thumbnail", {}).get("source")


def process_and_save(img_bytes: bytes, out_path: Path):
    img = Image.open(io.BytesIO(img_bytes)).convert("RGB")
    w, h = img.size
    # Recortar cuadrado desde arriba (foco en la cara)
    side = min(w, h)
    left = (w - side) // 2
    img = img.crop((left, 0, left + side, side))
    img = img.resize((SIZE, SIZE), Image.LANCZOS)
    img.save(out_path, "PNG")


import time, random

def download_one(slug, wiki):
    out = AVATAR_DIR / f"{slug}_512.png"
    if out.exists():
        print(f"▸ {slug}... ya existe, skip")
        return True
    for attempt in range(4):
        if attempt:
            wait = 3 + attempt * 2 + random.random() * 2
            print(f"  reintento {attempt} en {wait:.1f}s...", end=" ", flush=True)
            time.sleep(wait)
        try:
            url = get_image_url(wiki)
            if not url:
                raise ValueError("sin URL")
            req = urllib.request.Request(url, headers=HEADERS)
            with urllib.request.urlopen(req, timeout=20) as r:
                data = r.read()
            process_and_save(data, out)
            print(f"✓  ({len(data)//1024} KB)")
            return True
        except Exception as e:
            print(f"✗ {e}", end=" ", flush=True)
    print()
    return False

ok, fail = [], []
for slug, wiki in PLAYERS:
    print(f"▸ {slug}...", end=" ", flush=True)
    if download_one(slug, wiki):
        ok.append(slug)
    else:
        fail.append(slug)
    time.sleep(1.5 + random.random())

print(f"\n{'='*50}")
print(f"✓ {len(ok)} guardados en Avatar/")
if fail:
    print(f"✗ Fallaron: {fail}")
print("\nEjecuta scorer.py para regenerar el dashboard con los nuevos avatares.")
