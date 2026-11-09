"""Genera index.html a partir del Excel de juegos de mesa.

Uso:
    python build.py "ruta/al/Juegos de mesa.xlsx"

Lee la primera hoja, normaliza categorías (género, modalidad, curva), corrige
nombres mal escritos, saca el ID de YouTube de cada reseña (también de las celdas
que son hipervínculo con otro texto visible) e inyecta todo como JSON en
template.html. El Excel original no se modifica.
"""
import json
import re
import sys
from pathlib import Path

import openpyxl

HERE = Path(__file__).parent

# Nombre en el Excel (sin espacios sobrantes) -> (nombre a mostrar, nota opcional)
NOMBRES = {
    "Cycaldes (edicion legendaria) la normal no está": ("Cyclades", "Edición legendaria (la normal no está)"),
    "El Grande (el que le gusta al mama)": ("El Grande", "El que le gusta al Mama"),
    "Carcasone": ("Carcassonne", None),
    "Jauipur": ("Jaipur", None),
    "Risin sun": ("Rising Sun", None),
    "Wavelenght": ("Wavelength", None),
    "Explodingkittens": ("Exploding Kittens", None),
    "ArkNova": ("Ark Nova", None),
    "Codigo secreto": ("Código Secreto", None),
    "Lord of waterdeep": ("Lords of Waterdeep", None),
    "King of tokyo": ("King of Tokyo", None),
    "Gunsen: The batlle for Toshi Ranbo": ("Gunsen: The Battle for Toshi Ranbo", None),
    "HEAT": ("Heat", None),
    "MAGIC MAZE": ("Magic Maze", None),
    "Stone age": ("Stone Age", None),
    "Terraforming mars": ("Terraforming Mars", None),
    "Dead of winter": ("Dead of Winter", None),
    "Feed the kraken": ("Feed the Kraken", None),
    "it's a wonderful world": ("It's a Wonderful World", None),
    "libertalia": ("Libertalia", None),
    "Kingdom builder": ("Kingdom Builder", None),
    "So clover": ("So Clover!", None),
    "Twisted cryptids": ("Twisted Cryptids", None),
    "Leviathan wilds": ("Leviathan Wilds", None),
    "Survive the island": ("Survive the Island", None),
    "Through the desert": ("Through the Desert", None),
    "Say my name": ("Say My Name", None),
    "One round": ("One Round", None),
    "Garabatos fatidicos": ("Garabatos fatídicos", None),
    "Gloomhaven: fauces del leon": ("Gloomhaven: Fauces del León", None),
    "La tripulacion": ("La Tripulación", None),
    "La tripulacion: mar profundo": ("La Tripulación: Mar Profundo", None),
    "Las ruinas perdidas de Arnak": ("Las Ruinas Perdidas de Arnak", None),
    "The red cathedral": ("The Red Cathedral", None),
    "The Manhattan Project: Energy empire": ("The Manhattan Project: Energy Empire", None),
    "Eclipse: el segundo amanecer de la galaxia": ("Eclipse: El Segundo Amanecer de la Galaxia", None),
    "Sushi Go Party": ("Sushi Go Party!", None),
    "Aventureros al tren": ("¡Aventureros al Tren!", None),
}

# Género original -> familia para filtrar (el original se sigue mostrando en la ficha)
FAMILIAS = {
    "Eurogame": ["eurogame", "eurogame pesado", "eurogame económico", "eurogame familiar",
                 "eurogame cooperativo", "estrategia / tableau building"],
    "Control territorial": ["control territorial", "wargame ligero"],
    "Party": ["party game", "party game / deducción", "party game / roles ocultos",
              "destreza", "tiempo real"],
    "Deducción social": ["deducción social", "deducción"],
    "Temático / Supervivencia": ["temático / supervivencia", "temático"],
    "Rol / Combate": ["rol / dungeon crawler", "táctico / combate"],
    "Losetas": ["colocación de losetas"],
    "Cartas": ["cartas / familiar", "bazas (cartas)", "juego de bazas", "familiar / draft"],
    "Construcción de mazos": ["construcción de mazos"],
    "Riesgo / Carrera": ["carrera / forzar la suerte", "push your luck"],
    "Familiar": ["familiar", "familiar / gestión de mano", "familiar / gestión",
                 "familiar / colocación"],
}
GENERO_A_FAMILIA = {g: fam for fam, gs in FAMILIAS.items() for g in gs}

MODALIDAD = {
    "competitivo": "Competitivo",
    "cooperativo": "Cooperativo",
    "equipos": "Equipos",
    "competitivo por equipos": "Equipos",
    "semicooperativo": "Semicooperativo",
}

CURVA = {
    "muy baja": "Muy baja",
    "baja": "Baja",
    "media": "Media",
    "media - alta": "Media-alta",
    "alta": "Alta",
    "muy alta": "Muy alta",
}

YT_RE = re.compile(r"(?:youtube\.com/(?:watch\?(?:.*&)?v=|shorts/|embed/)|youtu\.be/)([A-Za-z0-9_-]{11})")


def limpio(v):
    if v is None:
        return None
    s = re.sub(r"[ \t]+", " ", str(v)).strip()
    return s or None


def numero(v):
    if v in (None, ""):
        return None
    try:
        return float(v)
    except (TypeError, ValueError):
        return None


def main(xlsx_path):
    ws = openpyxl.load_workbook(xlsx_path).worksheets[0]
    headers = [limpio(c.value) for c in ws[1]]
    col = {h: i for i, h in enumerate(headers) if h}

    def idx(nombre):
        if nombre not in col:
            raise SystemExit(f"Falta la columna '{nombre}' en el Excel. Columnas: {list(col)}")
        return col[nombre]

    c_link = idx("Link reseña")
    juegos, avisos = [], []

    for r in range(2, ws.max_row + 1):
        fila = [c.value for c in ws[r]]
        original = limpio(fila[idx("Juego")])
        if not original:
            continue
        nombre, nota = NOMBRES.get(original, (original, None))

        genero = limpio(fila[idx("Género")]) or ""
        familia = GENERO_A_FAMILIA.get(genero.lower())
        if not familia:
            familia = "Otros"
            avisos.append(f"Género sin familia: '{genero}' ({nombre})")

        modalidad_raw = limpio(fila[idx("Modalidad")]) or ""
        modalidad = MODALIDAD.get(modalidad_raw.lower(), modalidad_raw or None)

        curva_raw = (limpio(fila[idx("Curva de aprendizaje")]) or "").lower()
        curva = CURVA.get(curva_raw, curva_raw.capitalize() or None)

        celda_link = ws.cell(r, c_link + 1)
        url = (celda_link.hyperlink.target if celda_link.hyperlink else None) or limpio(celda_link.value)
        m = YT_RE.search(url or "")
        video = m.group(1) if m else None
        if url and not video:
            avisos.append(f"Link que no es de YouTube: {nombre} -> {url}")

        precio = numero(fila[idx("Precio")])
        desc = str(fila[idx("Descripción breve")] or "").strip()
        parrafos = [p.strip() for p in desc.split("\n") if p.strip()]

        juegos.append({
            "nombre": nombre,
            "nota": nota,
            "precio": precio if precio else None,
            "genero": genero,
            "familia": familia,
            "tematica": limpio(fila[idx("Temática")]),
            "modalidad": modalidad,
            "desc": parrafos,
            "jmin": numero(fila[idx("Jugadores mínimos")]),
            "jmax": numero(fila[idx("Jugadores máximos")]),
            "dur": numero(fila[idx("Duración máxima (min)")]),
            "comp": numero(fila[idx("Complejidad (1-5)")]),
            "azar": numero(fila[idx("Azar (1-5)")]),
            "rejug": numero(fila[idx("Rejugabilidad (1-5)")]),
            "curva": curva,
            "video": video,
            "url": url if url and url.startswith("http") else None,
        })

    juegos.sort(key=lambda j: j["nombre"].lower().lstrip("¡"))
    datos = json.dumps(juegos, ensure_ascii=False, separators=(",", ":")).replace("</", "<\\/")
    plantilla = (HERE / "template.html").read_text(encoding="utf-8")
    if "/*__DATA__*/" not in plantilla:
        raise SystemExit("template.html no tiene el marcador /*__DATA__*/")
    (HERE / "index.html").write_text(plantilla.replace("/*__DATA__*/[]", datos), encoding="utf-8")

    print(f"index.html generado con {len(juegos)} juegos")
    print(f"  sin video: {[j['nombre'] for j in juegos if not j['video']]}")
    print(f"  sin precio: {[j['nombre'] for j in juegos if not j['precio']]}")
    for a in avisos:
        print("  AVISO:", a)


if __name__ == "__main__":
    if len(sys.argv) != 2:
        raise SystemExit(__doc__)
    main(sys.argv[1])
