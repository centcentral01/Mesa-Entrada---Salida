# scripts/importar_salidas.py — Importa las notas de salida (npNNNN - <asunto>.doc) a mesa_entradas.salidas.
#
# Lee un zip o una carpeta de OneDrive con subcarpetas "Notas AAAA". De cada nota toma:
#   - año: de la carpeta "Notas AAAA" donde está
#   - número y asunto: del nombre del archivo (np0137 - Juzgado Civil 42 CABA responde oficio ...)
#   - fecha, "Nota N° 0137/25" y expediente: del texto del documento (Word por COM; PDF con pdfplumber)
#   - destinatario: lo que va antes del verbo en el nombre ("SSTA Sr Subsecretario" eleva ...)
#
# Números repetidos: si los archivos dicen lo mismo (el PDF de un Word, una copia) queda un solo
# registro; si son notas distintas con el mismo número, la segunda va con sufijo "b", la tercera "c".
#
# Sin --grabar solo arma la planilla de revisión (revision/salidas_<año>.csv). Con --grabar hace
# upsert por (anio, numero, sufijo): si se corre de nuevo, actualiza en lugar de duplicar.
#
# Uso:
#   python scripts/importar_salidas.py "C:/Users/Usuario/Downloads/OneDrive_1_16-10-2026.zip" --anio 2024
#   python scripts/importar_salidas.py "<carpeta con Notas 2026>" --grabar

import argparse
import csv
import re
import sys
import tempfile
import zipfile
from collections import defaultdict
from datetime import date
from pathlib import Path, PurePosixPath

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))

MESES = {"enero": 1, "febrero": 2, "marzo": 3, "abril": 4, "mayo": 5, "junio": 6, "julio": 7,
         "agosto": 8, "septiembre": 9, "setiembre": 9, "octubre": 10, "noviembre": 11, "diciembre": 12}

RE_ARCHIVO = re.compile(r"^np\s?(\d{1,4})\s*-?\s*(.*?)\.(docx?|pdf)$", re.I)
RE_CARPETA = re.compile(r"Notas\s+(\d{4})", re.I)
RE_FECHA = re.compile(r"(\d{1,2})\s+de\s+([a-záéíóú]+)\s+(?:de|del)\s+(\d{4})", re.I)
RE_NOTA = re.compile(r"Nota\s*N\s*[°º.o]*\s*:?\s*(\d{1,5})\s*/\s*(\d{2,4})", re.I)
RE_EX = re.compile(r"EX-\s?\d{4}-\s?\d+", re.I)
RE_EXPTE = re.compile(r"Expte\.?\s*(?:N(?:ro)?\.?\s*[°º]?\s*)?(\d[\d./-]*\d)", re.I)
RE_NO = re.compile(r"NO-\d{4}-\d+-APN-[A-Z#-]+", re.I)
RE_SSTA = re.compile(r"^SSTAU?\s+S[rR]\.?\s+S\w*ubsecretario$", re.I)
# En 2023 los nombres arrancan con "2023 - A <destinatario>".
RE_PREFIJO = re.compile(r"^(?:\d{4}\s*-\s*)?(?:A\s+(?=[A-ZÁÉÍÓÚ]))?")
# El destinatario es lo que va antes del primer verbo del nombre del archivo.
RE_VERBO = re.compile(r"\s+-?\s*\b(eleva|elevaci[oó]n|responde|respuesta|resp\.|contesta|informa|solicitud|solicita|"
                      r"consulta|remite|notificaci[oó]n|cierre|en relaci[oó]n|pedir|sobre|reclamo|alta|causa)\b", re.I)


def leer_entrada(origen: Path, anios):
    """[(anio, ruta relativa, nombre, bytes)] de las notas npNNNN dentro de carpetas "Notas AAAA"."""
    def anio_de(ruta):
        m = RE_CARPETA.search(ruta)
        return int(m.group(1)) if m else None

    salida = []
    if origen.suffix.lower() == ".zip":
        with zipfile.ZipFile(origen) as z:
            for i in z.infolist():
                nombre = PurePosixPath(i.filename).name
                anio = anio_de(i.filename)
                if not i.is_dir() and anio and RE_ARCHIVO.match(nombre) and (not anios or anio in anios):
                    salida.append((anio, i.filename, nombre, z.read(i)))
    else:
        for p in sorted(origen.rglob("*")):
            rel = p.relative_to(origen.parent).as_posix()
            anio = anio_de(rel)
            if p.is_file() and anio and RE_ARCHIVO.match(p.name) and (not anios or anio in anios):
                salida.append((anio, rel, p.name, p.read_bytes()))
    return salida


class LectorTexto:
    """Saca el texto de .doc/.docx con Word (una sola instancia) y de .pdf con pdfplumber."""

    def __init__(self):
        self.word = None
        self.tmp = tempfile.TemporaryDirectory()

    def texto(self, nombre, datos):
        ext = nombre.rsplit(".", 1)[-1].lower()
        # Nombre corto: los originales pasan los 260 caracteres de ruta de Windows.
        ruta = Path(self.tmp.name) / f"nota.{ext}"
        ruta.write_bytes(datos)
        if ext == "pdf":
            import pdfplumber
            with pdfplumber.open(ruta) as pdf:
                return "\n".join(p.extract_text() or "" for p in pdf.pages[:2])
        if self.word is None:
            import win32com.client
            self.word = win32com.client.DispatchEx("Word.Application")
            self.word.Visible = False
            self.word.DisplayAlerts = 0
        doc = self.word.Documents.Open(str(ruta), ReadOnly=True, AddToRecentFiles=False, ConfirmConversions=False)
        try:
            return doc.Content.Text.replace("\r", "\n").replace("\x07", "")
        finally:
            doc.Close(False)

    def cerrar(self):
        if self.word is not None:
            self.word.Quit()
        self.tmp.cleanup()


def limpiar(s):
    return re.sub(r"\s+", " ", s or "").strip(" -.")


def analizar(anio, ruta, nombre, texto):
    m = RE_ARCHIVO.match(nombre)
    numero, asunto = int(m.group(1)), limpiar(m.group(2))
    obs = []

    base = RE_PREFIJO.sub("", asunto)
    partes = RE_VERBO.split(base, maxsplit=1)
    destinatario = limpiar(partes[0]) if len(partes) > 1 else None
    if not destinatario and " - " in base:
        destinatario = limpiar(base.split(" - ", 1)[0])
    if destinatario and RE_SSTA.match(destinatario):
        destinatario = "SSTA Sr Subsecretario"

    cabeza = texto[:600]
    fecha = None
    if f := RE_FECHA.search(cabeza):
        mes = MESES.get(f.group(2).lower())
        try:
            fecha = date(int(f.group(3)), mes, int(f.group(1))) if mes else None
        except ValueError:
            pass
    if not fecha:
        obs.append("sin fecha en el documento")
    elif fecha.year != anio:
        obs.append(f"fecha {fecha:%d/%m/%Y} fuera del año {anio}")

    numero_documento = None
    if n := RE_NOTA.search(cabeza):
        numero_documento = f"{int(n.group(1)):04d}/{n.group(2)}"
        if int(n.group(1)) != numero:
            obs.append(f"el documento dice Nota N° {numero_documento}")
    elif texto.strip():
        obs.append("sin 'Nota N°' en el documento")
    else:
        obs.append("no se pudo leer el texto")

    expediente = None
    for fuente in (asunto, texto):
        if e := RE_EX.search(fuente):
            expediente = re.sub(r"\s+", "", e.group(0)).upper()
        elif e := RE_EXPTE.search(fuente):
            expediente = e.group(1)
        elif e := RE_NO.search(fuente):
            expediente = e.group(0).upper()
        if expediente:
            break

    cuerpo = texto[n.end():] if n else texto
    return {
        "anio": anio,
        "numero": numero,
        "sufijo": "",
        "clase": "Nota",
        "numero_documento": numero_documento or f"{numero:04d}/{str(anio)[-2:]}",
        "numero_expediente": expediente,
        "fecha_salida": fecha.isoformat() if fecha else None,
        "destinatario": destinatario,
        "asunto": base or asunto or nombre,
        "extracto": limpiar(cuerpo)[:1500] or None,
        "observaciones": obs,
        "ruta_archivo": ruta,
    }


def misma_nota(a, b):
    """True si dos archivos con el mismo número son la misma nota (mismo texto, o uno ilegible)."""
    ta, tb = (a["extracto"] or "")[:800].lower(), (b["extracto"] or "")[:800].lower()
    if len(ta) < 50 or len(tb) < 50:
        return True
    if "copy" in (a["ruta_archivo"] + b["ruta_archivo"]).lower():
        return True

    # Palabras en común: tolera los saltos de línea y el ruido del texto sacado de un PDF.
    def parecido(x, y):
        px, py = set(re.findall(r"\w{3,}", x)), set(re.findall(r"\w{3,}", y))
        return len(px & py) / max(1, len(px | py))

    # Mismo texto, o misma referencia y saludo (otra versión de la misma nota).
    return parecido(ta, tb) > 0.6 or parecido(ta[:80], tb[:80]) > 0.7


def resolver_repetidos(filas):
    """Une archivos de la misma nota y pone sufijo b, c... a notas distintas con el mismo número."""
    por_numero = defaultdict(list)
    for f in filas:
        por_numero[(f["anio"], f["numero"])].append(f)
    salida = []
    for grupo in por_numero.values():
        # Primero los Word (texto más confiable), las copias al final.
        grupo.sort(key=lambda f: (f["ruta_archivo"].lower().endswith(".pdf"), "copy" in f["ruta_archivo"].lower(),
                                  f["ruta_archivo"]))
        notas = []
        for f in grupo:
            igual = next((n for n in notas if misma_nota(n, f)), None)
            if igual:
                igual["observaciones"].append(f"otra versión o copia: {PurePosixPath(f['ruta_archivo']).name}")
            else:
                notas.append(f)
        for i, f in enumerate(notas):
            if len(notas) > 1:
                f["sufijo"] = "" if i == 0 else "bcdefgh"[i - 1]
                otras = [n["asunto"][:60] for n in notas if n is not f]
                f["observaciones"].append("número repetido con: " + " | ".join(otras))
            salida.append(f)
    for f in salida:
        f["observaciones"] = "; ".join(f["observaciones"]) or None
    return sorted(salida, key=lambda f: (f["anio"], f["numero"], f["sufijo"]))


def main():
    ap = argparse.ArgumentParser(description=__doc__)
    ap.add_argument("origen", type=Path, help="zip o carpeta con subcarpetas 'Notas AAAA'")
    ap.add_argument("--anio", type=int, action="append", help="importar solo este año (se puede repetir)")
    ap.add_argument("--grabar", action="store_true", help="grabar en Supabase (sin esto, solo la planilla)")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    archivos = leer_entrada(args.origen, set(args.anio or []))
    if not archivos:
        sys.exit("No se encontraron notas npNNNN dentro de carpetas 'Notas AAAA'.")

    lector = LectorTexto()
    filas = []
    try:
        for i, (anio, ruta, nombre, datos) in enumerate(archivos, 1):
            if int(RE_ARCHIVO.match(nombre).group(1)) == 0:
                continue  # np0000 / np00XX: modelo o sin número
            try:
                texto = lector.texto(nombre, datos)
            except Exception as e:
                texto = ""
                print(f"\n  ! {nombre}: no se pudo leer ({e})")
            filas.append(analizar(anio, ruta, nombre, texto))
            print(f"\r  {i}/{len(archivos)} leídas", end="", flush=True)
    finally:
        lector.cerrar()
    print()

    filas = resolver_repetidos(filas)
    revision = RAIZ / "revision"
    revision.mkdir(exist_ok=True)
    for anio in sorted({f["anio"] for f in filas}):
        del_anio = [f for f in filas if f["anio"] == anio]
        numeros = {f["numero"] for f in del_anio}
        faltan = [n for n in range(1, max(numeros) + 1) if n not in numeros]
        con_obs = [f for f in del_anio if f["observaciones"]]
        print(f"\n== {anio}: {len(del_anio)} notas, números {min(numeros)} a {max(numeros)}")
        print(f"Sin archivo: {faltan or 'ninguno'}")
        print(f"Con observaciones: {len(con_obs)}")
        for f in con_obs:
            print(f"  np{f['numero']:04d}{f['sufijo']}: {f['observaciones'][:150]}")
        csv_path = revision / f"salidas_{anio}.csv"
        with open(csv_path, "w", newline="", encoding="utf-8-sig") as fh:
            w = csv.DictWriter(fh, fieldnames=list(del_anio[0].keys()), delimiter=";")
            w.writeheader()
            w.writerows(del_anio)
        print(f"Planilla de revisión: {csv_path}")

    if args.grabar:
        from conexion import cliente
        tabla = cliente().table("salidas")
        for i in range(0, len(filas), 200):
            tabla.upsert(filas[i:i + 200], on_conflict="anio,numero,sufijo").execute()
        print(f"\nGrabadas {len(filas)} salidas en mesa_entradas.salidas")
    else:
        print("\nNo se grabó nada (falta --grabar).")


if __name__ == "__main__":
    main()
