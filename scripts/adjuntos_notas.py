# scripts/adjuntos_notas.py — Baja de Outlook (llombardi@) los adjuntos que son notas de salida del año.
#
# Recorre todas las carpetas del buzón (solo lectura) y guarda en correo/Notas <año>/ los adjuntos
# con nombre npNNNN..., "Nota CENT N° NNNN-AA" o "NOTA CENT NRO. NN-AAAA". Los deja con nombre
# "npNNNN - <asunto>.<ext>" para que los lea scripts/importar_salidas.py. Si el nombre del adjunto
# no trae asunto (NP0019~1.PDF), usa el asunto del correo. No duplica: mismo nombre y tamaño se saltea.
#
# Además arma revision/adjuntos_notas_<año>.csv: qué nota vino en qué correo (fecha, remitente, asunto).
#
# Uso: python scripts/adjuntos_notas.py [--anio 2026]

import argparse
import csv
import re
import sys
import tempfile
from datetime import datetime
from pathlib import Path

import win32com.client

sys.path.insert(0, str(Path(__file__).resolve().parent))
from correos_entrantes import CUENTA, OL_MAIL, almacen, smtp  # noqa: E402

RAIZ = Path(__file__).resolve().parent.parent
RE_NP = re.compile(r"^np\s?(\d{1,4})(?:~\d)?\s*-?\s*(.*?)\.(docx?|pdf)$", re.I)
RE_NOTA_CENT = re.compile(r"^nota\s+(?:oficial\s+)?cent\s*(?:n[°ºro.]*)\s*(\d{1,4})\s*[-/]\s*(\d{2,4})(.*?)\.(docx?|pdf)$", re.I)
INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def carpetas(raiz):
    for f in raiz.Folders:
        yield f
        yield from carpetas(f)


def nombre_nota(adjunto, asunto_correo, anio):
    """'npNNNN - asunto.ext' si el adjunto es una nota de salida del año; si no, None."""
    nombre = adjunto.strip()
    if m := RE_NP.match(nombre):
        numero, resto, ext = int(m.group(1)), m.group(2), m.group(3)
    elif (m := RE_NOTA_CENT.match(nombre)) and int(m.group(2)) % 100 == anio % 100:
        numero, resto, ext = int(m.group(1)), f"Nota CENT N° {m.group(1)}-{m.group(2)}{m.group(3)}", m.group(4)
    else:
        return None
    if numero == 0:
        return None
    resto = resto.strip(" -~")
    if len(resto) < 4:  # nombre corto de 8.3 o sin asunto: se usa el del correo
        resto = re.sub(r"^(RV|RE|FW|Fwd)\s*:\s*", "", asunto_correo, flags=re.I).strip(" .-")
    resto = INVALIDOS.sub("-", resto)[:120].strip(" .")
    return f"np{numero:04d} - {resto}.{ext.lower()}"


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--anio", type=int, default=datetime.now().year)
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")

    destino = RAIZ / "correo" / f"Notas {args.anio}"
    destino.mkdir(parents=True, exist_ok=True)
    desde = datetime(args.anio, 1, 1)

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    registro, nuevos = [], 0
    with tempfile.TemporaryDirectory() as tmp:
        for carpeta in carpetas(almacen(ns, CUENTA)):
            try:
                items = carpeta.Items
                items = items.Restrict(f"[ReceivedTime] >= '{desde:%d/%m/%Y} 00:00'")
            except Exception:
                continue
            for it in items:
                if getattr(it, "Class", None) != OL_MAIL or not it.Attachments.Count:
                    continue
                for a in it.Attachments:
                    final = nombre_nota(a.FileName or "", it.Subject or "", args.anio)
                    if not final:
                        continue
                    ruta = destino / final
                    temp = Path(tmp) / "adjunto"
                    a.SaveAsFile(str(temp))
                    datos = temp.read_bytes()
                    if not ruta.exists():
                        ruta.write_bytes(datos)
                        nuevos += 1
                    elif ruta.read_bytes() != datos:
                        i = 2
                        while (otra := ruta.with_name(f"{ruta.stem} ({i}){ruta.suffix}")).exists() \
                                and otra.read_bytes() != datos:
                            i += 1
                        if not otra.exists():
                            otra.write_bytes(datos)
                            nuevos += 1
                        ruta = otra
                    registro.append({
                        "numero": int(final[2:6]), "archivo": ruta.name, "adjunto_original": a.FileName,
                        "carpeta": carpeta.Name, "fecha": f"{it.ReceivedTime:%Y-%m-%d %H:%M}",
                        "remitente": it.SenderName, "email": smtp(it), "asunto_correo": it.Subject,
                        "correo_id": it.EntryID,
                    })

    registro.sort(key=lambda r: (r["numero"], r["fecha"]))
    numeros = sorted({r["numero"] for r in registro})
    print(f"Adjuntos de notas {args.anio}: {len(registro)} en correos, {len(set(r['archivo'] for r in registro))} archivos "
          f"distintos ({nuevos} nuevos) en {destino}")
    print(f"Números: {numeros}")
    salida = RAIZ / "revision" / f"adjuntos_notas_{args.anio}.csv"
    salida.parent.mkdir(exist_ok=True)
    with open(salida, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(registro[0].keys()), delimiter=";")
        w.writeheader()
        w.writerows(registro)
    print(f"Registro: {salida}")


if __name__ == "__main__":
    main()
