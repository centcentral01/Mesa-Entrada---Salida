# scripts/correos_entrantes.py — Lista los correos recibidos en llombardi@ que pueden ser notas entrantes.
#
# Solo lectura sobre Outlook (no mueve ni marca nada). Deja afuera lo que es circuito interno:
#   - remitentes de @cent.utn.edu.ar (salvo INCLUIR_SIEMPRE) y auditorías regionales (ar010@ ... ar140@)
#   - personas de la CENT o de la CNRT que escriben por otros temas (EXCLUIR_NOMBRES)
#   - la CCIV, las casillas propias de Luis y los auditores que también escriben a viaticoscent@
# Lo que queda se marca como "probable nota" si el asunto o los adjuntos lo sugieren.
#
# Uso: python scripts/correos_entrantes.py [--desde 2026-01-01]
# Resultado: revision/correos_entrantes_<año>.csv

import argparse
import csv
import re
import sys
from collections import Counter
from datetime import datetime
from pathlib import Path

import win32com.client

RAIZ = Path(__file__).resolve().parent.parent
CUENTA = "llombardi@cent.utn.edu.ar"
CUENTA_VIATICOS = "viaticoscent@cent.utn.edu.ar"
OL_MAIL = 43
PR_SMTP = "http://schemas.microsoft.com/mapi/proptag/0x5D01001F"     # PR_SENDER_SMTP_ADDRESS
PR_ATTACH_CONTENT_ID = "http://schemas.microsoft.com/mapi/proptag/0x3712001F"

EXCLUIR_DOMINIOS = ("cent.utn.edu.ar", "cciv.com.ar", "cciv.org.ar")
# De la CENT solo cuentan Neifert y Dirección Ejecutiva: por ellos llegan los oficios y sus respuestas.
INCLUIR_SIEMPRE = {"vneifert@cent.utn.edu.ar", "direccionejecutiva@cent.utn.edu.ar"}
EXCLUIR_PROPIAS = {"luislrdi@gmail.com", "llombardicopia@gmail.com"}
# Usuario (parte antes de @) o nombre visible: CENT por otros dominios y la gente de CNRT de todos los días.
EXCLUIR_NOMBRES = re.compile(r"amartinez|liliana|facundo|nahuel|nrodegher|cciv", re.I)
RE_AUDITORIA = re.compile(r"^ar\s?0*(1[0-4]0|[1-9]0)\b", re.I)   # ar010 ... ar140
RE_NOTA = re.compile(r"\b(oficio|nota|expte|expediente|c[ée]dula|juzgado|tribunal|fiscal[ií]a|intimaci|"
                     r"requerimiento|pedido de informe|solicitud|reclamo|carta documento|NO-\d{4}|EX-\d{4})", re.I)
RE_NO_NOTA = re.compile(r"(newsletter|webinar|unsubscribe|promo|oferta|resumen de cuenta|transferencia|"
                        r"verificaci[oó]n|c[oó]digo de seguridad|microsoft|fortinet|linkedin)", re.I)


def almacen(ns, cuenta):
    return next(s for s in ns.Folders if s.Name.lower() == cuenta)


def smtp(item):
    """Dirección real del remitente (las internas vienen como /O=EXCHANGELABS/...)."""
    try:
        dir_ = item.PropertyAccessor.GetProperty(PR_SMTP)
        if dir_:
            return dir_.lower()
    except Exception:
        pass
    try:
        if item.SenderEmailType == "EX":
            return item.Sender.GetExchangeUser().PrimarySmtpAddress.lower()
    except Exception:
        pass
    return (item.SenderEmailAddress or "").lower()


def adjuntos_reales(item):
    """Nombres de adjuntos, sin las imágenes de firma incrustadas."""
    nombres = []
    for a in item.Attachments:
        try:
            if a.PropertyAccessor.GetProperty(PR_ATTACH_CONTENT_ID):
                continue
        except Exception:
            pass
        nombres.append(a.FileName)
    return nombres


def correos(carpeta, desde):
    items = carpeta.Items
    items.Sort("[ReceivedTime]", True)
    for it in items:
        if it.Class != OL_MAIL:
            continue
        recibido = it.ReceivedTime.replace(tzinfo=None)
        if recibido < desde:
            break
        yield it, recibido


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--desde", default=f"{datetime.now().year}-01-01")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    desde = datetime.fromisoformat(args.desde)

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")

    # Auditores: quien escribe a viaticoscent@ manda lo mismo a llombardi@.
    de_viaticos = {smtp(it) for it, _ in correos(almacen(ns, CUENTA_VIATICOS).Folders["Bandeja de entrada"], desde)}

    excluidos, filas = Counter(), []
    for it, recibido in correos(almacen(ns, CUENTA).Folders["Bandeja de entrada"], desde):
        dir_ = smtp(it)
        usuario, _, dominio = dir_.partition("@")
        nombre = it.SenderName or ""
        motivo = (
            None if dir_ in INCLUIR_SIEMPRE else
            "dominio CENT/CCIV" if dominio.endswith(EXCLUIR_DOMINIOS) else
            "auditoría regional" if RE_AUDITORIA.match(usuario) or RE_AUDITORIA.match(nombre) else
            "casilla propia" if dir_ in EXCLUIR_PROPIAS else
            "CENT/CNRT habitual" if EXCLUIR_NOMBRES.search(usuario) or EXCLUIR_NOMBRES.search(nombre) else
            "auditor (escribe a viaticoscent)" if dir_ in de_viaticos else None
        )
        if motivo:
            excluidos[motivo] += 1
            continue
        asunto = it.Subject or ""
        adj = adjuntos_reales(it)
        texto = asunto + " " + " ".join(adj)
        probable = bool(RE_NOTA.search(texto)) and not RE_NO_NOTA.search(asunto + " " + dir_)
        filas.append({
            "fecha": f"{recibido:%Y-%m-%d %H:%M}",
            "remitente": nombre,
            "email": dir_,
            "asunto": asunto,
            "adjuntos": " | ".join(adj),
            "probable_nota": "sí" if probable else "",
            "correo_id": it.EntryID,
        })

    print(f"Desde {desde:%d/%m/%Y}: {sum(excluidos.values()) + len(filas)} correos en la bandeja de {CUENTA}")
    for m, n in excluidos.most_common():
        print(f"  excluidos por {m}: {n}")
    probables = [f for f in filas if f["probable_nota"]]
    print(f"  quedan {len(filas)}; probables notas: {len(probables)}\n")
    for f in probables:
        print(f"  {f['fecha'][:10]} | {f['remitente'][:30]:30} | {f['email'][:34]:34} | {f['asunto'][:70]}"
              + (f" | adj: {f['adjuntos'][:60]}" if f["adjuntos"] else ""))
    print("\nResto (no parecen notas), por remitente:")
    for (r, e), n in Counter((f["remitente"], f["email"]) for f in filas if not f["probable_nota"]).most_common():
        print(f"  {n:3} {r[:30]:30} {e}")

    salida = RAIZ / "revision" / f"correos_entrantes_{desde.year}.csv"
    salida.parent.mkdir(exist_ok=True)
    with open(salida, "w", newline="", encoding="utf-8-sig") as fh:
        w = csv.DictWriter(fh, fieldnames=list(filas[0].keys()) if filas else ["fecha"], delimiter=";")
        w.writeheader()
        w.writerows(filas)
    print(f"\nPlanilla: {salida}")


if __name__ == "__main__":
    main()
