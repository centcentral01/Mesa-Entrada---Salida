# scripts/entradas_correo.py — Arma las entradas del año a partir del correo de llombardi@ (solo lectura en Outlook).
#
# Fuentes:
#   - Bandeja de entrada: lo que reenvía Dirección Ejecutiva con un oficio (RV: OFICIO ..., - OFICIO -),
#     y las notas de organismos de afuera (sin las respuestas RE:/RV: de la misma conversación).
#   - Elementos enviados: los oficios que Luis le pasa a Neifert o a Dirección ("Oficio Judicial - ...").
# Une los correos de un mismo oficio por la carátula (apellidos/empresas en asunto y adjuntos) y busca
# con qué salida del año se respondió. El remitente de un oficio es el juzgado: se toma del destinatario
# de la salida que lo responde, si la hay.
#
# Sin --grabar solo arma revision/entradas_correo_<año>.csv. Con --grabar inserta las entradas
# (medio Correo, sin número de mesa) y vincula las salidas; si se corre de nuevo, saltea las ya cargadas
# (por correo_id).
#
# Uso: python scripts/entradas_correo.py [--anio 2026] [--grabar]

import argparse
import csv
import re
import sys
import unicodedata
from datetime import datetime
from pathlib import Path

import win32com.client

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(RAIZ / "scripts"))
from correos_entrantes import (CUENTA, EXCLUIR_DOMINIOS, EXCLUIR_NOMBRES, EXCLUIR_PROPIAS, RE_AUDITORIA,  # noqa: E402
                               RE_NO_NOTA, RE_NOTA, adjuntos_reales, almacen, correos, smtp)

DIRECCION = "direccionejecutiva@cent.utn.edu.ar"
RE_OFICIO = re.compile(r"oficio|c[ée]dula|juzgado|tribunal|requerimiento|procelac|fiscal", re.I)
RE_RESPUESTA = re.compile(r"^\s*((RE|RV|FW|Fwd)\s*:\s*)*(responde|contesta|envia)", re.I)
RE_RESPUESTA_HILO = re.compile(r"^\s*(RE|Fwd?)\s*:", re.I)
# Palabras que no distinguen un oficio de otro.
VACIAS = set("""oficio oficios judicial judiciales cent linea lineas danos perjuicios causa autos expte expediente
    responde contesta respuesta segundo nuevo otro otros otra otras despido juzgado juz tribunal trabajo civil
    comercial nacional primera instancia caba solicitud informacion entrada sres sr dra con para por del las los
    una pdf docx doc trans tran lesiones sin acc ley prov provincia plata secretaria cam camara instancia nro
    ordinario s.a srl s.r.l ahorro fines determinados transporte transportes""".split())


def normal(s):
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode().lower()
    return s


def claves(texto):
    """Apellidos y nombres propios de la carátula: palabras de 4+ letras que no son de relleno."""
    return {p for p in re.findall(r"[a-z]{4,}", normal(texto)) if p not in VACIAS}


def limpiar_asunto(s):
    return re.sub(r"^\s*((RE|RV|FW|Fwd)\s*:\s*)+", "", s or "", flags=re.I).strip(" -.")


def candidatos(ns, desde):
    st = almacen(ns, CUENTA)
    for it, rec in correos(st.Folders["Bandeja de entrada"], desde):
        dir_ = smtp(it)
        asunto, adj = it.Subject or "", adjuntos_reales(it)
        texto = asunto + " " + " ".join(adj)
        if dir_ == DIRECCION:
            if RE_OFICIO.search(texto) and not RE_RESPUESTA.search(asunto):
                yield "Dirección Ejecutiva", it, rec, adj
            continue
        usuario, _, dominio = dir_.partition("@")
        nombre = it.SenderName or ""
        if (dominio.endswith(EXCLUIR_DOMINIOS) or dir_ in EXCLUIR_PROPIAS or RE_AUDITORIA.match(usuario)
                or EXCLUIR_NOMBRES.search(usuario) or EXCLUIR_NOMBRES.search(nombre)):
            continue
        if RE_NOTA.search(texto) and not RE_NO_NOTA.search(asunto + " " + dir_) \
                and not RE_RESPUESTA_HILO.search(asunto):
            yield "Externo", it, rec, adj
    for it, rec in correos(st.Folders["Elementos enviados"], desde):
        asunto = it.Subject or ""
        if re.match(r"^\s*oficio", asunto, re.I) and it.Attachments.Count:
            yield "Enviado por Luis", it, rec, adjuntos_reales(it)


def salidas_del_anio(anio):
    from conexion import cliente
    return cliente().table("salidas").select("id,numero,sufijo,numero_documento,destinatario,asunto,fecha_salida,entrada_id") \
        .eq("anio", anio).execute().data


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--anio", type=int, default=datetime.now().year)
    ap.add_argument("--grabar", action="store_true")
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    desde = datetime(args.anio, 1, 1)

    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    correos_ = []
    for origen, it, rec, adj in candidatos(ns, desde):
        correos_.append({"origen": origen, "fecha": rec, "remitente": it.SenderName or "", "email": smtp(it),
                         "asunto": limpiar_asunto(it.Subject), "adjuntos": adj, "correo_id": it.EntryID,
                         "claves": claves((it.Subject or "") + " " + " ".join(adj))})
    correos_.sort(key=lambda c: c["fecha"])

    # Un mismo oficio puede aparecer reenviado por Dirección y enviado por Luis: se agrupa por carátula.
    entradas = []
    for c in correos_:
        grupo = next((e for e in entradas if len(e["claves"] & c["claves"]) >= 2
                      or (e["claves"] and e["claves"] == c["claves"])), None)
        if grupo:
            grupo["otros"].append(c)
            grupo["claves"] |= c["claves"]
        else:
            entradas.append({**c, "otros": []})

    salidas = salidas_del_anio(args.anio)
    for s in salidas:
        s["claves"] = claves(s["asunto"])
    for e in entradas:
        # Salidas posteriores a la entrada que comparten al menos un apellido de la carátula.
        e["salidas"] = [s for s in salidas if e["claves"] & s["claves"]
                        and (s["fecha_salida"] or "9999") >= f"{e['fecha']:%Y-%m-%d}"]

    filas = []
    print(f"{len(correos_)} correos candidatos → {len(entradas)} entradas\n")
    for e in entradas:
        resp = ", ".join(s["numero_documento"] + (f" {s['sufijo']}" if s["sufijo"] else "") for s in e["salidas"])
        # Un oficio lo manda el juzgado: si no hay salida que lo identifique, queda sin remitente
        # (Dirección o Luis solo lo reenviaron).
        if e["origen"] == "Externo":
            remitente = e["remitente"]
        else:
            remitente = next((s["destinatario"] for s in e["salidas"] if s["destinatario"]), None)
        clase = "Oficio" if RE_OFICIO.search(e["asunto"] + " ".join(e["adjuntos"])) else "Nota"
        asunto = e["asunto"]
        if len(asunto) < 25 and e["adjuntos"]:  # "Oficio Judicial" a secas: el adjunto dice más
            asunto = f"{asunto} - {Path(e['adjuntos'][0]).stem}"
        otros = "; ".join(f"{o['fecha']:%d/%m} {o['origen']}: {o['asunto'][:60]}" for o in e["otros"])
        print(f"{e['fecha']:%Y-%m-%d} | {e['origen']:19} | {clase:6} | {asunto[:75]:75} | resp: {resp or '-'}")
        if otros:
            print(f"{'':13}también: {otros[:150]}")
        filas.append({
            "fecha_recepcion": f"{e['fecha']:%Y-%m-%d}", "origen": e["origen"], "clase": clase,
            "remitente": remitente, "asunto": asunto, "adjuntos": " | ".join(e["adjuntos"]),
            "respondida_con": resp, "otros_correos": otros, "correo_id": e["correo_id"],
            "_salidas": [s["id"] for s in e["salidas"]],
        })

    salida = RAIZ / "revision" / f"entradas_correo_{args.anio}.csv"
    with open(salida, "w", newline="", encoding="utf-8-sig") as fh:
        campos = [k for k in filas[0] if not k.startswith("_")]
        w = csv.DictWriter(fh, fieldnames=campos, delimiter=";", extrasaction="ignore")
        w.writeheader()
        w.writerows(filas)
    print(f"\nPlanilla: {salida}")

    if not args.grabar:
        print("No se grabó nada (falta --grabar).")
        return
    from conexion import cliente
    db = cliente()
    cargados = {r["correo_id"] for r in db.table("entradas").select("correo_id").not_.is_("correo_id", "null")
                .execute().data}
    nuevas = 0
    for f in filas:
        if f["correo_id"] in cargados:
            continue
        fila = {
            "medio": "Correo", "anio": args.anio, "clase": f["clase"], "remitente": f["remitente"],
            "asunto": f["asunto"], "fecha_recepcion": f["fecha_recepcion"], "correo_id": f["correo_id"],
            "extracto": ("Adjuntos: " + f["adjuntos"]) if f["adjuntos"] else None,
            "observaciones": ("Otros correos del mismo tema: " + f["otros_correos"]) if f["otros_correos"] else None,
        }
        nueva = db.table("entradas").insert(fila).execute().data[0]
        nuevas += 1
        for sid in f["_salidas"]:
            # El trigger de salidas pasa la entrada a Respondido.
            db.table("salidas").update({"entrada_id": nueva["id"]}).eq("id", sid).is_("entrada_id", "null").execute()
    print(f"Grabadas {nuevas} entradas nuevas.")


if __name__ == "__main__":
    main()
