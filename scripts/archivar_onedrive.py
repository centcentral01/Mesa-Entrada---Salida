# scripts/archivar_onedrive.py — Copia a OneDrive los documentos de mesa de entradas del año y guarda la ruta.
#
#   OneDrive\Mesa de Entradas\<año>\Salidas\npNNNN - ....pdf      (desde correo/Notas <año>/)
#   OneDrive\Mesa de Entradas\<año>\Entradas\<fecha> - <asunto>\  (el correo como .msg + sus adjuntos)
#
# Outlook solo se lee (SaveAs del correo y de los adjuntos). Actualiza ruta_archivo en salidas y entradas
# con la ruta completa, para abrirla desde la app o desde un enlace. No pisa lo que ya está copiado.
#
# Uso: python scripts/archivar_onedrive.py [--anio 2026]

import argparse
import re
import shutil
import sys
from datetime import datetime
from pathlib import Path

import win32com.client

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ))
sys.path.insert(0, str(RAIZ / "scripts"))
from correos_entrantes import adjuntos_reales  # noqa: E402

ONEDRIVE = Path.home() / "Personal CENT" / "OneDrive - Universidad Tecnologica Nacional"
OL_MSG = 3
INVALIDOS = re.compile(r'[<>:"/\\|?*\x00-\x1f]')


def nombre_seguro(s, largo):
    return re.sub(r"\s+", " ", INVALIDOS.sub("-", s)).strip(" .-")[:largo].strip(" .-")


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--anio", type=int, default=datetime.now().year)
    args = ap.parse_args()
    sys.stdout.reconfigure(encoding="utf-8")
    from conexion import cliente
    db = cliente()

    base = ONEDRIVE / "Mesa de Entradas" / str(args.anio)
    dir_sal, dir_ent = base / "Salidas", base / "Entradas"
    dir_sal.mkdir(parents=True, exist_ok=True)
    dir_ent.mkdir(parents=True, exist_ok=True)

    # Salidas: los archivos ya bajados del correo, con el mismo nombre que figura en ruta_archivo.
    origen = RAIZ / "correo" / f"Notas {args.anio}"
    salidas = db.table("salidas").select("id,ruta_archivo").eq("anio", args.anio).execute().data
    copiadas = 0
    for s in salidas:
        nombre = Path(s["ruta_archivo"] or "").name
        if not nombre or not (origen / nombre).exists():
            continue
        destino = dir_sal / nombre
        if not destino.exists():
            shutil.copy2(origen / nombre, destino)
            copiadas += 1
        db.table("salidas").update({"ruta_archivo": str(destino)}).eq("id", s["id"]).execute()
    print(f"Salidas: {copiadas} copiadas a {dir_sal}")

    # Entradas por correo: una carpeta con el .msg y los adjuntos.
    ns = win32com.client.Dispatch("Outlook.Application").GetNamespace("MAPI")
    entradas = db.table("entradas").select("id,fecha_recepcion,asunto,correo_id") \
        .eq("medio", "Correo").not_.is_("correo_id", "null").execute().data
    for e in entradas:
        if not e["fecha_recepcion"].startswith(str(args.anio)):
            continue
        asunto = re.sub(r"^(oficio judicial|oficio)\s*[-:]*\s*", "", e["asunto"], flags=re.I)
        # Nombres cortos: la carpeta de OneDrive ya ocupa ~110 de los 260 caracteres de Windows.
        carpeta = dir_ent / nombre_seguro(f"{e['fecha_recepcion']} - {asunto}", 50)
        carpeta.mkdir(exist_ok=True)
        it = ns.GetItemFromID(e["correo_id"])
        if not (carpeta / "correo.msg").exists():
            it.SaveAs(str(carpeta / "correo.msg"), OL_MSG)
        reales = set(adjuntos_reales(it))
        for a in it.Attachments:
            if a.FileName in reales:
                nombre = Path(a.FileName)
                archivo = carpeta / (nombre_seguro(nombre.stem, 70) + nombre.suffix.lower())
                if not archivo.exists():
                    a.SaveAsFile(str(archivo))
        db.table("entradas").update({"ruta_archivo": str(carpeta)}).eq("id", e["id"]).execute()
        print(f"  {carpeta.name}: {', '.join(p.name for p in carpeta.iterdir())[:120]}")
    print(f"Entradas: {dir_ent}")


if __name__ == "__main__":
    main()
