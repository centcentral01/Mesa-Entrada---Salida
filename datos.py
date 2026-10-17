# datos.py — Lectura de mesa_entradas para la app. Trae las tablas completas (son chicas)
# y las cachea unos minutos; los filtros se aplican en pandas.
import pandas as pd
import streamlit as st

from conexion import cliente

PAGINA = 1000  # máximo de filas que devuelve la API por pedido


def _todas(tabla, columnas="*", orden="id"):
    db = cliente()
    filas, desde = [], 0
    while True:
        lote = db.table(tabla).select(columnas).order(orden).range(desde, desde + PAGINA - 1).execute().data
        filas.extend(lote)
        if len(lote) < PAGINA:
            return pd.DataFrame(filas)
        desde += PAGINA


@st.cache_data(ttl=300, show_spinner="Cargando salidas…")
def salidas():
    df = _todas("salidas")
    if df.empty:
        return df
    df["fecha_salida"] = pd.to_datetime(df["fecha_salida"]).dt.date
    df["nota"] = df["numero_documento"].fillna("") + df["sufijo"].map(lambda s: f" {s}" if s else "")
    return df.sort_values(["anio", "numero", "sufijo"], ascending=[False, False, True])


@st.cache_data(ttl=300, show_spinner="Cargando entradas…")
def entradas():
    df = _todas("v_entradas")
    if df.empty:
        return df
    for c in ("fecha_documento", "fecha_recepcion", "fecha_vencimiento"):
        df[c] = pd.to_datetime(df[c]).dt.date
    return df.sort_values(["fecha_recepcion", "id"], ascending=False)


def contiene(serie, texto):
    """Filtro de texto que ignora mayúsculas y acentos."""
    def normal(s):
        return (s.fillna("").str.lower().str.normalize("NFKD")
                .str.encode("ascii", "ignore").str.decode("ascii"))
    t = normal(pd.Series([texto])).iloc[0]
    return normal(serie).str.contains(t, regex=False)
