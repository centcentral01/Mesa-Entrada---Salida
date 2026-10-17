# pages/1_salidas.py — Consulta de notas de salida: filtros por año, destinatario, texto,
# número, expediente y fechas; detalle de la nota elegida y descarga de lo filtrado.
from datetime import date

import pandas as pd
import streamlit as st

import datos

st.title("📤 Notas de salida")

df = datos.salidas()
if df.empty:
    st.info("Todavía no hay notas de salida cargadas.")
    st.stop()

# --- Filtros -------------------------------------------------------------------
anios = sorted(df["anio"].unique(), reverse=True)
c1, c2, c3 = st.columns([1, 2, 2])
sel_anios = c1.multiselect("Año", anios, default=anios[:1])
# Destinatarios ordenados por cantidad de notas (los más frecuentes primero).
frecuentes = df["destinatario"].dropna().value_counts().index.tolist()
sel_dest = c2.multiselect("Destinatario", frecuentes, placeholder="Todos")
dest_txt = c3.text_input("Destinatario contiene", placeholder="ej. juzgado, DAJ, SSTA")

c4, c5, c6, c7 = st.columns([2, 1, 1, 1])
texto = c4.text_input("Buscar en asunto y texto de la nota", placeholder="ej. LINEA 17, oblea, Ibarra")
numero = c5.number_input("Nº de nota", min_value=0, step=1, value=0, help="0 = todas")
expte = c6.text_input("Expediente / referencia", placeholder="ej. EX-2025-85677674")
solo_obs = c7.checkbox("Solo para revisar", help="Notas con observaciones del importador")

with st.expander("Fechas"):
    f1, f2 = st.columns(2)
    desde = f1.date_input("Desde", value=None, format="DD/MM/YYYY")
    hasta = f2.date_input("Hasta", value=None, format="DD/MM/YYYY")

# --- Aplicar -------------------------------------------------------------------
v = df
if sel_anios:
    v = v[v["anio"].isin(sel_anios)]
if sel_dest:
    v = v[v["destinatario"].isin(sel_dest)]
if dest_txt:
    v = v[datos.contiene(v["destinatario"], dest_txt) | datos.contiene(v["asunto"], dest_txt)]
if texto:
    v = v[datos.contiene(v["asunto"], texto) | datos.contiene(v["extracto"], texto)]
if numero:
    v = v[v["numero"] == numero]
if expte:
    v = v[datos.contiene(v["numero_expediente"], expte) | datos.contiene(v["extracto"], expte)]
if solo_obs:
    v = v[v["observaciones"].notna()]
if desde:
    v = v[v["fecha_salida"].notna() & (v["fecha_salida"] >= desde)]
if hasta:
    v = v[v["fecha_salida"].notna() & (v["fecha_salida"] <= hasta)]

# --- Resultado -----------------------------------------------------------------
m1, m2, m3 = st.columns(3)
m1.metric("Notas", len(v))
m2.metric("Destinatarios distintos", v["destinatario"].nunique())
fechas = v["fecha_salida"].dropna()
m3.metric("Período", f"{min(fechas):%d/%m/%Y} – {max(fechas):%d/%m/%Y}" if len(fechas) else "—")

columnas = {
    "nota": st.column_config.TextColumn("Nota", width="small"),
    "fecha_salida": st.column_config.DateColumn("Fecha", format="DD/MM/YYYY", width="small"),
    "destinatario": st.column_config.TextColumn("Destinatario", width="medium"),
    "asunto": st.column_config.TextColumn("Asunto", width="large"),
    "numero_expediente": st.column_config.TextColumn("Expediente / ref.", width="medium"),
    "observaciones": st.column_config.TextColumn("Observaciones", width="medium"),
}
sel = st.dataframe(
    v[list(columnas)], column_config=columnas, hide_index=True, width="stretch",
    height=480, on_select="rerun", selection_mode="single-row", key="tabla_salidas",
)

st.download_button(
    "⬇️ Descargar lo filtrado (CSV)",
    v.drop(columns=["id", "created_at", "updated_at"], errors="ignore")
     .to_csv(index=False, sep=";").encode("utf-8-sig"),
    file_name=f"salidas_{date.today():%Y%m%d}.csv", mime="text/csv",
)

# --- Detalle -------------------------------------------------------------------
filas = sel.selection.rows
if filas:
    n = v.iloc[filas[0]]
    st.divider()
    st.subheader(f"Nota {n['nota']}")
    d1, d2, d3 = st.columns(3)
    d1.markdown(f"**Fecha:** {n['fecha_salida']:%d/%m/%Y}" if pd.notna(n["fecha_salida"]) else "**Fecha:** —")
    d2.markdown(f"**Destinatario:** {n['destinatario'] or '—'}")
    d3.markdown(f"**Expediente / ref.:** {n['numero_expediente'] or '—'}")
    st.markdown(f"**Asunto:** {n['asunto']}")
    if n["observaciones"]:
        st.warning(n["observaciones"])
    st.markdown("**Texto (comienzo de la nota):**")
    st.text(n["extracto"] or "—")
    st.caption(f"Archivo: {n['ruta_archivo']}")
else:
    st.caption("Tocá una fila para ver el detalle de la nota.")
