# pages/2_entradas.py — Consulta de entradas: filtros por año, emisor (remitente), estado y medio.
# Abre en "Falta contestar" (Pendiente o En Trámite), con las vencidas primero.
from datetime import date

import streamlit as st

import datos

st.title("📥 Entradas")

df = datos.entradas()
if df.empty:
    st.info("Todavía no hay entradas cargadas. Se suman con las notas en papel escaneadas y los correos.")
    st.stop()

df["anio_recepcion"] = df["fecha_recepcion"].map(lambda f: f.year if f else None)
ESTADOS = ["Pendiente", "En Trámite", "Respondido", "Archivado"]

c1, c2, c3, c4 = st.columns([1, 2, 2, 1])
anios = sorted(df["anio_recepcion"].dropna().unique(), reverse=True)
sel_anios = c1.multiselect("Año", anios, default=anios[:1])
frecuentes = df["remitente"].dropna().value_counts().index.tolist()
sel_rem = c2.multiselect("Emisor (remitente)", frecuentes, placeholder="Todos")
sel_est = c3.multiselect("Estado", ESTADOS, default=["Pendiente", "En Trámite"])
sel_medio = c4.multiselect("Medio", ["Papel", "Correo"], placeholder="Todos")

c5, c6 = st.columns([3, 1])
texto = c5.text_input("Buscar en asunto, extracto, Nº de documento o expediente")
solo_venc = c6.checkbox("Solo vencidas")

v = df
if sel_anios:
    v = v[v["anio_recepcion"].isin(sel_anios)]
if sel_rem:
    v = v[v["remitente"].isin(sel_rem)]
if sel_est:
    v = v[v["estado"].isin(sel_est)]
if sel_medio:
    v = v[v["medio"].isin(sel_medio)]
if texto:
    v = v[datos.contiene(v["asunto"], texto) | datos.contiene(v["extracto"], texto)
          | datos.contiene(v["numero_documento"], texto) | datos.contiene(v["numero_expediente"], texto)]
if solo_venc:
    v = v[v["vencida"]]
v = v.sort_values(["vencida", "fecha_vencimiento", "fecha_recepcion"], ascending=[False, True, False])

m1, m2, m3 = st.columns(3)
m1.metric("Entradas", len(v))
m2.metric("Vencidas", int(v["vencida"].sum()))
m3.metric("Respondidas", int(v["respondida"].sum()))

columnas = {
    "numero": st.column_config.NumberColumn("Nº mesa", format="%d", width="small"),
    "medio": st.column_config.TextColumn("Medio", width="small"),
    "fecha_recepcion": st.column_config.DateColumn("Recibida", format="DD/MM/YYYY", width="small"),
    "remitente": st.column_config.TextColumn("Emisor", width="medium"),
    "clase": st.column_config.TextColumn("Clase", width="small"),
    "asunto": st.column_config.TextColumn("Asunto", width="large"),
    "estado": st.column_config.TextColumn("Estado", width="small"),
    "fecha_vencimiento": st.column_config.DateColumn("Vence", format="DD/MM/YYYY", width="small"),
    "vencida": st.column_config.CheckboxColumn("Vencida", width="small"),
    "respuestas": st.column_config.TextColumn("Respondida con", width="medium"),
}
st.dataframe(v[list(columnas)], column_config=columnas, hide_index=True, width="stretch", height=480)

st.download_button(
    "⬇️ Descargar lo filtrado (CSV)",
    v.drop(columns=["id", "created_at", "updated_at"], errors="ignore")
     .to_csv(index=False, sep=";").encode("utf-8-sig"),
    file_name=f"entradas_{date.today():%Y%m%d}.csv", mime="text/csv",
)
