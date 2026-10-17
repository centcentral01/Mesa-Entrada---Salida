# app.py — Punto de entrada de la app Streamlit "Mesa de Entradas" (CENT).
# Consulta de notas de salida y de entradas guardadas en Supabase (esquema mesa_entradas).
# Correr con: streamlit run app.py

import sys
from pathlib import Path

ruta_base = Path(__file__).resolve().parent
if str(ruta_base) not in sys.path:
    sys.path.insert(0, str(ruta_base))

import streamlit as st

st.set_page_config(page_title="Mesa de Entradas", page_icon="📬", layout="wide")

pg_salidas = st.Page("pages/1_salidas.py", title="Notas de salida", icon="📤", default=True)
pg_entradas = st.Page("pages/2_entradas.py", title="Entradas", icon="📥")

with st.sidebar:
    st.markdown("**📬 Mesa de Entradas — CENT**")
    if st.button("🔄 Actualizar datos", width="stretch"):
        st.cache_data.clear()

st.navigation([pg_salidas, pg_entradas]).run()
