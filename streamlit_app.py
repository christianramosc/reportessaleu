# -*- coding: utf-8 -*-
"""Punto de entrada: herramientas de MG Colima (una página por herramienta)."""
import streamlit as st

import ui

st.set_page_config(page_title="Herramientas MG Colima", page_icon="📊", layout="wide")


def acceso_permitido():
    """Pide contraseña solo si existe APP_PASSWORD en los secrets."""
    try:
        clave = st.secrets.get("APP_PASSWORD", "")
    except Exception:
        clave = ""
    if not clave or st.session_state.get("acceso_ok"):
        return True
    st.title("Herramientas MG Colima")
    intento = st.text_input("Contraseña", type="password")
    if intento:
        if intento == clave:
            st.session_state["acceso_ok"] = True
            st.rerun()
        st.error("Contraseña incorrecta.")
    return False


if not acceso_permitido():
    st.stop()

ui.aplicar_estilo()

paginas = [
    st.Page("paginas/kpis_capacitacion.py", title="KPIs de capacitación",
            icon=":material/school:", url_path="kpis", default=True),
    st.Page("paginas/cuadres_reportes.py", title="Cuadres para reportes",
            icon=":material/fact_check:", url_path="cuadres"),
]
st.navigation(paginas, position="top").run()
