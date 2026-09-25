# -*- coding: utf-8 -*-
"""Punto de entrada: herramientas de MG Colima (una página por herramienta)."""
import streamlit as st

import ui

st.set_page_config(page_title="Herramientas MG Colima", page_icon="📊", layout="wide")
ui.aplicar_estilo()

paginas = [
    st.Page("paginas/kpis_capacitacion.py", title="KPIs de capacitación",
            icon=":material/school:", url_path="kpis", default=True),
    st.Page("paginas/cuadre_saleu.py", title="Cuadre Sale-U",
            icon=":material/call:", url_path="cuadre-saleu"),
]
st.navigation(paginas, position="top").run()
