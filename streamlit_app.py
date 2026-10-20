# -*- coding: utf-8 -*-
"""Punto de entrada: herramientas de MG Colima (una página por herramienta)."""
import hashlib
import pathlib
import sys

import streamlit as st

import zona  # noqa: F401  (hora de México: el servidor corre en UTC)

_RAIZ = pathlib.Path(__file__).resolve().parent


def _recargar_si_cambio():
    """Streamlit Cloud a veces actualiza el código sin reiniciar la app y los módulos (cuadres/, conexion…)
    se quedan con la versión vieja en memoria mientras las páginas ya son nuevas. Si cambió algún .py,
    se descargan para que se importen otra vez."""
    h = hashlib.sha1()
    for p in sorted(_RAIZ.glob("*.py")) + sorted(_RAIZ.glob("cuadres/*.py")):
        h.update(p.name.encode())
        h.update(p.read_bytes())
    firma = h.hexdigest()
    previa = getattr(sys, "_firma_app_mg", None)
    if previa and previa != firma:
        for nombre, modulo in list(sys.modules.items()):
            archivo = getattr(modulo, "__file__", None)
            if nombre != "__main__" and archivo and pathlib.Path(archivo).resolve().is_relative_to(_RAIZ):
                del sys.modules[nombre]
    sys._firma_app_mg = firma


_recargar_si_cambio()

import ui  # noqa: E402  (después de recargar)

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
    st.Page("paginas/inicio.py", title="Inicio", icon=":material/home:", url_path="inicio", default=True),
    st.Page("paginas/actualizar.py", title="Actualizar", icon=":material/upload_file:", url_path="actualizar"),
    st.Page("paginas/seguimiento.py", title="Seguimiento", icon=":material/fact_check:", url_path="seguimiento"),
    st.Page("paginas/blueservice.py", title="BlueService", icon=":material/directions_car:", url_path="blueservice"),
    st.Page("paginas/mantenimiento.py", title="Mantenimiento", icon=":material/build:", url_path="mantenimiento"),
    st.Page("paginas/kpis_capacitacion.py", title="KPIs de capacitación",
            icon=":material/school:", url_path="kpis"),
]
st.navigation(paginas, position="top").run()
