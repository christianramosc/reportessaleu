# -*- coding: utf-8 -*-
"""Conexión de la app con la base de seguimiento (Google Sheet) y la Copia de Financiamiento.

Secrets (Streamlit Cloud → Settings → Secrets, o .streamlit/secrets.toml):
    SEGUIMIENTO_SHEET_URL   = "https://docs.google.com/spreadsheets/d/…"   # base de seguimiento
    FINANCIAMIENTO_SHEET_URL = "https://docs.google.com/spreadsheets/d/…"  # opcional
    [gcp_service_account]   # contenido del JSON de la cuenta de servicio
    type = "service_account"
    …

Sin secrets la app funciona en "modo Excel": subes el .xlsx de seguimiento, la app lo
actualiza y lo descargas al final.
"""
import pandas as pd
import streamlit as st

from cuadres.base import BaseGoogle, BaseExcel, TAB_CAPTURAS
from cuadres import fuentes


def _secret(nombre, defecto=None):
    try:
        return st.secrets.get(nombre, defecto)
    except Exception:
        return defecto


def google_configurado():
    return bool(_secret("SEGUIMIENTO_SHEET_URL")) and bool(_secret("gcp_service_account"))


@st.cache_resource(show_spinner="Conectando con el Google Sheet…")
def _base_google(url):
    return BaseGoogle(dict(_secret("gcp_service_account")), url)


def obtener_base():
    """La base activa (Google Sheet o Excel subido) o None. Siempre con datos frescos."""
    if google_configurado():
        url = _secret("SEGUIMIENTO_SHEET_URL")
        base = _base_google(url)
        if not isinstance(base, BaseGoogle):
            # la conexión guardada es de una versión anterior del código (la app se actualizó): rehacerla
            _base_google.clear()
            base = _base_google(url)
        base.refrescar()
        return base
    return st.session_state.get("base_excel")


def capturas_guardadas(base):
    if base is None or TAB_CAPTURAS not in base.pestanas():
        return pd.DataFrame(columns=fuentes.CAP_COLS + fuentes.CAP_EXTRA)
    df = base.leer(TAB_CAPTURAS).drop(columns="_fila", errors="ignore")
    fecha = df.get("Fecha de carga")
    df = df.drop(columns="Fecha de carga", errors="ignore").fillna("").astype(str)
    df["Fecha de carga"] = fecha if fecha is not None else ""
    return df


def tablas_financiamiento(base):
    """(tablas, titulo) de la Copia de Financiamiento si está configurada en Google; si no (None, None)."""
    url = _secret("FINANCIAMIENTO_SHEET_URL")
    if url and isinstance(base, BaseGoogle):
        return fuentes.tablas_de_google(base.cliente, url)
    return None, None


def panel_base(donde=st):
    """Muestra con qué base se trabaja. En modo Excel pide el archivo. Devuelve la base o None."""
    if google_configurado():
        try:
            base = obtener_base()
        except Exception as e:
            donde.error(f"No pude abrir el Google Sheet: {e}. Revisa que esté compartido con la cuenta de "
                        "servicio como Editor.")
            return None
        donde.success(f"Conectado a **{base.titulo}** · [abrir el Sheet]({base.url})", icon=":material/cloud_done:")
        return base
    caja = donde.container(border=True)
    caja.markdown("**Modo Excel (sin conexión a Google)** · Sube tu Excel de seguimiento; al final descargas la "
                  "versión actualizada. Para trabajar directo sobre el Google Sheet configura los *secrets* "
                  "(ver README).")
    f = caja.file_uploader("Excel de seguimiento (Seguimiento_Rechazos_MG_Colima.xlsx)", type=["xlsx"], key="up_base_excel")
    if f is not None and st.session_state.get("_base_excel_nombre") != (f.name, f.size):
        st.session_state["base_excel"] = BaseExcel(f.getvalue(), f.name)
        st.session_state["_base_excel_nombre"] = (f.name, f.size)
    base = st.session_state.get("base_excel")
    if base is not None:
        caja.download_button("⬇️ Descargar Excel actualizado", base.guardar(), file_name=base.titulo,
                             mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                             key=f"dl_base_{id(donde)}")
    return base
