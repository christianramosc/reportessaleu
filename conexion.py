# -*- coding: utf-8 -*-
"""Conexión de la app con la base de seguimiento (Google Sheet) y la Bitácora de F&I.

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

from cuadres.base import BaseGoogle, BaseExcel, TAB_CAPTURAS, TAB_BITACORA, TAB_FI
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
    if "Hoja" in df.columns:
        df["Hoja"] = df["Hoja"].map(fuentes.etiqueta_hoja)
    fecha = df.get("Fecha de carga")
    df = df.drop(columns="Fecha de carga", errors="ignore").fillna("").astype(str)
    df["Fecha de carga"] = fecha if fecha is not None else ""
    return df


def bitacora_guardada(base):
    """Solicitudes de todas las financieras para los cuadres: la pestaña Bitacora del Sheet
    (capturas de Cetelem + Bitácora de F&I). Si no existe, se arma con las capturas de Cetelem guardadas."""
    if base is None:
        return fuentes.solicitudes_de_bitacora(None)
    sol = fuentes.solicitudes_de_bitacora(None)
    if TAB_BITACORA in base.pestanas():
        sol = fuentes.solicitudes_de_bitacora(base.leer(TAB_BITACORA).drop(columns="_fila", errors="ignore"))
    if sol.empty:
        caps = capturas_guardadas(base)
        if caps.empty:
            return sol
        cons = fuentes.consolidar(fuentes.registros_capturas(caps))
        sol = fuentes.solicitudes_de_bitacora(pd.DataFrame(fuentes.filas_bitacora(cons, {}),
                                                           columns=fuentes.BITACORA_COLS))
    return _unificar_vendedores(base, sol)


def _unificar_vendedores(base, sol):
    """F&I escribe al asesor por su primer nombre ("CARLA"); se usa el nombre completo de Catalogos."""
    from cuadres import actualizar
    catalogo = actualizar._catalogo_asesores(base)
    if catalogo:
        sol["Vendedor"] = sol["Vendedor"].map(lambda v: actualizar._asesor(v, catalogo) or v)
    sol["Vendedor"] = sol["Vendedor"].map(lambda v: v.upper() if v else "(sin vendedor)")
    return sol


def registros_fi_guardados(base):
    """(registros, fecha de carga) de la copia de la Bitácora de F&I guardada en el Sheet."""
    if base is None or TAB_FI not in base.pestanas():
        return [], None
    df = base.leer(TAB_FI).drop(columns="_fila", errors="ignore")
    fecha = next((v for v in df.get("Fecha de carga", []) if v), None) if not df.empty else None
    return fuentes.registros_fi_de_tabla(df), fecha


def registros_fi(base):
    """(registros, de dónde): la Bitácora de F&I de Google si está configurada; si no, la copia guardada."""
    try:
        tablas, titulo = tablas_financiamiento(base)
    except Exception:
        tablas, titulo = None, None
    if tablas is not None:
        regs, _ = fuentes.registros_financiamiento(tablas)
        if regs:
            return regs, f"Google Sheets ({titulo})"
    regs, fecha = registros_fi_guardados(base)
    if regs:
        cuando = f", cargada el {fecha:%d/%m/%Y}" if hasattr(fecha, "strftime") else ""
        return regs, f"copia guardada en el Sheet{cuando}"
    return [], None


def bitacora_fi(base):
    """(solicitudes, de dónde) para el cuadre de Sale-U: las solicitudes de la Bitácora de F&I (todas las
    financieras), con el estatus oficial de Cetelem cuando también están en sus capturas."""
    regs_fi, origen = registros_fi(base)
    if not regs_fi:
        return fuentes.solicitudes_de_bitacora(None), None
    regs = fuentes.registros_capturas(capturas_guardadas(base)) + regs_fi
    sol = [s for s in fuentes.consolidar(regs) if fuentes.es_de_fi(s)]
    df = pd.DataFrame(fuentes.filas_bitacora(sol, {}), columns=fuentes.BITACORA_COLS)
    return _unificar_vendedores(base, fuentes.solicitudes_de_bitacora(df)), origen


def describir_bitacora(sol):
    """'121 solicitudes · may 2026 – oct 2026 · CETELEM 102, SCOTIA 3, …'"""
    if sol.empty:
        return "sin solicitudes"
    con = sol[sol["_anio"] > 0].sort_values(["_anio", "_mes"])
    rango = f"{con['Mes'].iloc[0].lower()} – {con['Mes'].iloc[-1].lower()}" if not con.empty else "sin fechas"
    fin = sol["Fin_sol"].value_counts()
    fin = ", ".join(f"{k} {v}" for k, v in fin.items())
    return f"{len(sol)} solicitudes · {rango} · {fin}"


def tablas_financiamiento(base):
    """(tablas, titulo) de la Bitácora de F&I si está configurada en Google; si no (None, None)."""
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
