# -*- coding: utf-8 -*-
"""
Página: Mantenimiento

  · Un botón rehace el dashboard (pestaña Resumen) y aplica el formato y el mantenimiento del Sheet
    (cuadres/formato_sheet.py: columnas, colores, respaldos, protecciones, vistas por asesor, orden de pestañas).
  · Historial de cambios que hizo la app.
"""
import streamlit as st

import ui
import conexion
from comunes import aviso_excel
from cuadres.base import TAB_HISTORIAL

ui.hero("Mantenimiento", "Dashboard, formato y respaldos del Sheet, e historial de cambios.")

base = conexion.panel_base()
if base is None:
    st.stop()

try:
    from cuadres import columnas
    falta_orden = columnas.pendiente(base)
except Exception:
    falta_orden = []
if falta_orden:
    st.warning("Falta acomodar las columnas de " + ", ".join(f"«{h}»" for h in falta_orden)
               + " (fecha, nombre y teléfono al inicio). Presiona **Rehacer dashboard y formato**.",
               icon=":material/view_column:")

if st.session_state.get("msg_mant"):
    st.success("Listo:\n" + "\n".join(f"- {m}" for m in st.session_state.pop("msg_mant")) + aviso_excel(base))

with st.container(border=True):
    st.markdown("**Rehacer dashboard y formato**")
    st.caption("Rehace la pestaña Resumen (gráficas e indicadores) y da mantenimiento al Sheet: fecha, nombre y teléfono como "
               "primeras columnas (congeladas) y sin MARCA en Finalizados, columnas, colores y "
               "fechas, un solo respaldo por pestaña, columnas calculadas y Bitacora protegidas (solo advertencia), "
               "encabezados congelados, vistas por asesor, Finalizados ligados a su rechazo y pestañas en orden. "
               "Úsalo si agregaste un asesor o una financiera al catálogo, o una vez al mes. No duplica nada.")
    if st.button("Rehacer dashboard y formato", type="primary", key="btn_mant", icon=":material/build:"):
        with st.spinner("Ajustando el Sheet…"):
            try:
                from cuadres import formato_sheet
                st.session_state["msg_mant"] = formato_sheet.aplicar(base)
                conexion.recargar()
                st.rerun()
            except Exception as e:
                st.error(f"No se pudo aplicar: {e}")

st.markdown("#### :material/history: Historial de cambios")
if TAB_HISTORIAL in base.pestanas():
    h = base.leer(TAB_HISTORIAL).drop(columns="_fila", errors="ignore")
    if h.empty:
        st.info("Todavía no hay cambios.")
    else:
        h = h.iloc[::-1]
        origen = st.multiselect("Origen", sorted(h["Origen"].dropna().astype(str).unique()), key="f_hist",
                                placeholder="Todos")
        if origen:
            h = h[h["Origen"].isin(origen)]
        st.dataframe(h.head(500).astype(str), hide_index=True, width="stretch", height=520)
else:
    st.info("Todavía no hay cambios.")
