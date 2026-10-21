# -*- coding: utf-8 -*-
"""
Página: Mantenimiento

  · Un botón rehace el dashboard (pestaña Resumen) y aplica el formato y el mantenimiento del Sheet
    (cuadres/formato_sheet.py: columnas, colores, respaldos, protecciones, vistas por asesor, orden de pestañas).
  · Cruce de Créditos Finalizados con la pestaña Polizas (cuadres/cruce_polizas.py).
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

with st.container(border=True):
    st.markdown("**Cruce de Créditos Finalizados con Pólizas**")
    st.caption("Compara «MG Colima Créditos Finalizados» con la pestaña «Polizas» (la llena la app de seguros) y "
               "propone corregir teléfono, nombre, correo y VIN. Liga cada cliente por VIN, por teléfono con nombre "
               "parecido o por nombre. Ya vienen marcados los datos que faltaban o estaban mal; los que son "
               "distintos en las dos fuentes vienen sin marcar, para que decidas tú. Lo que desmarques no se vuelve "
               "a proponer.")
    if st.button("Revisar contra Pólizas", key="btn_cruce_pol", icon=":material/compare_arrows:"):
        with st.spinner("Cruzando con Pólizas…"):
            try:
                from cuadres import cruce_polizas
                st.session_state["cruce_pol"] = cruce_polizas.proponer(base)
            except Exception as e:
                st.session_state.pop("cruce_pol", None)
                st.error(f"No se pudo revisar: {e}")
    prop_pol = st.session_state.get("cruce_pol")
    if prop_pol is not None:
        if prop_pol.empty:
            st.success("Créditos Finalizados ya coincide con Pólizas.")
        else:
            m = st.columns(3)
            m[0].metric("Cambios propuestos", len(prop_pol))
            m[1].metric("Marcados", int(prop_pol["Aplicar"].sum()))
            m[2].metric("Por decidir", int((~prop_pol["Aplicar"].astype(bool)).sum()),
                        help="Dato distinto en Finalizados y en la póliza: márcalo si el de la póliza es el bueno.")
            vista = st.data_editor(
                prop_pol.drop(columns="_valor"), hide_index=True, width="stretch", key="editor_cruce_pol",
                height=min(460, 60 + 35 * len(prop_pol)),
                disabled=["Fila", "Cliente", "Campo", "Antes", "Después", "Ligado por", "Motivo"],
                column_config={"Aplicar": st.column_config.CheckboxColumn(width="small"),
                               "Motivo": st.column_config.TextColumn(width="large")})
            prop_pol = prop_pol.assign(Aplicar=vista["Aplicar"].values)
            n = int(prop_pol["Aplicar"].sum())
            if st.button(f"Aplicar {n} cambio(s)", type="primary", key="btn_aplicar_pol", icon=":material/done:",
                         disabled=not len(prop_pol)):
                with st.spinner("Escribiendo en el Sheet…"):
                    try:
                        from cuadres import cruce_polizas
                        escritos, ign = cruce_polizas.aplicar(base, prop_pol)
                        st.session_state.pop("cruce_pol", None)
                        st.session_state["msg_mant"] = [f"Créditos Finalizados: {escritos} dato(s) corregido(s) con "
                                                        f"Pólizas" + (f"; {ign} ignorado(s)" if ign else "")]
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
