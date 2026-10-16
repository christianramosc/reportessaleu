# -*- coding: utf-8 -*-
"""
Página: Base de seguimiento

Nutre el Google Sheet de seguimiento con las bitácoras mensuales de Cetelem y la
Bitácora de F&I: propone cambios, los muestra, eliges y se aplican.
"""
import streamlit as st

import ui
import conexion
from cuadres import fuentes, actualizar, pendientes
from cuadres.base import TAB_HISTORIAL

ui.hero("Base de seguimiento",
        "Actualiza el Google Sheet de rechazos con las bitácoras de Cetelem y la Bitácora de F&I.")

base = conexion.panel_base()
if base is None:
    st.stop()

tab_act, tab_pend, tab_hist = st.tabs(["🔄 Actualizar con bitácoras", "✅ Pendientes", "🕘 Historial de cambios"])

with tab_act:
    if st.session_state.get("msg_act"):
        st.success(st.session_state.pop("msg_act"))
    guardadas = conexion.capturas_guardadas(base)
    meses_guardados = list(dict.fromkeys(guardadas["Hoja"])) if not guardadas.empty else []
    c1, c2 = st.columns(2)
    archivos_cet = c1.file_uploader(
        "Bitácoras mensuales de Cetelem", type=["xlsx"], accept_multiple_files=True, key="up_cet_base",
        help="Las que te pasa el F&I (JUNIO_2026_FINAL.xlsx, Bitacora_Oct26…). Cada mes que subas reemplaza "
             "al mismo mes guardado en el Sheet.")
    c1.caption("Guardadas en el Sheet: " + (", ".join(meses_guardados) if meses_guardados else "ninguna todavía"))

    tablas_fi, titulo_fi = None, None
    try:
        tablas_fi, titulo_fi = conexion.tablas_financiamiento(base)
    except Exception as e:
        c2.warning(f"No pude leer la Bitácora de F&I de Google: {e}")
    if tablas_fi is not None:
        c2.info(f"Bitácora de F&I: se lee de Google Sheets (**{titulo_fi}**).", icon=":material/link:")
    else:
        f_fi = c2.file_uploader("Bitácora de F&I (opcional)", type=["xlsx"], key="up_fi_base",
                                help="La bitácora de todos los créditos (todas las financieras). Si configuras FINANCIAMIENTO_SHEET_URL en los secrets, "
                                     "se lee sola desde Google Sheets.")
        if f_fi is not None:
            tablas_fi, titulo_fi = fuentes.tablas_de_excel(f_fi.getvalue()), f_fi.name

    puede = bool(archivos_cet) or not guardadas.empty or tablas_fi is not None
    if st.button("Revisar cambios", type="primary", disabled=not puede, key="btn_revisar"):
        with st.spinner("Leyendo bitácoras y comparando con el Sheet…"):
            try:
                nuevas = fuentes.leer_capturas([(f.name, f.getvalue()) for f in archivos_cet or []])
                capturas = fuentes.combinar_capturas(guardadas, nuevas) if not nuevas.empty else guardadas
                regs = fuentes.registros_capturas(capturas)
                hojas_fi = []
                if tablas_fi is not None:
                    regs_fi, hojas_fi = fuentes.registros_financiamiento(tablas_fi)
                    regs += regs_fi
                sol = fuentes.consolidar(regs)
                prop, ctx = actualizar.proponer(base, sol)
                st.session_state["act"] = dict(prop=prop, ctx=ctx, sol=sol, capturas=capturas,
                                               meses=list(dict.fromkeys(capturas["Hoja"])) if not capturas.empty else [],
                                               hojas_fi=hojas_fi, cambia_capturas=not nuevas.empty)
            except Exception as e:
                st.session_state.pop("act", None)
                st.error(f"No se pudo revisar: {e}")

    act = st.session_state.get("act")
    if act:
        prop = act["prop"]
        st.caption("Cetelem: " + (", ".join(act["meses"]) or "—") + "  ·  F&I: " + (", ".join(act["hojas_fi"]) or "—"))
        m = st.columns(4)
        m[0].metric("Rechazos nuevos", int((prop["Tipo"] == "Nuevo rechazo").sum()))
        m[1].metric("Correcciones", int((prop["Tipo"] == "Corrección").sum()))
        m[2].metric("Comentarios", int((prop["Tipo"] == "Comentario").sum()))
        m[3].metric("Solicitudes en bitácoras", len(act["sol"]))
        if prop.empty:
            st.success("La base ya está al día con estas bitácoras.")
        else:
            st.markdown("Revisa los cambios. **Desmarca** los que no quieras aplicar; los desmarcados se anotan "
                        "como *ignorados* y no se vuelven a proponer.")
            vista = st.data_editor(
                prop[["Aplicar", "Tipo", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo"]].assign(
                    Fila=prop["Fila"].map(lambda f: "nueva" if f is None or f != f else str(int(f)))),
                hide_index=True, width="stretch", height=min(520, 60 + 35 * len(prop)), key="editor_prop",
                disabled=["Tipo", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo"],
                column_config={"Aplicar": st.column_config.CheckboxColumn(width="small"),
                               "Fila": st.column_config.TextColumn(width="small"),
                               "Motivo": st.column_config.TextColumn(width="large")})
            prop = prop.assign(Aplicar=vista["Aplicar"].values)
        n = int(prop["Aplicar"].sum()) if not prop.empty else 0
        etiqueta = f"Aplicar {n} cambio(s) en el Sheet" if n else "Guardar bitácoras en el Sheet"
        if st.button(etiqueta, type="primary", key="btn_aplicar",
                     disabled=prop.empty and not act["cambia_capturas"]):
            with st.spinner("Escribiendo en el Sheet…"):
                try:
                    r = actualizar.aplicar(base, prop, act["sol"], act["ctx"], act["capturas"])
                    st.session_state.pop("act", None)
                    st.session_state["msg_act"] = (
                        f"Listo: {r['nuevos']} rechazo(s) nuevo(s), {r['correcciones']} corrección(es), "
                        f"{r['ignorados']} ignorado(s)."
                        + (f" Respaldo previo en la pestaña oculta «{r['respaldo']}»." if r["respaldo"] else "")
                        + (" Descarga el Excel actualizado con el botón de arriba." if isinstance(base, conexion.BaseExcel) else ""))
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo aplicar: {e}")

with tab_pend:
    if st.session_state.get("msg_pend"):
        st.success(st.session_state.pop("msg_pend"))
    st.markdown("Revisa los pendientes de la pestaña **Revision**. Los que ya tienen el dato capturado se marcan "
                "como *Corregido* y se les quita el **morado y la nota** en la hoja. También quita la marca de los "
                "que cambiaste a mano a *Corregido* o *No aplica*.")
    if st.button("Revisar pendientes", type="primary", key="btn_rev_pend"):
        with st.spinner("Revisando pendientes contra el Sheet…"):
            try:
                st.session_state["pend"] = pendientes.revisar(base)
            except Exception as e:
                st.session_state.pop("pend", None)
                st.error(f"No se pudieron revisar: {e}")
    pend = st.session_state.get("pend")
    if pend is not None:
        if pend.empty:
            st.success("No hay pendientes en la pestaña Revision.")
        else:
            cuenta = pend["Resultado"].value_counts()
            k = st.columns(4)
            k[0].metric("Resueltos", int(cuenta.get("Resuelto", 0)))
            k[1].metric("Marcados a mano", int(cuenta.get("Quitar marca", 0)))
            k[2].metric("Siguen pendientes", int(cuenta.get("Sigue pendiente", 0)))
            k[3].metric("Revisar a mano", int(cuenta.get("Revisar a mano", 0)))
            orden = {"Resuelto": 0, "Quitar marca": 1, "Sigue pendiente": 2, "Revisar a mano": 3}
            vista = pend.sort_values("Resultado", key=lambda x: x.map(orden))
            colores = {"Resuelto": "#C6EFCE", "Quitar marca": "#C6EFCE", "Sigue pendiente": "#FFEB9C",
                       "Revisar a mano": "#EDEFF3"}
            st.dataframe(vista[["Resultado", "Hoja", "Celda", "Cliente", "Campo", "Detalle", "Problema"]].style.map(
                             lambda v: f"background-color: {colores.get(v, '')}" if v in colores else "",
                             subset=["Resultado"]),
                         hide_index=True, width="stretch", height=min(520, 60 + 35 * len(vista)))
            n = int(pend["Resultado"].isin(["Resuelto", "Quitar marca"]).sum())
            if st.button(f"Limpiar {n} pendiente(s) resuelto(s)", type="primary", disabled=n == 0, key="btn_limpiar"):
                with st.spinner("Quitando marcas y actualizando Revision…"):
                    try:
                        hechos = pendientes.aplicar(base, pend)
                        st.session_state.pop("pend", None)
                        st.session_state["msg_pend"] = (
                            f"Listo: {hechos} pendiente(s) limpiado(s); quedó en el Historial."
                            + (" Descarga el Excel actualizado con el botón de arriba." if isinstance(base, conexion.BaseExcel) else ""))
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo limpiar: {e}")

with tab_hist:
    if TAB_HISTORIAL in base.pestanas():
        h = base.leer(TAB_HISTORIAL).drop(columns="_fila", errors="ignore")
        if h.empty:
            st.info("Todavía no hay cambios registrados.")
        else:
            h = h.iloc[::-1]
            origen = st.multiselect("Origen", sorted(h["Origen"].dropna().astype(str).unique()), key="f_hist",
                                    placeholder="Todos los orígenes")
            if origen:
                h = h[h["Origen"].isin(origen)]
            st.dataframe(h.head(500).astype(str), hide_index=True, width="stretch", height=520)
    else:
        st.info("Todavía no hay cambios registrados.")
