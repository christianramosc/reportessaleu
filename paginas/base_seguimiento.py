# -*- coding: utf-8 -*-
"""
Página: Base de seguimiento

Nutre el Google Sheet de seguimiento con las bitácoras mensuales de Cetelem y la
Bitácora de F&I: propone cambios, los muestra, eliges y se aplican.
"""
import streamlit as st

import ui
import conexion
from cuadres import fuentes, actualizar, pendientes, saleu_sheet
from cuadres.base import TAB_HISTORIAL, TAB_FINALIZADOS

ui.hero("Base de seguimiento", "Actualiza el Sheet de rechazos con las bitácoras de Cetelem y F&I.")
ui.guia_mes()

base = conexion.panel_base()
if base is None:
    st.stop()



def aviso_excel():
    return " Descarga el Excel arriba." if isinstance(base, conexion.BaseExcel) else ""


tab_act, tab_pend, tab_dash, tab_hist = st.tabs(["🔄 Actualizar", "✅ Pendientes", "📊 Dashboard", "🕘 Historial"])

with tab_dash:
    if st.session_state.get("msg_dash"):
        st.success(st.session_state.pop("msg_dash"))
    st.markdown("La pestaña **Resumen** se actualiza sola (son fórmulas; el año va en B2). Rehazla si agregaste "
                "un asesor o una financiera al catálogo.")
    if st.button("Rehacer dashboard", type="primary", key="btn_dash"):
        with st.spinner("Escribiendo el dashboard…"):
            try:
                from cuadres import dashboard
                n = dashboard.construir(base)
                st.session_state["msg_dash"] = f"Listo: Resumen con {n} gráficas." + aviso_excel()
                st.rerun()
            except Exception as e:
                st.error(f"No se pudo hacer el dashboard: {e}")

    st.divider()
    st.markdown("**Formato del Sheet**")
    if st.session_state.get("msg_fmt"):
        st.success("Listo:\n" + "\n".join(f"- {m}" for m in st.session_state.pop("msg_fmt")) + aviso_excel())
    st.caption("Ajusta columnas, colores y fechas, y da mantenimiento: borra respaldos viejos, protege las columnas "
               "calculadas y la Bitacora (solo advertencia), congela encabezados, crea vistas por asesor, liga "
               "Finalizados con su rechazo y ordena las pestañas. Si ya está aplicado, no duplica nada.")
    if st.button("Aplicar formato", key="btn_formato", icon=":material/format_paint:"):
        with st.spinner("Ajustando el Sheet…"):
            try:
                from cuadres import formato_sheet
                st.session_state["msg_fmt"] = formato_sheet.aplicar(base)
                st.rerun()
            except Exception as e:
                st.error(f"No se pudo aplicar el formato: {e}")

with tab_act:
    if st.session_state.get("msg_act"):
        st.success(st.session_state.pop("msg_act"))
    if st.session_state.get("aviso_act"):
        st.warning(st.session_state.pop("aviso_act"))
    guardadas = conexion.capturas_guardadas(base)
    meses_guardados = list(dict.fromkeys(guardadas["Hoja"])) if not guardadas.empty else []
    c1, c2 = st.columns(2)
    archivos_cet = c1.file_uploader(
        "Bitácoras de Cetelem", type=["xlsx"], accept_multiple_files=True, key="up_cet_base",
        help="Las que te pasa el F&I. Cada mes reemplaza al mismo mes guardado.")
    c1.caption("Guardadas: " + (", ".join(meses_guardados) if meses_guardados else "ninguna"))

    tablas_fi, titulo_fi = None, None
    try:
        tablas_fi, titulo_fi = conexion.tablas_financiamiento(base)
    except Exception as e:
        c2.warning(f"No se pudo leer la Bitácora de F&I: {e}")
    if tablas_fi is not None:
        c2.info(f"Bitácora de F&I: **{titulo_fi}**", icon=":material/link:")
    else:
        f_fi = c2.file_uploader("Bitácora de F&I (opcional)", type=["xlsx"], key="up_fi_base",
                                help="Todas las financieras. Con FINANCIAMIENTO_SHEET_URL en los secrets se lee sola.")
        if f_fi is not None:
            tablas_fi, titulo_fi = fuentes.tablas_de_excel(f_fi.getvalue()), f_fi.name
    fi_guardados, fi_fecha = conexion.registros_fi_guardados(base)
    if tablas_fi is None:
        c2.caption("Guardada: " + (
            f"{len(fi_guardados)} registros" + (f" ({fi_fecha:%d/%m/%Y})" if hasattr(fi_fecha, "strftime") else "")
            if fi_guardados else "ninguna") + ". Lo nuevo se suma a lo guardado.")

    puede = bool(archivos_cet) or not guardadas.empty or tablas_fi is not None or bool(fi_guardados)
    if st.button("Revisar cambios", type="primary", disabled=not puede, key="btn_revisar"):
        with st.spinner("Comparando con el Sheet…"):
            try:
                nuevas = fuentes.leer_capturas([(f.name, f.getvalue()) for f in archivos_cet or []])
                capturas = fuentes.combinar_capturas(guardadas, nuevas) if not nuevas.empty else guardadas
                regs = fuentes.registros_capturas(capturas)
                hojas_fi, regs_fi = [], []
                if tablas_fi is not None:
                    regs_fi, hojas_fi = fuentes.registros_financiamiento(tablas_fi)
                    regs_fi = fuentes.combinar_fi(fi_guardados, regs_fi)      # se acumula con lo ya guardado
                elif fi_guardados:              # la copia guardada: así la Bitacora no pierde lo de F&I
                    regs_fi, hojas_fi = fi_guardados, ["copia guardada en el Sheet"]
                regs += regs_fi
                sol = fuentes.consolidar(regs)
                prop, ctx = actualizar.proponer(base, sol)
                st.session_state["act"] = dict(prop=prop, ctx=ctx, sol=sol, capturas=capturas,
                                               meses=list(dict.fromkeys(capturas["Hoja"])) if not capturas.empty else [],
                                               hojas_fi=hojas_fi, cambia_capturas=not nuevas.empty,
                                               regs_fi=regs_fi if tablas_fi is not None else None)
            except Exception as e:
                st.session_state.pop("act", None)
                st.error(f"No se pudo revisar: {e}")

    act = st.session_state.get("act")
    if act:
        prop = act["prop"]
        st.caption("Cetelem: " + (", ".join(act["meses"]) or "—") + "  ·  F&I: " + (", ".join(act["hojas_fi"]) or "—"))
        m = st.columns(5)
        m[0].metric("Rechazos nuevos", int((prop["Tipo"] == "Nuevo rechazo").sum()))
        m[1].metric("Correcciones", int((prop["Tipo"] == "Corrección").sum()))
        m[2].metric("Datos faltantes", int((prop["Tipo"] == "Dato faltante").sum()),
                    help="Celdas vacías que se llenan con la bitácora (cuadre por teléfono o nombre).")
        m[3].metric("Comentarios", int((prop["Tipo"] == "Comentario").sum()))
        m[4].metric("Solicitudes", len(act["sol"]))
        if prop.empty:
            st.success("La base ya está al día.")
        else:
            st.markdown("**Desmarca** lo que no quieras aplicar; no se volverá a proponer.")
            vista = st.data_editor(
                prop[["Aplicar", "Tipo", "Hoja", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo"]].assign(
                    Hoja=prop["Hoja"].map(lambda h: "Finalizados" if h == TAB_FINALIZADOS else "Rechazos"),
                    Fila=prop["Fila"].map(lambda f: "nueva" if f is None or f != f else str(int(f)))),
                hide_index=True, width="stretch", height=min(520, 60 + 35 * len(prop)), key="editor_prop",
                disabled=["Tipo", "Hoja", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo"],
                column_config={"Aplicar": st.column_config.CheckboxColumn(width="small"),
                               "Fila": st.column_config.TextColumn(width="small"),
                               "Motivo": st.column_config.TextColumn(width="large")})
            prop = prop.assign(Aplicar=vista["Aplicar"].values)
        n = int(prop["Aplicar"].sum()) if not prop.empty else 0
        etiqueta = f"Aplicar {n} cambio(s)" if n else "Guardar bitácoras"
        if st.button(etiqueta, type="primary", key="btn_aplicar"):
            with st.spinner("Escribiendo en el Sheet…"):
                try:
                    r = actualizar.aplicar(base, prop, act["sol"], act["ctx"], act["capturas"], act.get("regs_fi"))
                    st.session_state.pop("act", None)
                    if r.get("aviso"):
                        st.session_state["aviso_act"] = r["aviso"]
                    st.session_state["msg_act"] = (
                        f"Listo: {r['nuevos']} nuevo(s) · {r['correcciones']} corrección(es) · "
                        f"{r.get('faltantes', 0)} dato(s) llenado(s) · {r['ignorados']} ignorado(s)."
                        + (f" Respaldo: «{r['respaldo']}»." if r["respaldo"] else "") + aviso_excel())
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo aplicar: {e}")

with tab_pend:
    if st.session_state.get("msg_pend"):
        st.success(st.session_state.pop("msg_pend"))
    st.markdown("Los pendientes de **Revision** que ya tienen el dato (o que marcaste a mano) se dan por resueltos, "
                "se les quita el morado y la nota, y pasan a «Revision resueltos».")
    if st.button("Revisar pendientes", type="primary", key="btn_rev_pend"):
        with st.spinner("Revisando…"):
            try:
                st.session_state["pend"] = pendientes.revisar(base)
            except Exception as e:
                st.session_state.pop("pend", None)
                st.error(f"No se pudieron revisar: {e}")
    pend = st.session_state.get("pend")
    if pend is not None:
        if pend.empty:
            st.success("No hay pendientes.")
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
            if st.button(f"Limpiar {n} resuelto(s)", type="primary", disabled=n == 0, key="btn_limpiar"):
                with st.spinner("Limpiando…"):
                    try:
                        hechos = pendientes.aplicar(base, pend)
                        st.session_state.pop("pend", None)
                        st.session_state["msg_pend"] = (
                            f"Listo: {hechos} pendiente(s) limpiado(s); los resueltos pasaron a «Revision resueltos»."
                            + aviso_excel())
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo limpiar: {e}")

with tab_hist:
    if st.session_state.get("msg_rest"):
        st.success(st.session_state.pop("msg_rest"))
    try:
        bajas = saleu_sheet.bajas_por_error(base)
    except Exception:
        bajas = None
    if bajas is not None and not bajas.empty:
        with st.container(border=True):
            st.warning(f"**{len(bajas)} rechazo(s) se bajaron por error a «No Registrado»** "
                       "(versión anterior de la app).",
                       icon=":material/undo:")
            st.dataframe(bajas[["Fila", "Cliente", "Antes", "Fecha"]].rename(columns={"Fecha": "Cuándo se bajó"}),
                         hide_index=True, width="stretch", height=min(380, 60 + 35 * len(bajas)))
            if st.button(f"Restaurar {len(bajas)}", type="primary", key="btn_restaurar_su"):
                with st.spinner("Escribiendo en el Sheet…"):
                    try:
                        n = saleu_sheet.restaurar(base, bajas)
                        st.session_state["msg_rest"] = f"Listo: {n} restaurado(s) a Registrado."
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo restaurar: {e}")
    if st.session_state.get("msg_desh"):
        st.success(st.session_state.pop("msg_desh"))
    try:
        dudosos = saleu_sheet.por_telefono_dudosos(base)
    except Exception:
        dudosos = None
    if dudosos is not None and not dudosos.empty:
        with st.container(border=True):
            st.warning(f"**{len(dudosos)} celda(s) se llenaron solo porque el teléfono coincide con otra persona** "
                       "(p. ej. un familiar). Puedes regresarlas.", icon=":material/undo:")
            st.dataframe(dudosos[["Fila", "Cliente", "Campo", "Valor", "Regresar a", "Sale-U", "Fecha"]].assign(
                             Fila=dudosos["Fila"].astype(str), Valor=dudosos["Valor"].astype(str)),
                         hide_index=True, width="stretch", height=min(380, 60 + 35 * len(dudosos)))
            if st.button(f"Deshacer {len(dudosos)}", type="primary", key="btn_deshacer_su"):
                with st.spinner("Escribiendo en el Sheet…"):
                    try:
                        n = saleu_sheet.deshacer_dudosos(base, dudosos)
                        st.session_state["msg_desh"] = f"Listo: {n} celda(s) regresaron a como estaban."
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo deshacer: {e}")
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
