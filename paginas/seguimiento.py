# -*- coding: utf-8 -*-
"""
Página: Seguimiento

  · Resumen: indicadores del año, gráficas, seguimientos vencidos (con PDF por asesor), pendientes de Revision y
    quién falta en Sale-U (PDF y Excel).
  · Pendientes: revisa los pendientes de la pestaña Revision y limpia los resueltos.
"""
import datetime as dt

import pandas as pd
import streamlit as st

import ui
import conexion
from comunes import ESTADO, aviso_excel
from cuadres import pendientes, resumen, saleu_fuente
from cuadres import saleu as cu
from cuadres.salidas import telefono

ui.hero("Seguimiento", "Indicadores del Sheet, vencidos, pendientes y quién falta en Sale-U.")
ui.guia_mes()

base = conexion.panel_base()
if base is None:
    st.stop()

tab_res, tab_pend = st.tabs(["📊 Resumen", "✅ Pendientes"])

with tab_res:
    if True:
        try:
            c_anio, _ = st.columns([1, 5])
            anio = c_anio.selectbox("Año", resumen.anios(base), key="anio_res")
            k = resumen.indicadores(base, anio)
        except Exception as e:
            k = None
            st.error(f"No se pudo leer el Sheet: {e}")
        if k:
            pct = lambda a, b: f"{a / b:.0%}" if b else "—"
            al_dia = k["al_dia"] / k["ya_toca"] if k["ya_toca"] else 1
            fila1, fila2 = st.columns(3), st.columns(3)
            ui.tarjeta(fila1[0], "event_busy", "Rechazos en el año", k["rechazos"],
                       " · ".join(f"{f} {n}" for f, n in sorted(k["por_fin"].items(), key=lambda x: -x[1])))
            ui.tarjeta(fila1[1], "event_available", "Seguimiento al día", pct(k["al_dia"], k["ya_toca"]),
                       f"{k['al_dia']} de {k['ya_toca']} que ya cumplieron 180 días",
                       ESTADO["bien"] if al_dia >= 0.9 else ESTADO["aviso"])
            ui.tarjeta(fila1[2], "call", "Contacto efectivo", pct(k["contestaron"], k["contactados"]),
                       f"{k['contestaron']} de {k['contactados']} contactados"
                       + (f" · {k['sin_contesta']} sin dato" if k["sin_contesta"] else ""))
            ui.tarjeta(fila2[0], "how_to_reg", "Registrados en Sale-U", pct(k["registrados"], k["saleu_total"]),
                       f"{k['falta_alta']} falta alta · {k['falta_credito']} falta crédito"
                       + (f" · {k['aprob_sin_dato']} por confirmar" if k["aprob_sin_dato"] else ""))
            ui.tarjeta(fila2[1], "alarm", "Seguimientos vencidos", len(k["vencidos"]),
                       f"{len(k['por_vencer'])} vencen en 15 días",
                       ESTADO["critico"] if len(k["vencidos"]) else ESTADO["bien"])
            ui.tarjeta(fila2[2], "rule", "Pendientes en Revision", k["pendientes"], "por resolver",
                       ESTADO["aviso"] if k["pendientes"] else ESTADO["bien"])

            import altair as alt
            g1, g2 = st.columns([3, 2])
            if not k["por_mes"].empty:
                meses = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]
                pm = k["por_mes"].assign(Mes=k["por_mes"]["mes"].map(lambda m: meses[m - 1]))
                orden_fin = ["CETELEM", "BBVA", "Otras"]
                g1.markdown("**Rechazos por mes**")
                g1.altair_chart(alt.Chart(pm).mark_bar(cornerRadiusTopLeft=3, cornerRadiusTopRight=3, size=22,
                                                      stroke="white", strokeWidth=2).encode(
                    x=alt.X("Mes:N", sort=meses, title=None, axis=alt.Axis(labelAngle=0)),
                    y=alt.Y("sum(rechazos):Q", title=None, axis=alt.Axis(tickMinStep=1, grid=True)),
                    color=alt.Color("fin:N", title=None, sort=orden_fin, legend=alt.Legend(orient="top"),
                                    scale=alt.Scale(domain=orden_fin, range=["#2a78d6", "#eb6834", "#1baf7a"])),
                    order=alt.Order("orden:Q"),
                    tooltip=[alt.Tooltip("Mes:N"), alt.Tooltip("fin:N", title="Financiera"),
                             alt.Tooltip("rechazos:Q", title="Rechazos")]
                ).transform_calculate(orden="indexof(['CETELEM','BBVA','Otras'], datum.fin)").properties(height=260),
                    width="stretch")
            embudo = pd.DataFrame({"Etapa": ["Ya cumplieron 180 días", "Con 1er seguimiento", "Contestaron"],
                                   "Clientes": [k["ya_toca"], k["al_dia"], k["contestaron"]]})
            g2.markdown("**Seguimiento de los que ya cumplieron 180 días**")
            base_e = alt.Chart(embudo).encode(
                y=alt.Y("Etapa:N", sort=list(embudo["Etapa"]), title=None, axis=alt.Axis(labelLimit=180)),
                x=alt.X("Clientes:Q", title=None, axis=None))
            g2.altair_chart((base_e.mark_bar(cornerRadiusTopRight=4, cornerRadiusBottomRight=4, size=26).encode(
                color=alt.Color("Etapa:N", legend=None, sort=list(embudo["Etapa"]),
                                scale=alt.Scale(range=["#86b6ef", "#2a78d6", "#104281"])),
                tooltip=["Etapa", "Clientes"])
                + base_e.mark_text(align="left", dx=4, fontWeight="bold", color=ui.kc.TEXTO).encode(text="Clientes:Q")
            ).properties(height=260), width="stretch")

            def lista(df, cols, nombres):
                vista = df[cols].copy()
                vista.columns = nombres
                return vista

            v, pv = k["vencidos"], k["por_vencer"]
            with st.expander(f"Seguimientos vencidos ({len(v)})", icon=":material/alarm:", expanded=bool(len(v))):
                if v.empty:
                    st.success("No hay seguimientos vencidos.")
                else:
                    st.dataframe(lista(v.assign(tel=v["tel"].map(telefono), fseg=v["fseg"].map(cu.fecha_txt)),
                                       ["nombre", "tel", "asesor", "fseg", "dias", "id"],
                                       ["Cliente", "Celular", "Asesor", "Tocaba el", "Días de atraso", "ID"]),
                                 hide_index=True, width="stretch")
            if not v.empty:
                from cuadres import reporte_vencidos
                c1, c2, _ = st.columns([2, 1, 2])
                quien = c1.selectbox("PDF de vencidos para", ["Todos los asesores (una página cada uno)"]
                                     + reporte_vencidos.asesores(v), key="asesor_venc")
                elegido = None if quien.startswith("Todos") else quien
                c2.download_button("PDF de vencidos", reporte_vencidos.pdf(v, elegido),
                                   file_name=f"Vencidos_{(elegido or 'todos').replace(' ', '_')}_{dt.date.today():%Y%m%d}.pdf",
                                   mime="application/pdf", icon=":material/picture_as_pdf:", key="dl_venc",
                                   width="stretch")
            with st.expander(f"Vencen en los próximos 15 días ({len(pv)})", icon=":material/schedule:"):
                if pv.empty:
                    st.info("Ninguno vence en los próximos 15 días.")
                else:
                    st.dataframe(lista(pv.assign(tel=pv["tel"].map(telefono), fseg=pv["fseg"].map(cu.fecha_txt),
                                                 dias=-pv["dias"]),
                                       ["nombre", "tel", "asesor", "fseg", "dias", "id"],
                                       ["Cliente", "Celular", "Asesor", "Toca el", "Faltan días", "ID"]),
                                 hide_index=True, width="stretch")
            with st.expander(f"Pendientes en Revision ({k['pendientes']})", icon=":material/rule:"):
                rv = k["revision"]
                if rv.empty:
                    st.success("No hay pendientes.")
                else:
                    st.dataframe(rv[[c for c in ["Hoja", "Celda", "Cliente", "Campo", "Motivo / qué falta"]
                                     if c in rv.columns]], hide_index=True, width="stretch")

        # ------------------------------------------------ reporte de faltantes en Sale-U
        st.divider()
        st.markdown("#### :material/person_add: Faltan en Sale-U")
        st.caption("Aprobados y Rechazados sin registro, según el Sheet (lo marca el cuadre de Sale-U): «No Registrado» "
                   "= falta darlos de alta; «Falta crédito» = ya son lead (reporte Fuente), falta su crédito.")
        try:
            res_su = (st.session_state.get("todo") or {}).get("su") or {}
            fuente_su = res_su.get("fuente_df")
            if fuente_su is None:
                fuente_su = saleu_fuente.guardada(base)       # leads guardados en el Sheet por el último cuadre
            faltan, confirmar = resumen.faltantes_saleu(base, fuente=fuente_su, credito=res_su.get("credito_df"))
        except Exception as e:
            faltan = confirmar = None
            st.error(f"No se pudo armar la lista: {e}")
        if faltan is not None:
            m = st.columns(5)
            m[0].metric("Por registrar", len(faltan),
                        help=f"{int((faltan['Estatus'] == resumen.APROBADO).sum())} aprobados · "
                             f"{int((faltan['Estatus'] == resumen.RECHAZADO).sum())} rechazados")
            m[1].metric("Dar de alta", int((faltan["Falta"] == resumen.FALTA_ALTA).sum()),
                        help="No están en Sale-U (ni Fuente ni Crédito).")
            m[2].metric("Registrar crédito", int((faltan["Falta"] == resumen.FALTA_CREDITO).sum()),
                        help="Ya son lead (reporte Fuente); falta registrar su crédito.")
            m[3].metric("Aprobados", int((faltan["Estatus"] == resumen.APROBADO).sum()))
            m[4].metric("Por confirmar", len(confirmar), help="Aprobados sin dato de Sale-U.")
            if st.button("Generar PDF y Excel", type="primary", icon=":material/picture_as_pdf:", key="btn_pdf_su"):
                with st.spinner("Armando el reporte…"):
                    try:
                        from cuadres import reporte_saleu
                        nombre = f"PorRegistrar_SaleU_{dt.date.today():%Y%m%d}"
                        st.session_state["pdf_su"] = (reporte_saleu.pdf(faltan, confirmar), nombre + ".pdf",
                                                      reporte_saleu.excel(faltan, confirmar), nombre + ".xlsx")
                    except Exception as e:
                        st.error(f"No se pudo generar el reporte: {e}")
            if st.session_state.get("pdf_su") and len(st.session_state["pdf_su"]) == 4:
                datos_pdf, nombre_pdf, datos_xl, nombre_xl = st.session_state["pdf_su"]
                d1, d2 = st.columns(2)
                d1.download_button("Descargar PDF", datos_pdf, file_name=nombre_pdf, mime="application/pdf",
                                   icon=":material/download:", key="dl_pdf_su", width="stretch")
                d2.download_button("Descargar Excel", datos_xl, file_name=nombre_xl, icon=":material/download:",
                                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                   key="dl_xl_su", width="stretch")
            if fuente_su is None:
                st.caption("Para llenar «Cómo está en Sale-U» y «Fuente», sube una vez el reporte Fuente "
                           "en Actualizar: sus leads se guardan en el Sheet (pestaña oculta «Fuente Sale-U»).")
            else:
                st.caption(f"«Cómo está en Sale-U» y «Fuente»: {len(fuente_su)} leads del reporte Fuente "
                           f"({saleu_fuente.periodo_txt(fuente_su)}).")
            if not faltan.empty:
                vista = faltan.assign(Fecha=faltan["Fecha"].map(cu.fecha_txt),
                                      Mensualidad=faltan["Mensualidad"].map(
                                          lambda v: f"${v:,.2f}" if v == v and v else ""))
                st.dataframe(vista[["Falta", "Estatus", "Cliente", resumen.COL_SALEU, "Fuente", "Telefono", "Vendedor", "Financiera", "Fecha", "Dias",
                                    "Mensualidad"]].rename(columns={"Telefono": "Celular", "Dias": "Días"}),
                             hide_index=True, width="stretch")

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
                            + aviso_excel(base))
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo limpiar: {e}")
