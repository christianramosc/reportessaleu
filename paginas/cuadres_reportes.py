# -*- coding: utf-8 -*-
"""
Página: Resumen y cuadres

  · Resumen: indicadores del Sheet y PDF de quién falta en Sale-U
  · Sale-U vs bitácora (todas las financieras)
  · BlueService vs Inventario (por VIN)

Se abre desde streamlit_app.py (navegación de la app).
"""

import datetime as dt
import os
import tempfile
import pandas as pd
import streamlit as st

import ui
import conexion

from cuadres import saleu as cu
from cuadres import blueservice as cb
from cuadres import saleu_sheet, saleu_fuente
from cuadres.base import a_excel, norm
from cuadres.salidas import etiqueta, telefono
from cuadres import resumen

# =============================================================================
# Utilidades
# =============================================================================
def guardar_temporal(archivos, carpeta):
    """Guarda los archivos subidos con su nombre original (el mes se lee del nombre)."""
    rutas = []
    for a in archivos:
        ruta = os.path.join(carpeta, a.name)
        with open(ruta, "wb") as f:
            f.write(a.getbuffer())
        rutas.append(ruta)
    return rutas


def leer_bytes(ruta):
    with open(ruta, "rb") as f:
        return f.read()


def colorear(valor):
    verde, amarillo, rojo = "#C6EFCE", "#FFEB9C", "#FFC7CE"
    v = str(valor)
    if v in ("Falta crédito", "Registrar crédito"):
        return f"background-color: {amarillo}"
    if v == "Dar de alta":
        return f"background-color: {rojo}"
    if v in ("OK", "Coincide por folio", "Coincide (teléfono y nombre)", "Cuadra"):
        return f"background-color: {verde}"
    if (v.startswith("NO ") or v == "No cuadra" or "pero NO" in v or v.startswith("Falta")
            or v.startswith("En BlueService, no")):
        return f"background-color: {rojo}"
    if (v == "Revisar" or v.startswith("Coincide por") or v.startswith("Posible")
            or v.startswith("Cuadra –")):
        return f"background-color: {amarillo}"
    return ""


def tabla(df, columnas, columnas_color=(), alto=None):
    cols = [c for c in columnas if c in df.columns]
    if df.empty:
        st.info("Sin registros.")
        return
    vista = df[cols].copy()
    vista.columns = [etiqueta(c) for c in cols]
    estilo = vista.style
    color_cols = [etiqueta(c) for c in columnas_color if c in cols]
    if color_cols:
        estilo = estilo.map(colorear, subset=color_cols)
    extra = {"height": alto} if alto else {}
    st.dataframe(estilo, width="stretch", hide_index=True, **extra)


def filtro_vendedor(df, clave, columna="Vendedor"):
    if df.empty or columna not in df.columns:
        return df
    opciones = sorted(v for v in df[columna].dropna().unique() if str(v).strip())
    elegidos = st.multiselect("Vendedor", opciones, key=clave, placeholder="Todos")
    return df[df[columna].isin(elegidos)] if elegidos else df


def valor_resumen(resumen, texto):
    fila = resumen[resumen["Concepto"].str.strip().str.startswith(texto)]
    return fila["Valor"].iloc[0] if not fila.empty else "—"


def descargas(res, prefijo):
    c1, c2, _ = st.columns([1, 1, 3])
    c1.download_button("Excel", res["xlsx_bytes"], file_name=res["xlsx_nombre"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch", key=f"{prefijo}_x", icon=":material/table_view:")
    c2.download_button("PDF", res["pdf_bytes"], file_name=res["pdf_nombre"],
                       mime="application/pdf", width="stretch", key=f"{prefijo}_p",
                       icon=":material/picture_as_pdf:")


def ejecutar_bs(archivo_bs, archivo_inv, hoja):
    with tempfile.TemporaryDirectory() as tmp:
        r_bs, r_inv = guardar_temporal([archivo_bs, archivo_inv], tmp)
        res = cb.procesar(r_bs, r_inv, carpeta=tmp, hoja=hoja)
        res["xlsx_bytes"], res["pdf_bytes"] = leer_bytes(res["xlsx"]), leer_bytes(res["pdf"])
        res["xlsx_nombre"], res["pdf_nombre"] = os.path.basename(res["xlsx"]), os.path.basename(res["pdf"])
    return res


def ejecutar(modulo, archivo_base, capturas):
    """Corre procesar() en una carpeta temporal y regresa tablas + bytes de los archivos."""
    with tempfile.TemporaryDirectory() as tmp:
        ruta_base = guardar_temporal([archivo_base], tmp)[0]
        rutas_cet = guardar_temporal(capturas, tmp)
        res = modulo.procesar(ruta_base, rutas_cet, carpeta=tmp)
        res["xlsx_bytes"], res["pdf_bytes"] = leer_bytes(res["xlsx"]), leer_bytes(res["pdf"])
        res["xlsx_nombre"], res["pdf_nombre"] = os.path.basename(res["xlsx"]), os.path.basename(res["pdf"])
    return res


def revisar_reportes(ruta_su, ruta_fuente):
    """Error claro si los reportes vienen cambiados de lugar."""
    if saleu_fuente.es_fuente(ruta_su):
        raise ValueError("En «Reporte de Crédito» subiste el reporte Fuente. Ahí va el reporteLeadsconCredito; "
                         "el de Fuente va en el segundo campo.")
    if ruta_fuente and not saleu_fuente.es_fuente(ruta_fuente):
        raise ValueError("El archivo de «Reporte Fuente» no es el reporte Fuente de Sale-U "
                         "(ReporteFuentes_Agencia_MG-Colima…).")


def ejecutar_su_base(base, archivo_su, clientes, archivo_fuente=None):
    """Cuadre Sale-U contra los clientes del Sheet (MG Colima Rechazos, todas las financieras, y Créditos
    Finalizados) + cambios propuestos para el Sheet (registro en Sale-U y datos faltantes). Con el reporte Fuente
    (opcional) separa a quién falta dar de alta de a quién solo le falta el crédito. No escribe: los cambios se
    aplican con el botón «Aplicar» después de revisarlos."""
    with tempfile.TemporaryDirectory() as tmp:
        ruta_su = guardar_temporal([archivo_su], tmp)[0]
        ruta_fu = guardar_temporal([archivo_fuente], tmp)[0] if archivo_fuente else None
        revisar_reportes(ruta_su, ruta_fu)
        fu_nuevo = saleu_fuente.leer_fuente(ruta_fu) if ruta_fu else None
        # se suma a los leads ya guardados en el Sheet: así un reporte del mes no pierde los leads de meses anteriores
        fu = saleu_fuente.combinar(saleu_fuente.guardada(base), fu_nuevo)
        guardados = saleu_fuente.guardar(base, fu) if fu_nuevo is not None else 0
        res = {}
        if not clientes.empty:
            res = cu.procesar(ruta_su, carpeta=tmp, bitacora=clientes, fuente="la bitácora", ruta_fuente=fu)
            res["xlsx_bytes"], res["pdf_bytes"] = leer_bytes(res["xlsx"]), leer_bytes(res["pdf"])
            res["xlsx_nombre"], res["pdf_nombre"] = os.path.basename(res["xlsx"]), os.path.basename(res["pdf"])
        su = cu.leer_saleu(ruta_su)
    cambios, (ini, fin) = saleu_sheet.proponer(base, su, fuente=fu)
    periodo = f"{ini:%d/%m/%Y} – {fin:%d/%m/%Y}"
    res["registro"] = cambios
    res["periodo_saleu"] = periodo
    res["con_fuente"] = fu is not None
    res["fuente_guardados"], res["fuente_subida"] = guardados, fu_nuevo is not None
    res["fuente_df"], res["credito_df"] = fu, su       # para «Cómo está en Sale-U» en la lista del Resumen
    res["periodo_fuente"] = saleu_fuente.periodo_txt(fu)
    res["datos"] = pd.DataFrame(saleu_sheet.proponer_datos(base, su),
                                columns=["Hoja", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo", "_valor"])
    res["escritos"], res["datos_escritos"], res["aplicado"] = 0, 0, False
    return res


# =============================================================================
# Barra lateral: ajustes
# =============================================================================
with st.sidebar.expander("Ajustes avanzados", icon=":material/tune:"):
    st.caption("Cambian los resultados de los cuadres. Déjalos como están salvo que sepas por qué.")
    fuerte = st.slider("Parecido: solo por nombre", 0.6, 1.0, 0.80, 0.05, help="Cuando el teléfono es distinto.")
    con_tel = st.slider("Parecido: con teléfono igual", 0.3, 1.0, 0.50, 0.05, help="Cuando el teléfono coincide.")
    cu.UMBRAL_NOMBRE_FUERTE, cu.UMBRAL_NOMBRE_CON_TEL = fuerte, con_tel
    contra = st.selectbox("Contrapropuesta en Sale-U", ["Revisar", "Aceptado", "Rechazado"])
    cu.ESTATUS_ESPERADO["CONTRAPROPUESTA"] = {"Aceptado": "ACEPTADO",
                                              "Rechazado": "RECHAZADO"}.get(contra)


# =============================================================================
# Página
# =============================================================================
ui.hero("Resumen y cuadres", "Indicadores del Sheet, cuadres y sus reportes en Excel y PDF.")
ui.guia_mes()

try:
    base = conexion.obtener_base()      # una sola lectura del Sheet por ejecución (cuenta en la cuota de Google)
except Exception as e:
    base = None
    st.error(f"No se pudo abrir el Google Sheet: {e}. Mientras, puedes subir los archivos.",
             icon=":material/cloud_off:")

tab_res, tab_su, tab_bs = st.tabs(["📊 Resumen", "🟦 Sale-U", "🔧 BlueService"])

HOJA_CORTA = lambda h: "Bitacora" if h == "Bitacora" else "Rechazos"
AGREGAR_MES = "Para agregar un mes: **Base de seguimiento → Actualizar**."
AYUDA_CAPTURAS = "Las pide el F&I. Un Excel por mes (ej. JULIO_2026_FINAL.xlsx); puedes subir varios."

# ---------------------------------------------------------------- RESUMEN
ESTADO = dict(bien="#0ca30c", aviso="#c98500", critico="#d03b3b", neutro=ui.kc.AZUL_OSCURO)
with tab_res:
    if base is None:
        st.info("Conecta el Google Sheet (o sube el Excel en **Base de seguimiento**) para ver el resumen.",
                icon=":material/cloud_off:")
    else:
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
            res_su = st.session_state.get("res_su") or {}
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
                st.caption("Para llenar «Cómo está en Sale-U» y «Fuente», haz una vez el cuadre de Sale-U con el reporte "
                           "Fuente: sus leads se guardan en el Sheet (pestaña oculta «Fuente Sale-U»).")
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

# ---------------------------------------------------------------- SALE-U
with tab_su:
    c1, c2 = st.columns(2)
    f_su = c1.file_uploader("Reporte de Crédito (Sale-U)", type=["csv", "xlsx"], key="up_su",
                            help="reporteLeadsconCredito…, tal cual se descarga (sin abrirlo en Excel).")
    f_fu = c1.file_uploader("Reporte Fuente (opcional)", type=["csv", "xlsx"], key="up_fuente",
                            help="ReporteFuentes_Agencia_MG-Colima…: todos los leads. Con él se separa a quién falta "
                                 "dar de alta de a quién solo le falta el crédito. Sus leads se guardan en el Sheet y "
                                 "se suman a los anteriores, así basta subir el del mes.")
    if base is not None:
        bit = conexion.clientes_saleu(base)
        if not bit.empty:
            c2.success(f"Contra la bitácora completa: {conexion.describir_creditos(bit)}",
                       icon=":material/cloud_done:")
            c2.caption("Todas las financieras y estatus del periodo del reporte, más los Finalizados (por fecha de entrega; "
                       "los que no la tienen entran todos). " + AGREGAR_MES)
        else:
            c2.warning("No hay créditos en la bitácora ni en Rechazos / Finalizados.",
                       icon=":material/warning:")
        if st.button("Hacer cuadre", type="primary", disabled=not f_su, key="btn_su"):
            with st.spinner("Cruzando…"):
                try:
                    st.session_state["res_su"] = ejecutar_su_base(base, f_su, bit, f_fu)
                except Exception as e:
                    st.session_state.pop("res_su", None)
                    st.error(f"No se pudo hacer el cuadre: {e}")
    else:
        f_cet2 = c2.file_uploader("Capturas de Cetelem", type=["xlsx"], accept_multiple_files=True,
                                  key="up_cet_su", help=AYUDA_CAPTURAS)
        c2.caption("Sin el Sheet conectado solo se cuadra Cetelem.")

        if st.button("Hacer cuadre", type="primary", disabled=not (f_su and f_cet2), key="btn_su"):
            with st.spinner("Cruzando…"):
                try:
                    st.session_state["res_su"] = ejecutar(cu, f_su, f_cet2)
                except Exception as e:
                    st.session_state.pop("res_su", None)
                    st.error(f"No se pudo hacer el cuadre: {e}")

    res = st.session_state.get("res_su")
    if res and "registro" in res:
        reg = res["registro"]
        cambian = reg[reg["Cambia"]]
        if res.get("fuente_subida"):
            st.caption(f"Reporte Fuente guardado en el Sheet (pestaña oculta «Fuente Sale-U»): "
                       f"{res.get('fuente_guardados', 0)} leads acumulados.")
        st.markdown(f"**Registro en Sale-U** · Crédito {res['periodo_saleu']}"
                    + (f" · Fuente {res['periodo_fuente']}" if res.get("con_fuente") else " · sin reporte Fuente"))
        datos = res.get("datos", pd.DataFrame())
        aplicado = res.get("aplicado", False)
        escritos_datos = res.get("datos_escritos", 0)
        k = st.columns(6)
        k[0].metric("Revisados", len(reg), help="Rechazos de MG Colima Rechazos y Aprobados y Rechazados de la Bitacora del periodo del reporte de Crédito "
                                           "(más los «No Registrado» de otras fechas que salen en Fuente).")
        k[1].metric("Registrados", int((reg["Después"] == saleu_sheet.REGISTRADO).sum()))
        k[2].metric("Falta crédito", int((reg["Después"] == saleu_sheet.FALTA_CREDITO).sum()),
                    help="Están en el reporte Fuente, pero no en el de Crédito." if res.get("con_fuente")
                    else "Sube el reporte Fuente para separarlos de los «No Registrado».")
        k[3].metric("No Registrado", int((reg["Después"] == saleu_sheet.NO_REGISTRADO).sum()),
                    help="No están en Sale-U: falta darlos de alta.")
        k[4].metric("Celdas actualizadas" if aplicado else "Celdas por actualizar",
                    res["escritos"] if aplicado else len(cambian),
                    help="Registrado / Falta crédito / No Registrado. Nunca baja de nivel lo que ya dice el Sheet.")
        k[5].metric("Datos llenados" if aplicado else "Datos faltantes", escritos_datos if aplicado else len(datos),
                    help="Celdas vacías que se llenan con Sale-U o la bitácora (si ambas lo traen, manda la bitácora).")
        if aplicado:
            st.success(f"Listo: {res['escritos']} registro(s) en Sale-U y {escritos_datos} dato(s) escritos en el Sheet.")
        elif len(cambian) or len(datos):
            n_cambios = len(cambian) + len(datos)
            st.info("Revisa el detalle de abajo y aplica los cambios en el Sheet.", icon=":material/visibility:")
            if st.button(f"Aplicar {n_cambios} cambio(s) en el Sheet", type="primary", key="btn_aplicar_su"):
                with st.spinner("Escribiendo…"):
                    try:
                        res["escritos"] = saleu_sheet.aplicar(base, reg, res["periodo_saleu"])
                        res["datos_escritos"] = saleu_sheet.aplicar_datos(base, datos.to_dict("records"))
                        res["aplicado"] = True
                        st.rerun()
                    except Exception as e:
                        st.error(f"No se pudo aplicar: {e}")
        else:
            st.success("El Sheet ya está al día con este reporte.")
        if len(datos):
            with st.expander(f"Datos {'llenados' if aplicado else 'por llenar'} ({len(datos)})"):
                tabla(datos.assign(Hoja=datos["Hoja"].map(lambda h: "Finalizados" if "Finalizados" in h else "Rechazos"),
                                   Después=datos["Después"].astype(str)),
                      ["Hoja", "Fila", "Cliente", "Campo", "Después", "Motivo"])
        rev = reg[reg["Revisar"] != ""] if "Revisar" in reg.columns else reg.iloc[0:0]
        if len(rev):
            st.info(f"**{len(rev)} marcado(s) como Registrado no salen en los reportes.** Se dejaron igual; "
                    "confírmalos en Sale-U.", icon=":material/fact_check:")
            tabla(rev.assign(Fecha_ingreso=rev["Fecha_ingreso"].map(lambda d: d.strftime("%d/%m/%Y") if d else ""),
                             Hoja=rev["Hoja"].map(HOJA_CORTA)),
                  ["Hoja", "Fila", "Cliente", "Fecha_ingreso", "Antes", "Después", "Revisar"])
        with st.expander(f"Detalle del registro ({len(cambian)} cambio(s) de {len(reg)})"):
            tabla(reg.assign(Fecha_ingreso=reg["Fecha_ingreso"].map(lambda d: d.strftime("%d/%m/%Y") if d else ""),
                             Hoja=reg["Hoja"].map(HOJA_CORTA)),
                  ["Hoja", "Fila", "Cliente", "Fecha_ingreso", "Antes", "Después", "Como", "Cliente_SaleU", "Estatus_SaleU",
                   "Financiera_SaleU"] + (["Cliente_Fuente", "Alta_Fuente"] if res.get("con_fuente") else []))
        if isinstance(base, conexion.BaseExcel):
            st.info("Modo Excel: descarga el archivo en **Base de seguimiento**.")
    if res and "resumen" in res:
        r = res["resumen"]
        fuente = res.get("fuente", "la bitácora")
        st.success(f"{fuente[0].upper()}{fuente[1:]}: {', '.join(res['meses'])}  ·  "
                   f"Financieras: {', '.join(res.get('financieras', []))}")
        m = st.columns(5)
        m[0].metric("En bitácora", valor_resumen(r, "Clientes únicos"))
        m[1].metric("En Sale-U", valor_resumen(r, f"Clientes {cu.de(fuente)} encontrados"))
        m[2].metric("Faltan", valor_resumen(r, f"Clientes {cu.de(fuente)} NO"))
        m[3].metric("% cuadre", valor_resumen(r, "Porcentaje"))
        m[4].metric("Estatus distinto", valor_resumen(r, "Estatus que no cuadra"))
        descargas(res, "su")

        t1, t2, t4, t5 = st.tabs([
            f"Faltan en Sale-U ({len(res['faltantes'])})",
            f"A revisar ({len(res['revisar'])})",
            "Completo", "Resumen"])
        with t1:
            if "registro" in res:
                est = res["faltantes"]["Estatus_final"]
                n_hoja = int((res["registro"]["Después"] == saleu_sheet.NO_REGISTRADO).sum())
                n_cred = int((res["registro"]["Después"] == saleu_sheet.FALTA_CREDITO).sum())
                por_est = " · ".join(f"{k.capitalize()} {v}" for k, v in est.value_counts().items())
                por_falta = (" · " + " · ".join(f"{k} {v}" for k, v in res["faltantes"]["Falta"].value_counts().items())
                             if "Falta" in res["faltantes"] else "")
                st.caption(f"Por estatus: {por_est}{por_falta} · En el Sheet: **{n_hoja}** *No Registrado* y "
                           f"**{n_cred}** *Falta crédito*.")
            df = filtro_vendedor(res["faltantes"], "fv_su")
            vista = cu._faltantes_con_vendedor(df).assign(Historial=df["Historial"].values,
                                                          Observaciones=df["Observaciones"].values)
            vista["Fecha"] = vista["Fecha"].map(cu.fecha_txt)
            vista["Mensualidad"] = vista["Mensualidad"].map(lambda v: f"${v:,.2f}" if v == v and v else "")
            vista["Plazo"] = vista["Plazo"].map(lambda v: str(int(v)) if v == v and v else "")
            con_f = "Falta" in df
            tabla(vista, (["Falta"] if con_f else []) + cu.COLS_FALTANTES[:2] + ["Vendedor"] + cu.COLS_FALTANTES[2:]
                  + [c for c in (cu.COL_SALEU, "Fuente") if c in vista] + ["Historial", "Observaciones"],
                  ["Falta"] if con_f else ())
        with t2:
            tabla(res["revisar"], ["Resultado", "Cuadre_estatus", "Cliente_bitacora", "Tel_bitacora", cu.COL_SALEU,
                                   "Cliente_SaleU", "Tel_SaleU", "Estatus_SaleU", "Financiera_SaleU", "Fuente_SaleU",
                                   "Observaciones"], ["Resultado", "Cuadre_estatus"])
        with t4:
            df = filtro_vendedor(res["cuadre"], "fv_su_todo")
            tabla(df, ["Resultado", "Cuadre_estatus", "Cliente_bitacora", "Tel_bitacora", "Vendedor",
                       "Financiera", "Historial", "Cliente_SaleU", "Tel_SaleU", "Asesor_SaleU", "Estatus_SaleU",
                       "Financiera_SaleU", "Fuente_SaleU", "Observaciones"], ["Resultado", "Cuadre_estatus"], alto=550)
        with t5:
            st.dataframe(r.astype(str), hide_index=True, width="stretch")

# ---------------------------------------------------------------- BLUESERVICE
with tab_bs:
    c1, c2 = st.columns(2)
    f_pdi = c1.file_uploader("PDIs de BlueService", type=["xlsx"], key="up_pdi",
                             help="Ej. PDIs_MG_Colima.xlsx")
    f_inv = c2.file_uploader("Inventario", type=["xlsx"], key="up_inv",
                             help="Ej. INVENTARIO_MG_COLIMA….xlsx")
    hoja = None
    if f_inv:
        try:
            hojas = pd.ExcelFile(f_inv).sheet_names
            hoja = c2.selectbox("Hoja a cuadrar", hojas, index=hojas.index(cb.hoja_por_defecto(hojas)),
                                help="Las demás hojas dicen dónde está cada VIN (vendidos, traslados, servicio).")
        except Exception as e:
            c2.error(f"No se pudieron leer las hojas: {e}")

    if st.button("Hacer cuadre", type="primary", disabled=not (f_pdi and f_inv and hoja), key="btn_bs"):
        with st.spinner("Cruzando VINs…"):
            try:
                st.session_state["res_bs"] = ejecutar_bs(f_pdi, f_inv, hoja)
            except Exception as e:
                st.session_state.pop("res_bs", None)
                st.error(f"No se pudo hacer el cuadre: {e}")

    res = st.session_state.get("res_bs")
    if res:
        r, h = res["resumen"], res["hoja"]
        st.success(f"Hoja: {h}  ·  Corte: {cb.fmt(res['corte'])}")
        m = st.columns(5)
        m[0].metric(f"En {h.title()}", valor_resumen(r, "Unidades en la hoja"))
        m[1].metric("Cuadran", valor_resumen(r, f"Unidades de {h} que cuadran"))
        m[2].metric("Faltan en BS", valor_resumen(r, f"Unidades de {h} que faltan"))
        m[3].metric("Solo en BS", valor_resumen(r, "En BlueService sin venta"),
                    help="En BlueService sin venta, pero no en el inventario.")
        m[4].metric("A revisar", valor_resumen(r, "Unidades con datos"))
        descargas(res, "bs")

        cols_inv = ["Resultado", "VIN", "Modelo_inventario", "Version", "Color", "Ubicacion",
                    "Entrada_inventario", "Ultimo_PDI_inventario", "Proximo_PDI_inventario",
                    "Etapa_BS", "Status_BS", "Observaciones"]
        t1, t2, t3, t4, t5, t6 = st.tabs([
            f"Faltan en BS ({len(res['falta_bs'])})",
            f"Solo en BS ({len(res['falta_inv'])})",
            f"A revisar ({len(res['revisar'])})",
            f"{h.title()} completo", f"Vendidas en BS ({len(res['vendidas'])})", "Resumen"])
        with t1:
            st.caption(f"Unidades en {h.lower()} que no salen en BlueService.")
            tabla(res["falta_bs"], cols_inv, ["Resultado"])
        with t2:
            df = res["falta_inv"]
            if not df.empty:
                sit = st.multiselect("Situación", sorted(df["Situacion"].unique()), key="f_sit",
                                     placeholder="Todas")
                if sit:
                    df = df[df["Situacion"].isin(sit)]
                st.bar_chart(res["falta_inv"].groupby("Situacion").size().rename("Unidades"),
                             color="#C00000", height=200, horizontal=True)
            tabla(df, ["Situacion", "VIN", "Modelo_BS", "Dealer_origen", "Dealer_actual",
                       "Fecha_inventario_BS", "Etapa_BS", "Status_BS", "Aparece_en_inventario",
                       "Observaciones"], ["Situacion"])
        with t3:
            st.caption("VIN mal capturados, PDI inválidos o distintos y vendidas que siguen en inventario.")
            tabla(res["revisar"], ["Resultado", "VIN", "VIN_BlueService"] + cols_inv[2:], ["Resultado"])
        with t4:
            tabla(res["cuadre_inv"], ["Resultado", "VIN", "VIN_BlueService"] + cols_inv[2:] + ["Fila_excel"],
                  ["Resultado"], alto=550)
        with t5:
            st.caption("Unidades con fecha de venta en BlueService (informativo).")
            tabla(res["vendidas"], ["Situacion", "VIN", "Modelo_BS", "Fecha_venta_BS", "Etapa_BS",
                                    "Status_BS", "Aparece_en_inventario"])
        with t6:
            st.dataframe(r.astype(str), hide_index=True, width="stretch")
