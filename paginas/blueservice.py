# -*- coding: utf-8 -*-
"""
Página: BlueService

Cuadre por VIN de los PDIs de BlueService contra el inventario de MG Colima (Excel y PDF).
"""
import pandas as pd
import streamlit as st

import ui
from comunes import descargas, tabla, valor_resumen
from cuadres import blueservice as cb

ui.encabezado("BlueService", "Cuadre por VIN de los PDIs de BlueService contra el inventario.")


def ejecutar_bs(archivo_bs, archivo_inv, hoja):
    return cb.procesar(archivo_bs.getvalue(), archivo_inv.getvalue(), hoja=hoja)


ui.seccion("Archivos", "El reporte de PDIs de BlueService y el inventario de MG Colima.")
caja = st.container(border=True)
c1, c2 = caja.columns(2)
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
        ui.error("No se pudieron leer las hojas del inventario", e, donde=c2)

if ui.barra(caja, alinear="left").button("Hacer cuadre", type="primary", disabled=not (f_pdi and f_inv and hoja),
                                         key="btn_bs", icon=":material/compare_arrows:"):
    with st.spinner("Cruzando VINs…"):
        try:
            st.session_state["res_bs"] = ejecutar_bs(f_pdi, f_inv, hoja)
        except Exception as e:
            st.session_state.pop("res_bs", None)
            ui.error("No se pudo hacer el cuadre", e)

res = st.session_state.get("res_bs")
if res:
    r, h = res["resumen"], res["hoja"]
    t, d = st.columns([3, 2], vertical_alignment="bottom")
    ui.seccion("Resultado", f"Hoja: {h}  ·  Corte: {cb.fmt(res['corte'])}", donde=t)
    descargas(res, "bs", donde=d)
    m = st.columns(5)
    m[0].metric(f"En {h.title()}", valor_resumen(r, "Unidades en la hoja"))
    m[1].metric("Cuadran", valor_resumen(r, f"Unidades de {h} que cuadran"))
    m[2].metric("Faltan en BS", valor_resumen(r, f"Unidades de {h} que faltan"))
    m[3].metric("Solo en BS", valor_resumen(r, "En BlueService sin venta"),
                help="En BlueService sin venta, pero no en el inventario.")
    m[4].metric("A revisar", valor_resumen(r, "Unidades con datos"))
    st.write("")

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
                         color=ui.ESTADO["critico"], height=200, horizontal=True)
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
