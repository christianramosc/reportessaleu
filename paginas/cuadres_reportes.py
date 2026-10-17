# -*- coding: utf-8 -*-
"""
Página: Cuadres para reportes

  · Seguimiento de rechazos vs bitácora (solo rechazos, todas las financieras)
  · Sale-U vs bitácora (todas las financieras)
  · BlueService vs Inventario (por VIN)

Se abre desde streamlit_app.py (navegación de la app).
"""

import os
import tempfile
import pandas as pd
import streamlit as st

import ui
import conexion

from cuadres import seguimiento as cs
from cuadres import saleu as cu
from cuadres import blueservice as cb
from cuadres import saleu_sheet
from cuadres.base import a_excel, norm
from cuadres.salidas import etiqueta

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


def ejecutar_seg_base(base, bitacora):
    """Cuadre de seguimiento: pestañas de Colima de la base vs rechazos de la bitácora (todas las financieras)."""
    with tempfile.TemporaryDirectory() as tmp:
        hojas = [h for h in base.pestanas() if cs.FILTRO_HOJAS in norm(h) and not h.startswith("_")]
        ruta_seg = os.path.join(tmp, "Seguimiento_base.xlsx")
        with open(ruta_seg, "wb") as f:
            f.write(a_excel(base, hojas))
        res = cs.procesar(ruta_seg, carpeta=tmp, bitacora=bitacora)
        res["xlsx_bytes"], res["pdf_bytes"] = leer_bytes(res["xlsx"]), leer_bytes(res["pdf"])
        res["xlsx_nombre"], res["pdf_nombre"] = os.path.basename(res["xlsx"]), os.path.basename(res["pdf"])
    return res


def ejecutar_su_base(base, archivo_su, clientes):
    """Cuadre Sale-U contra los clientes del Sheet (MG Colima Rechazos, todas las financieras, y Créditos
    Finalizados) + cambios propuestos para el Sheet (registro en Sale-U y datos faltantes). No escribe:
    los cambios se aplican con el botón «Aplicar» después de revisarlos."""
    with tempfile.TemporaryDirectory() as tmp:
        ruta_su = guardar_temporal([archivo_su], tmp)[0]
        res = {}
        if not clientes.empty:
            res = cu.procesar(ruta_su, carpeta=tmp, bitacora=clientes, fuente="la bitácora")
            res["xlsx_bytes"], res["pdf_bytes"] = leer_bytes(res["xlsx"]), leer_bytes(res["pdf"])
            res["xlsx_nombre"], res["pdf_nombre"] = os.path.basename(res["xlsx"]), os.path.basename(res["pdf"])
        su = cu.leer_saleu(ruta_su)
    cambios, (ini, fin) = saleu_sheet.proponer(base, su)
    periodo = f"{ini:%d/%m/%Y} – {fin:%d/%m/%Y}"
    res["registro"] = cambios
    res["periodo_saleu"] = periodo
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
    for m in (cs, cu):
        m.UMBRAL_NOMBRE_FUERTE, m.UMBRAL_NOMBRE_CON_TEL = fuerte, con_tel
    contra = st.selectbox("Contrapropuesta en Sale-U", ["Revisar", "Aceptado", "Rechazado"])
    cu.ESTATUS_ESPERADO["CONTRAPROPUESTA"] = {"Aceptado": "ACEPTADO",
                                              "Rechazado": "RECHAZADO"}.get(contra)


# =============================================================================
# Página
# =============================================================================
ui.hero("Cuadres para reportes", "Elige el cuadre y descarga el resultado en Excel y PDF.")
ui.guia_mes()

try:
    base = conexion.obtener_base()      # una sola lectura del Sheet por ejecución (cuenta en la cuota de Google)
except Exception as e:
    base = None
    st.error(f"No se pudo abrir el Google Sheet: {e}. Mientras, puedes subir los archivos.",
             icon=":material/cloud_off:")

tab_seg, tab_su, tab_bs = st.tabs(["📒 Seguimiento", "🟦 Sale-U", "🔧 BlueService"])

AGREGAR_MES = "Para agregar un mes: **Base de seguimiento → Actualizar**."
AYUDA_CAPTURAS = "Las pide el F&I. Un Excel por mes (ej. JULIO_2026_FINAL.xlsx); puedes subir varios."

# ---------------------------------------------------------------- SEGUIMIENTO
with tab_seg:
    st.markdown("Busca en el seguimiento los **rechazos** de la bitácora (todas las financieras).")
    if base is not None:
        bit = conexion.bitacora_guardada(base, con_sheet=False)    # solo bitácoras: el Sheet es lo que se revisa
        st.success(f"Bitácora: {conexion.describir_bitacora(bit)}", icon=":material/cloud_done:")
        st.caption(AGREGAR_MES)
        if st.button("Hacer cuadre", type="primary", disabled=bit.empty, key="btn_seg"):
            with st.spinner("Cruzando…"):
                try:
                    st.session_state["res_seg"] = ejecutar_seg_base(base, bit)
                except Exception as e:
                    st.session_state.pop("res_seg", None)
                    st.error(f"No se pudo hacer el cuadre: {e}")
    else:
        c1, c2 = st.columns(2)
        f_seg = c1.file_uploader("Excel de seguimiento", type=["xlsx"], key="up_seg",
                                 help="Solo se leen las pestañas de COLIMA.")
        f_cet = c2.file_uploader("Capturas de Cetelem", type=["xlsx"], accept_multiple_files=True,
                                 key="up_cet_seg", help=AYUDA_CAPTURAS)
        c2.caption("Sin el Sheet conectado solo se cuadra Cetelem.")

        if st.button("Hacer cuadre", type="primary", disabled=not (f_seg and f_cet), key="btn_seg"):
            with st.spinner("Cruzando…"):
                try:
                    st.session_state["res_seg"] = ejecutar(cs, f_seg, f_cet)
                except Exception as e:
                    st.session_state.pop("res_seg", None)
                    st.error(f"No se pudo hacer el cuadre: {e}")

    res = st.session_state.get("res_seg")
    if res:
        r = res["resumen"]
        st.success(f"Bitácora: {', '.join(res['meses'])}  ·  Financieras: {', '.join(res.get('financieras', []))}"
                   f"  ·  Pestañas: {', '.join(res['hojas'])}")
        m = st.columns(5)
        m[0].metric("Rechazados", valor_resumen(r, "Clientes únicos rechazados"))
        m[1].metric("Encontrados", valor_resumen(r, "Clientes rechazados encontrados"))
        m[2].metric("Faltan", valor_resumen(r, "Clientes rechazados NO"))
        m[3].metric("% cuadre", valor_resumen(r, "Porcentaje"))
        m[4].metric("A corregir", valor_resumen(r, "Datos a corregir"))
        descargas(res, "seg")

        t1, t2, t3, t4, t5 = st.tabs([
            f"Faltan en seguimiento ({len(res['faltantes'])})",
            f"A corregir ({len(res['revisar'])})",
            f"No están en la bitácora ({len(res['solo_periodo'])})",
            "Completo", "Resumen"])
        with t1:
            df = filtro_vendedor(res["faltantes"], "fv_seg")
            if not df.empty:
                st.bar_chart(df.groupby("Vendedor").size().rename("Clientes"), color="#C00000",
                             height=60 + 28 * df["Vendedor"].nunique(), horizontal=True)
            tabla(df, ["Cliente_bitacora", "Tel_bitacora", "Vendedor", "Financiera", "Folios", "Historial",
                       "Observaciones"])
        with t2:
            tabla(res["revisar"], ["Resultado", "Cuadre_pestana", "Cliente_bitacora", "Tel_bitacora",
                                   "Cliente_seguimiento", "Tel_seguimiento", "Folio_seguimiento",
                                   "Pestanas_encontrado", "Observaciones"],
                  ["Resultado", "Cuadre_pestana"])
        with t3:
            st.caption("Rechazos del seguimiento, dentro de los meses de la bitácora, que ahí no salen como "
                       "rechazo. *Observaciones* dice con qué estatus aparecen.")
            tabla(res["solo_periodo"], ["Hoja", "Fila_excel", "Fecha_ingreso", "Folio_norm", "Nombre",
                                        "Telefono", "Asesor", "Observaciones"])
            with st.expander("Fuera del periodo, sin fecha u otras financieras"):
                tabla(res["solo"], ["Clasificacion", "Hoja", "Fila_excel", "Fecha_ingreso", "Folio_norm",
                                    "Nombre", "Telefono", "Financiera", "Asesor", "Observaciones"],
                      ["Clasificacion"])
        with t4:
            df = filtro_vendedor(res["cuadre"], "fv_seg_todo")
            tabla(df, cs.COLS_CUADRE, ["Resultado", "Cuadre_pestana"], alto=550)
        with t5:
            st.dataframe(r.astype(str), hide_index=True, width="stretch")

# ---------------------------------------------------------------- SALE-U
with tab_su:
    c1, c2 = st.columns(2)
    f_su = c1.file_uploader("Reporte de Sale-U", type=["csv", "xlsx"], key="up_su",
                            help="reporteLeadsconCredito…, tal cual se descarga (sin abrirlo en Excel).")
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
                    st.session_state["res_su"] = ejecutar_su_base(base, f_su, bit)
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
        st.markdown(f"**Registro en Sale-U** · {res['periodo_saleu']}")
        datos = res.get("datos", pd.DataFrame())
        aplicado = res.get("aplicado", False)
        escritos_datos = res.get("datos_escritos", 0)
        k = st.columns(4)
        k[0].metric("Rechazos revisados", len(reg))
        k[1].metric("Registrados", int((reg["Después"] == saleu_sheet.REGISTRADO).sum()))
        k[2].metric("Celdas actualizadas" if aplicado else "Celdas por actualizar",
                    res["escritos"] if aplicado else len(cambian),
                    help="Registrado / No Registrado. Nunca baja un «Registrado» a «No Registrado».")
        k[3].metric("Datos llenados" if aplicado else "Datos faltantes", escritos_datos if aplicado else len(datos),
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
            st.info(f"**{len(rev)} marcado(s) como Registrado no salen en este reporte.** Se dejaron igual; "
                    "confírmalos en Sale-U.", icon=":material/fact_check:")
            tabla(rev.assign(Fecha_ingreso=rev["Fecha_ingreso"].map(lambda d: d.strftime("%d/%m/%Y") if d else "")),
                  ["Fila", "Cliente", "Fecha_ingreso", "Antes", "Después"])
        with st.expander(f"Detalle del registro ({len(cambian)} cambio(s) de {len(reg)})"):
            tabla(reg.assign(Fecha_ingreso=reg["Fecha_ingreso"].map(lambda d: d.strftime("%d/%m/%Y") if d else "")),
                  ["Fila", "Cliente", "Fecha_ingreso", "Antes", "Después", "Como", "Cliente_SaleU", "Estatus_SaleU",
                   "Financiera_SaleU"])
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
                por_est = " · ".join(f"{k.capitalize()} {v}" for k, v in est.value_counts().items())
                st.caption(f"Por estatus: {por_est} · En el Sheet: **{n_hoja}** rechazos *No Registrado*.")
            df = filtro_vendedor(res["faltantes"], "fv_su")
            if not df.empty:
                st.bar_chart(df.groupby("Vendedor").size().rename("Clientes"), color="#C00000",
                             height=60 + 28 * df["Vendedor"].nunique(), horizontal=True)
            vista = cu._faltantes_con_vendedor(df).assign(Historial=df["Historial"].values,
                                                          Observaciones=df["Observaciones"].values)
            vista["Fecha"] = vista["Fecha"].map(cu.fecha_txt)
            vista["Mensualidad"] = vista["Mensualidad"].map(lambda v: f"${v:,.2f}" if v == v and v else "")
            vista["Plazo"] = vista["Plazo"].map(lambda v: str(int(v)) if v == v and v else "")
            tabla(vista, cu.COLS_FALTANTES[:2] + ["Vendedor"] + cu.COLS_FALTANTES[2:] + ["Historial", "Observaciones"])
        with t2:
            tabla(res["revisar"], ["Resultado", "Cuadre_estatus", "Cliente_bitacora", "Tel_bitacora",
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
