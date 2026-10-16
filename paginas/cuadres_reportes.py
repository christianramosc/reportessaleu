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
from cuadres import fuentes, saleu_sheet
from cuadres.base import a_excel, norm

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
    vista.columns = [c.replace("_", " ") for c in cols]
    estilo = vista.style
    color_cols = [c.replace("_", " ") for c in columnas_color if c in cols]
    if color_cols:
        estilo = estilo.map(colorear, subset=color_cols)
    extra = {"height": alto} if alto else {}
    st.dataframe(estilo, width="stretch", hide_index=True, **extra)


def filtro_vendedor(df, clave, columna="Vendedor"):
    if df.empty or columna not in df.columns:
        return df
    opciones = sorted(v for v in df[columna].dropna().unique() if str(v).strip())
    elegidos = st.multiselect("Filtrar por vendedor", opciones, key=clave, placeholder="Todos los vendedores")
    return df[df[columna].isin(elegidos)] if elegidos else df


def valor_resumen(resumen, texto):
    fila = resumen[resumen["Concepto"].str.strip().str.startswith(texto)]
    return fila["Valor"].iloc[0] if not fila.empty else "—"


def descargas(res, prefijo):
    c1, c2, _ = st.columns([1, 1, 3])
    c1.download_button("⬇️ Descargar Excel", res["xlsx_bytes"], file_name=res["xlsx_nombre"],
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       width="stretch", key=f"{prefijo}_x")
    c2.download_button("⬇️ Descargar PDF", res["pdf_bytes"], file_name=res["pdf_nombre"],
                       mime="application/pdf", width="stretch", key=f"{prefijo}_p")


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


def ejecutar_su_base(base, archivo_su, bitacora, escribir):
    """Cuadre Sale-U contra la bitácora (todas las financieras) + registro en Sale-U escrito en la base."""
    with tempfile.TemporaryDirectory() as tmp:
        ruta_su = guardar_temporal([archivo_su], tmp)[0]
        res = {}
        if not bitacora.empty:
            res = cu.procesar(ruta_su, carpeta=tmp, bitacora=bitacora)
            res["xlsx_bytes"], res["pdf_bytes"] = leer_bytes(res["xlsx"]), leer_bytes(res["pdf"])
            res["xlsx_nombre"], res["pdf_nombre"] = os.path.basename(res["xlsx"]), os.path.basename(res["pdf"])
        su = cu.leer_saleu(ruta_su)
    cambios, (ini, fin) = saleu_sheet.proponer(base, su)
    periodo = f"{ini:%d/%m/%Y} – {fin:%d/%m/%Y}"
    res["registro"] = cambios
    res["periodo_saleu"] = periodo
    res["escritos"] = saleu_sheet.aplicar(base, cambios, periodo) if escribir else 0
    return res


# =============================================================================
# Barra lateral: ajustes
# =============================================================================
with st.sidebar:
    st.header("⚙️ Ajustes")
    st.caption("Qué tan estricta es la comparación de nombres (aplica a los dos cuadres).")
    fuerte = st.slider("Coincidencia solo por nombre", 0.6, 1.0, 0.80, 0.05,
                       help="Parecido mínimo para aceptar un cliente con teléfono distinto.")
    con_tel = st.slider("Nombre cuando el teléfono coincide", 0.3, 1.0, 0.50, 0.05,
                        help="Si el teléfono es igual, cuánto debe parecerse el nombre.")
    for m in (cs, cu):
        m.UMBRAL_NOMBRE_FUERTE, m.UMBRAL_NOMBRE_CON_TEL = fuerte, con_tel

    st.divider()
    st.subheader("Sale-U")
    st.caption("Cómo debe verse en Sale-U una contrapropuesta:")
    contra = st.selectbox("Contrapropuesta", ["No se juzga (revisar)", "Aceptado", "Rechazado"])
    cu.ESTATUS_ESPERADO["CONTRAPROPUESTA"] = {"Aceptado": "ACEPTADO",
                                              "Rechazado": "RECHAZADO"}.get(contra)


# =============================================================================
# Página
# =============================================================================
ui.hero("Cuadres para reportes",
        "Elige el cuadre, sube los archivos fuente y descarga el resultado en Excel y PDF.")

tab_seg, tab_su, tab_bs = st.tabs(["📒 Seguimiento de rechazos vs bitácora",
                                   "🟦 Sale-U vs bitácora",
                                   "🔧 BlueService vs Inventario"])

AGREGAR_MES = ("La bitácora es la pestaña **Bitacora** del Sheet: junta las capturas de Cetelem y la Bitácora de F&I "
               "(todas las financieras). Para agregar un mes, súbelo en **Base de seguimiento → Actualizar con "
               "bitácoras** y aplica; de ahí la toman todos los cuadres.")
AYUDA_CAPTURAS = ("Se solicitan al F&I. Son todas las capturas mensuales de la financiera CETELEM "
                  "(un Excel por mes, ej. JULIO_2026_FINAL.xlsx, AGOSTO_2026_FINAL.xlsx). "
                  "Puedes subir varios meses a la vez.")

# ---------------------------------------------------------------- SEGUIMIENTO
with tab_seg:
    base = conexion.obtener_base()
    st.markdown("Compara la base de seguimiento contra las **solicitudes RECHAZADAS** de la bitácora: capturas de "
                "Cetelem y Bitácora de F&I, de **todas las financieras** (la base solo registra rechazos).")
    if base is not None:
        bit = conexion.bitacora_guardada(base)
        st.success(f"Seguimiento: se lee de **{base.titulo}** · Bitácora: {conexion.describir_bitacora(bit)}",
                   icon=":material/cloud_done:")
        st.caption(AGREGAR_MES)
        if st.button("Hacer cuadre", type="primary", disabled=bit.empty, key="btn_seg"):
            with st.spinner("Cruzando folios, teléfonos y nombres…"):
                try:
                    st.session_state["res_seg"] = ejecutar_seg_base(base, bit)
                except Exception as e:
                    st.session_state.pop("res_seg", None)
                    st.error(f"No se pudo hacer el cuadre: {e}")
    else:
        c1, c2 = st.columns(2)
        f_seg = c1.file_uploader("Seguimiento Rechazos Financiera MG", type=["xlsx"], key="up_seg",
                                 help="Sube el archivo con el formato 'Seguimiento Rechazos Financiera MG.xlsx'. "
                                      "Solo se leen las pestañas de COLIMA.")
        c1.caption("📄 Formato: **Seguimiento Rechazos Financiera MG.xlsx** · Conecta el Google Sheet (ver README) "
                   "para no tener que subirlo.")
        f_cet = c2.file_uploader("Capturas mensuales CETELEM", type=["xlsx"], accept_multiple_files=True,
                                 key="up_cet_seg", help=AYUDA_CAPTURAS)
        c2.caption("📄 Se piden al **F&I**: todas las capturas mensuales de la financiera CETELEM. Sin la base "
                   "conectada solo se cuadra Cetelem.")

        if st.button("Hacer cuadre", type="primary", disabled=not (f_seg and f_cet), key="btn_seg"):
            with st.spinner("Cruzando folios, teléfonos y nombres…"):
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
        m[0].metric("Clientes rechazados", valor_resumen(r, "Clientes únicos rechazados"))
        m[1].metric("Encontrados", valor_resumen(r, "Clientes rechazados encontrados"))
        m[2].metric("Faltan capturar", valor_resumen(r, "Clientes rechazados NO"))
        m[3].metric("% de cuadre", valor_resumen(r, "Porcentaje"))
        m[4].metric("Datos a corregir", valor_resumen(r, "Datos a corregir"))
        descargas(res, "seg")

        t1, t2, t3, t4, t5 = st.tabs([
            f"Rechazos que faltan en seguimiento ({len(res['faltantes'])})",
            f"A corregir / revisar ({len(res['revisar'])})",
            f"En seguimiento, no en la bitácora ({len(res['solo_periodo'])})",
            "Cuadre completo", "Resumen"])
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
            st.caption("Rechazos en el seguimiento, de una financiera que trae la bitácora y con fecha dentro de sus meses, que no aparecen como rechazados en la bitácora. La columna Observaciones dice con qué estatus sí aparecen.")
            tabla(res["solo_periodo"], ["Hoja", "Fila_excel", "Fecha_ingreso", "Folio_norm", "Nombre",
                                        "Telefono", "Asesor", "Observaciones"])
            with st.expander("Ver también fuera del periodo, sin fecha y otras financieras"):
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
    base = conexion.obtener_base()
    c1, c2 = st.columns(2)
    f_su = c1.file_uploader("Excel de Sale - U", type=["csv", "xlsx"], key="up_su",
                            help="Sube el reporte de Sale - U tal cual se descarga "
                                 "(reporteLeadsconCredito…), sin abrirlo ni guardarlo antes en Excel.")
    c1.caption("📄 Archivo: **reporte de Sale - U** (reporteLeadsconCredito), tal cual se descarga.")
    if base is not None:
        bit = conexion.bitacora_guardada(base)
        c2.success(f"Bitácora desde **{base.titulo}**: {conexion.describir_bitacora(bit)}",
                   icon=":material/cloud_done:")
        c2.caption(AGREGAR_MES)
        escribir = c1.toggle("Escribir «Cliente con Registro en SalesU» en el Sheet", value=True, key="tg_su",
                             help="Marca Registrado / No Registrado a cada rechazo cuya fecha de ingreso cae dentro "
                                  "del periodo del reporte de Sale-U. Lo demás no se toca.")
        if st.button("Hacer cuadre", type="primary", disabled=not f_su, key="btn_su"):
            with st.spinner("Cruzando teléfonos y nombres…"):
                try:
                    st.session_state["res_su"] = ejecutar_su_base(base, f_su, bit, escribir)
                except Exception as e:
                    st.session_state.pop("res_su", None)
                    st.error(f"No se pudo hacer el cuadre: {e}")
    else:
        f_cet2 = c2.file_uploader("Capturas mensuales CETELEM", type=["xlsx"], accept_multiple_files=True,
                                  key="up_cet_su", help=AYUDA_CAPTURAS)
        c2.caption("📄 Se piden al **F&I**: todas las capturas mensuales de la financiera CETELEM. Sin la base "
                   "conectada solo se cuadra Cetelem.")

        if st.button("Hacer cuadre", type="primary", disabled=not (f_su and f_cet2), key="btn_su"):
            with st.spinner("Cruzando teléfonos y nombres…"):
                try:
                    st.session_state["res_su"] = ejecutar(cu, f_su, f_cet2)
                except Exception as e:
                    st.session_state.pop("res_su", None)
                    st.error(f"No se pudo hacer el cuadre: {e}")

    res = st.session_state.get("res_su")
    if res and "registro" in res:
        reg = res["registro"]
        cambian = reg[reg["Cambia"]]
        st.markdown(f"**Registro en Sale-U en el Sheet** · periodo del reporte {res['periodo_saleu']}")
        k = st.columns(3)
        k[0].metric("Rechazos del periodo revisados", len(reg))
        k[1].metric("Registrados en Sale-U", int((reg["Después"] == saleu_sheet.REGISTRADO).sum()))
        k[2].metric("Celdas actualizadas en el Sheet", res["escritos"])
        if res["escritos"]:
            st.success(f"Se actualizó «Cliente con Registro en SalesU» en {res['escritos']} fila(s); quedó en el Historial.")
        elif len(cambian):
            st.warning(f"{len(cambian)} fila(s) cambiarían, pero la escritura en el Sheet estaba apagada.")
        with st.expander(f"Ver detalle ({len(cambian)} cambio(s) de {len(reg)} revisados)"):
            tabla(reg.assign(Fecha_ingreso=reg["Fecha_ingreso"].map(lambda d: d.strftime("%d/%m/%Y") if d else "")),
                  ["Fila", "Cliente", "Fecha_ingreso", "Antes", "Después", "Como", "Cliente_SaleU", "Estatus_SaleU",
                   "Financiera_SaleU"])
        if isinstance(base, conexion.BaseExcel):
            st.info("Modo Excel: descarga el archivo actualizado en la página «Base de seguimiento».")
    if res and "resumen" in res:
        r = res["resumen"]
        st.success(f"Bitácora: {', '.join(res['meses'])}  ·  Financieras: {', '.join(res.get('financieras', []))}")
        m = st.columns(5)
        m[0].metric("Clientes en bitácora", valor_resumen(r, "Clientes únicos"))
        m[1].metric("En Sale-U", valor_resumen(r, "Clientes de la bitácora encontrados"))
        m[2].metric("Faltan en Sale-U", valor_resumen(r, "Clientes de la bitácora NO"))
        m[3].metric("% de cuadre", valor_resumen(r, "Porcentaje"))
        m[4].metric("Estatus no cuadra", valor_resumen(r, "Estatus que no cuadra"))
        descargas(res, "su")

        t1, t2, t3, t4, t5 = st.tabs([
            f"Faltan en Sale-U ({len(res['faltantes'])})",
            f"A revisar ({len(res['revisar'])})",
            f"Solo en Sale-U ({len(res['solo_cet'])})",
            "Cuadre completo", "Resumen"])
        with t1:
            df = filtro_vendedor(res["faltantes"], "fv_su")
            if not df.empty:
                st.bar_chart(df.groupby("Vendedor").size().rename("Clientes"), color="#C00000",
                             height=60 + 28 * df["Vendedor"].nunique(), horizontal=True)
            tabla(df, ["Cliente_bitacora", "Tel_bitacora", "Vendedor", "Financiera", "Folios", "Historial",
                       "Estatus_final"])
        with t2:
            tabla(res["revisar"], ["Resultado", "Cuadre_estatus", "Cliente_bitacora", "Tel_bitacora",
                                   "Cliente_SaleU", "Tel_SaleU", "Estatus_SaleU", "Financiera_SaleU",
                                   "Observaciones"], ["Resultado", "Cuadre_estatus"])
        with t3:
            st.caption("Registros en Sale-U que no aparecen en la bitácora (pueden ser de meses anteriores). "
                       "Los de financieras que la bitácora no trae se ven en el Excel descargable.")
            tabla(res["solo_cet"], ["Nombre", "Telefono", "Asesor", "Estatus_SaleU", "Financiera",
                                    "Fecha_Alta_Credito", "Observaciones"])
        with t4:
            df = filtro_vendedor(res["cuadre"], "fv_su_todo")
            tabla(df, ["Resultado", "Cuadre_estatus", "Cliente_bitacora", "Tel_bitacora", "Vendedor",
                       "Financiera", "Historial", "Cliente_SaleU", "Tel_SaleU", "Asesor_SaleU", "Estatus_SaleU",
                       "Financiera_SaleU", "Observaciones"], ["Resultado", "Cuadre_estatus"], alto=550)
        with t5:
            st.dataframe(r.astype(str), hide_index=True, width="stretch")

# ---------------------------------------------------------------- BLUESERVICE
with tab_bs:
    c1, c2 = st.columns(2)
    f_pdi = c1.file_uploader("Reporte de PDIs de BlueService", type=["xlsx"], key="up_pdi",
                             help="Ej. PDIs_MG_Colima.xlsx")
    f_inv = c2.file_uploader("Excel de inventario", type=["xlsx"], key="up_inv",
                             help="Ej. INVENTARIO_MG_COLIMA….xlsx")
    hoja = None
    if f_inv:
        try:
            hojas = pd.ExcelFile(f_inv).sheet_names
            hoja = c2.selectbox("Hoja de inventario a cuadrar", hojas,
                                index=hojas.index(cb.hoja_por_defecto(hojas)),
                                help="Las demás hojas se usan para decir dónde aparece cada VIN "
                                     "(vendidos, traslados, servicio).")
        except Exception as e:
            c2.error(f"No pude leer las hojas del archivo: {e}")

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
        st.success(f"Hoja de inventario: {h}  ·  Fecha de corte: {cb.fmt(res['corte'])}")
        m = st.columns(5)
        m[0].metric(f"Unidades en {h.title()}", valor_resumen(r, "Unidades en la hoja"))
        m[1].metric("Cuadran con BlueService", valor_resumen(r, f"Unidades de {h} que cuadran"))
        m[2].metric("Faltan en BlueService", valor_resumen(r, f"Unidades de {h} que faltan"))
        m[3].metric("En BlueService, no en inventario", valor_resumen(r, "En BlueService sin venta"))
        m[4].metric("Datos a revisar", valor_resumen(r, "Unidades con datos"))
        descargas(res, "bs")

        cols_inv = ["Resultado", "VIN", "Modelo_inventario", "Version", "Color", "Ubicacion",
                    "Entrada_inventario", "Ultimo_PDI_inventario", "Proximo_PDI_inventario",
                    "Etapa_BS", "Status_BS", "Observaciones"]
        t1, t2, t3, t4, t5, t6 = st.tabs([
            f"Faltan en BlueService ({len(res['falta_bs'])})",
            f"En BlueService, no en {h.title()} ({len(res['falta_inv'])})",
            f"A revisar ({len(res['revisar'])})",
            f"{h.title()} completo", f"Vendidas en BS ({len(res['vendidas'])})", "Resumen"])
        with t1:
            st.caption(f"Unidades físicas en {h.lower()} que no aparecen en el reporte de PDIs de BlueService.")
            tabla(res["falta_bs"], cols_inv, ["Resultado"])
        with t2:
            df = res["falta_inv"]
            if not df.empty:
                sit = st.multiselect("Filtrar por situación", sorted(df["Situacion"].unique()), key="f_sit",
                                     placeholder="Todas las situaciones")
                if sit:
                    df = df[df["Situacion"].isin(sit)]
                st.bar_chart(res["falta_inv"].groupby("Situacion").size().rename("Unidades"),
                             color="#C00000", height=200, horizontal=True)
            tabla(df, ["Situacion", "VIN", "Modelo_BS", "Dealer_origen", "Dealer_actual",
                       "Fecha_inventario_BS", "Etapa_BS", "Status_BS", "Aparece_en_inventario",
                       "Observaciones"], ["Situacion"])
        with t3:
            st.caption("VIN mal capturados, fechas de PDI inválidas, unidades vendidas que siguen en "
                       "inventario y diferencias de PDI.")
            tabla(res["revisar"], ["Resultado", "VIN", "VIN_BlueService"] + cols_inv[2:], ["Resultado"])
        with t4:
            tabla(res["cuadre_inv"], ["Resultado", "VIN", "VIN_BlueService"] + cols_inv[2:] + ["Fila_excel"],
                  ["Resultado"], alto=550)
        with t5:
            st.caption("Informativo: unidades que BlueService ya tiene con fecha de venta.")
            tabla(res["vendidas"], ["Situacion", "VIN", "Modelo_BS", "Fecha_venta_BS", "Etapa_BS",
                                    "Status_BS", "Aparece_en_inventario"])
        with t6:
            st.dataframe(r.astype(str), hide_index=True, width="stretch")
