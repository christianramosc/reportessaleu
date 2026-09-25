# -*- coding: utf-8 -*-
"""Página: Cuadre Sale-U, MG Colima (compara por teléfono seguimiento vs. Sale-U)."""
import io
from datetime import date

import streamlit as st

import cuadre_core as cc
import ui


@st.cache_data(show_spinner="Cuadrando teléfonos…")
def procesar_cuadre(saleu_bytes: bytes, seg_bytes: bytes, hojas: tuple):
    saleu = cc.cargar_saleu(io.BytesIO(saleu_bytes))
    tablas = [cc.cargar_hoja(io.BytesIO(seg_bytes), h) for h in hojas]
    cuadre = cc.cuadrar(saleu, tablas)
    obs = cc.generar_observaciones(cuadre, saleu)
    return cuadre, obs, cc.generar_excel(cuadre, obs)


@st.cache_data(show_spinner=False)
def hojas_disponibles(seg_bytes: bytes):
    return cc.hojas_colima(io.BytesIO(seg_bytes))


ui.hero("Cuadre Sale-U", "Rechazos y créditos finalizados contra la exportación de Sale-U, MG Colima",
        "Cruce por teléfono")

c1, c2 = st.columns(2)
f_saleu = ui.archivo_persistente("Exportación de Sale-U (.xlsx)", "cuadre_saleu", ["xlsx"], donde=c1)
f_seg = ui.archivo_persistente("Seguimiento de rechazos (.xlsx)", "cuadre_seg", ["xlsx"], donde=c2)
if not (f_saleu and f_seg):
    st.markdown("""<div class="vacio"><b>Sube los dos archivos para comenzar</b><br>
    La exportación de leads de Sale-U y el archivo de seguimiento con las hojas de
    MG Colima Rechazos y MG Colima Créditos Finalizados.</div>""", unsafe_allow_html=True)
    st.stop()

try:
    disponibles = hojas_disponibles(f_seg[1])
except Exception as e:  # noqa: BLE001
    st.error(f"No se pudo leer el archivo de seguimiento. {e}")
    st.stop()
if not disponibles:
    st.error("El archivo de seguimiento no tiene hojas de MG Colima (Rechazos o Créditos Finalizados).")
    st.stop()

elegidas = st.multiselect("Hojas a cuadrar", disponibles, default=disponibles, key="cuadre_hojas")
if not elegidas:
    st.info("Elige al menos una hoja para cuadrar.")
    st.stop()

try:
    cuadre, obs, excel = procesar_cuadre(f_saleu[1], f_seg[1], tuple(elegidas))
except KeyError as e:
    st.error(str(e).strip("'\""))
    st.stop()

falt = cuadre[cuadre["Estatus Cuadre"] == cc.FALTA]
total, n_ok = len(cuadre), int((cuadre["Estatus Cuadre"] == cc.EN_SALEU).sum())

st.subheader("Resumen")
m = st.columns(4)
m[0].metric("Registros capturados", total)
m[1].metric("En Sale-U", n_ok, f"{n_ok / total * 100:.1f}%" if total else None)
m[2].metric("Faltan en Sale-U", len(falt), delta_color="inverse")
m[3].metric("Clientes faltantes", falt["Teléfono"].fillna(falt["Nombre de Cliente"]).nunique())

res = (cuadre.groupby(["Origen", "Estatus Cuadre"]).size().unstack(fill_value=0)
       .reindex(columns=[cc.EN_SALEU, cc.FALTA], fill_value=0))
res["Total"] = res.sum(axis=1)
res["% cuadrado"] = (res[cc.EN_SALEU] / res["Total"] * 100).round(1).astype(str) + "%"
st.dataframe(res, width="stretch")

st.subheader("Observaciones importantes")
for nivel, texto in obs:
    {"alta": st.error, "media": st.warning, "ok": st.success, "info": st.info}[nivel](texto)

st.download_button("Descargar Excel del cuadre", excel,
                   file_name=f"Cuadre_SaleU_MG_Colima_{date.today():%Y%m%d}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   type="primary", icon=":material/download:")

st.subheader("Detalle")
cols_falt = ["Origen", "Fecha ingreso", "Nombre de Cliente", "Teléfono", "Financiera", "Asesor",
             "Seguimiento SalesU (Excel)", "Posible coincidencia por nombre", "Observaciones"]
tabs = st.tabs([f"Faltan en Sale-U ({len(falt)})", "Faltantes por asesor", "Cuadre completo"])
with tabs[0]:
    st.dataframe(falt[cols_falt], width="stretch", hide_index=True)
with tabs[1]:
    if len(falt):
        pa = falt.groupby(["Asesor", "Origen"]).size().unstack(fill_value=0)
        pa["Total"] = pa.sum(axis=1)
        st.dataframe(pa.sort_values("Total", ascending=False), width="stretch")
    else:
        st.success("No hay faltantes.")
with tabs[2]:
    filtro = st.radio("Mostrar", ["Todos", cc.EN_SALEU, cc.FALTA], horizontal=True, key="cuadre_filtro")
    vista = cuadre if filtro == "Todos" else cuadre[cuadre["Estatus Cuadre"] == filtro]
    st.dataframe(vista, width="stretch", hide_index=True)
