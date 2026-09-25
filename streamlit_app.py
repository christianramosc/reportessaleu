# -*- coding: utf-8 -*-
"""App de Streamlit: KPIs de capacitación Ventas / Posventa · MG Colima."""
import pandas as pd
import streamlit as st

import kpis_core as kc

st.set_page_config(page_title="KPIs Capacitación · MG Colima", page_icon="📊", layout="wide")


@st.cache_data(show_spinner="Generando gráficas...")
def procesar(contenido: bytes):
    return kc.generar(contenido)


st.title("📊 KPIs de Capacitación · MG Colima")
st.caption("Sube el archivo PLAN DE ACCIÓN (.xlsx). Se leen las hojas "
           f"**{kc.HOJAS['VENTAS']}** y **{kc.HOJAS['POSVENTA']}**.")

archivo = st.file_uploader("Archivo Excel", type=["xlsx", "xlsm"])
if archivo is None:
    st.info("Esperando archivo…")
    st.stop()

try:
    out = procesar(archivo.getvalue())
except Exception as e:  # noqa: BLE001
    st.error(f"No se pudo procesar el archivo: {e}")
    st.stop()

# ---------- descargas generales
c1, c2 = st.columns(2)
c1.download_button("⬇️ Descargar PDF completo", out["pdf"], file_name=kc.SALIDA,
                   mime="application/pdf", use_container_width=True)
c2.download_button("⬇️ Descargar todas las imágenes (ZIP)", out["zip"],
                   file_name="KPIs_Capacitacion_imagenes.zip", mime="application/zip",
                   use_container_width=True)

st.caption("Los totales y avances se recalculan desde las filas de cada colaborador. "
           "Avance por módulo = promedio del avance individual (concluidas / total); "
           "avance total = promedio de los 4 módulos.")

ETIQUETAS = {
    "0_resumen_completo": "Resumen completo",
    "1_tarjetas_modulos": "Tarjetas por módulo",
    "2_avance_por_letra": "Avance por letra",
    "3_avance_promedio_general": "Avance promedio general",
    "4_detalle_por_colaborador": "Detalle por colaborador",
}


def imagen(area, clave):
    nombre = f"{area.lower()}_{clave}.png"
    st.image(out["imagenes"][nombre], use_container_width=True)
    st.download_button(f"⬇️ {ETIQUETAS[clave]} (PNG)", out["imagenes"][nombre],
                       file_name=nombre, mime="image/png", key=nombre)


for area, tab in zip(kc.HOJAS, st.tabs([a.title() for a in kc.HOJAS])):
    d = out["resumen"][area]
    with tab:
        # métricas rápidas
        cols = st.columns(5)
        cols[0].metric("Avance total", kc.pct(d["general"]))
        for col, k in zip(cols[1:], kc.ORDEN):
            r = d["modulos"][k]
            col.metric(kc.MODULOS[k][0], kc.pct(r["avance"]),
                       f"{r['concluidas']:.0f} de {r['total']:.0f}", delta_color="off")

        imagen(area, "1_tarjetas_modulos")
        a, b = st.columns(2)
        with a:
            imagen(area, "2_avance_por_letra")
        with b:
            imagen(area, "3_avance_promedio_general")

        with st.expander("Detalle por colaborador"):
            imagen(area, "4_detalle_por_colaborador")
            filas = []
            for k in kc.ORDEN:
                for x in d["tablas"][k]:
                    filas.append({"Colaborador": x["nombre"], "Letra": x["letra"],
                                  "Módulo": kc.MODULOS[k][0], "Avance": x["avance"]})
            df = (pd.DataFrame(filas)
                  .pivot_table(index=["Colaborador", "Letra"], columns="Módulo",
                               values="Avance", sort=False)
                  [[kc.MODULOS[k][0] for k in kc.ORDEN]]
                  .reset_index())
            st.dataframe(df.style.format({kc.MODULOS[k][0]: "{:.0%}" for k in kc.ORDEN}),
                         use_container_width=True, hide_index=True)

        with st.expander("Página de resumen completa"):
            imagen(area, "0_resumen_completo")
