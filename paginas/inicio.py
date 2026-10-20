# -*- coding: utf-8 -*-
"""
Página: Inicio

Lo más importante al abrir la app: los indicadores del año, las prioridades, qué archivos faltan por subir y el
resumen ejecutivo de una página en PDF.
"""
import datetime as dt
import re

import streamlit as st

import ui
import conexion
from comunes import ESTADO
from cuadres import cargas, reporte_ejecutivo

ui.hero("Inicio", "Lo más importante del seguimiento de MG Colima.", f"Corte {dt.date.today():%d/%m/%Y}")

base = conexion.panel_base()
if base is None:
    st.stop()

try:
    d = reporte_ejecutivo.datos(base)
except Exception as e:
    st.error(f"No se pudo leer el Sheet: {e}")
    st.stop()
k = d["k"]
pct = lambda a, b: f"{a / b:.0%}" if b else "—"
al_dia = k["al_dia"] / k["ya_toca"] if k["ya_toca"] else 1
f1, f2 = st.columns(3), st.columns(3)
ui.tarjeta(f1[0], "event_busy", "Rechazos en el año", k["rechazos"],
           " · ".join(f"{f} {n}" for f, n in sorted(k["por_fin"].items(), key=lambda x: -x[1])))
ui.tarjeta(f1[1], "event_available", "Seguimiento al día", pct(k["al_dia"], k["ya_toca"]),
           f"{k['al_dia']} de {k['ya_toca']} que ya cumplieron 180 días",
           ESTADO["bien"] if al_dia >= 0.9 else ESTADO["aviso"])
ui.tarjeta(f1[2], "call", "Contacto efectivo", pct(k["contestaron"], k["contactados"]),
           f"{k['contestaron']} de {k['contactados']} contactados")
ui.tarjeta(f2[0], "how_to_reg", "Registrados en Sale-U", pct(k["registrados"], k["saleu_total"]),
           f"{k['falta_alta']} falta alta · {k['falta_credito']} falta crédito")
ui.tarjeta(f2[1], "alarm", "Seguimientos vencidos", len(k["vencidos"]), f"{len(k['por_vencer'])} vencen en 15 días",
           ESTADO["critico"] if len(k["vencidos"]) else ESTADO["bien"])
ui.tarjeta(f2[2], "rule", "Pendientes en Revision", k["pendientes"], "por resolver",
           ESTADO["aviso"] if k["pendientes"] else ESTADO["bien"])

c1, c2 = st.columns([3, 2])
with c1.container(border=True):
    st.markdown(":material/lightbulb: **Prioridades**")
    prio = reporte_ejecutivo.prioridades(d)
    for p in prio or ["Todo al día."]:
        st.markdown("- " + re.sub(r"</?b>", "**", p))
with c2.container(border=True):
    st.markdown(":material/upload_file: **Archivos**")
    try:
        avisos = [e for e in cargas.estado(base) if e["Aviso"]]
    except Exception:
        avisos = []
    for e in avisos:
        st.markdown(f":material/warning: {e['Aviso']}")
    if not avisos:
        st.markdown(":material/check_circle: Todo lo subido está al día.")
    try:
        st.page_link("paginas/actualizar.py", label="Ir a Actualizar", icon=":material/arrow_forward:")
    except Exception:  # fuera de la navegación (pruebas) no hay a dónde ligar
        pass

if st.button("Resumen ejecutivo (PDF)", type="primary", icon=":material/picture_as_pdf:", key="btn_ejec"):
    with st.spinner("Armando el resumen…"):
        try:
            st.session_state["pdf_ejec"] = reporte_ejecutivo.pdf(base, d=d)
        except Exception as e:
            st.error(f"No se pudo generar: {e}")
if st.session_state.get("pdf_ejec"):
    st.download_button("Descargar resumen ejecutivo", st.session_state["pdf_ejec"],
                       file_name=f"Resumen_ejecutivo_MG_Colima_{dt.date.today():%Y%m%d}.pdf", mime="application/pdf",
                       icon=":material/download:", key="dl_ejec")
