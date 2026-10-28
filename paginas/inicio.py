# -*- coding: utf-8 -*-
"""
Página: Inicio

Lo más importante al abrir la app: los indicadores del año, las prioridades, qué archivos faltan por subir, un
tablero de gráficas interactivas (sin nombres de clientes ni de asesores) y el resumen ejecutivo en PDF.
"""
import datetime as dt
import re

import streamlit as st

import ui
import conexion
from comunes import ESTADO
from cuadres import cargas, reporte_ejecutivo, resumen
import tablero

ui.encabezado("Inicio", "Lo más importante del seguimiento de MG Colima.", f"Corte {dt.date.today():%d/%m/%Y}")

base = conexion.panel_base()
if base is None:
    st.stop()

anios = resumen.anios(base)
c_anio, c_gen, c_pdf = st.columns([6, 2, 2], vertical_alignment="bottom")
anio = c_anio.segmented_control("Año", anios[:4], default=anios[0], key="anio_inicio") or anios[0]
try:
    d = reporte_ejecutivo.datos(base, anio=anio)
except Exception as e:
    st.error(f"No se pudo leer el Sheet: {e}")
    st.stop()

# resumen ejecutivo: se arma con un clic y se descarga al lado (el del año que se armó)
if c_gen.button("Armar resumen PDF", icon=":material/picture_as_pdf:", key="btn_ejec", width="stretch"):
    with st.spinner("Armando el resumen…"):
        try:
            st.session_state["pdf_ejec"] = (anio, reporte_ejecutivo.pdf(base, d=d))
        except Exception as e:
            st.error(f"No se pudo generar: {e}")
listo = st.session_state.get("pdf_ejec")
listo = listo if isinstance(listo, tuple) and listo[0] == anio else None
c_pdf.download_button("Descargar resumen", listo[1] if listo else b"", disabled=listo is None, type="primary",
                      file_name=f"Resumen_ejecutivo_MG_Colima_{anio}_{dt.date.today():%Y%m%d}.pdf",
                      mime="application/pdf", icon=":material/download:", key="dl_ejec", width="stretch",
                      help=None if listo else "Primero arma el resumen.")
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
ui.tarjeta(f2[0], "how_to_reg", "Registrados en Sale-U", pct(k.get("registrados", 0), k.get("saleu_total", 0)),
           f"{k.get('falta_alta', 0)} falta alta · {k.get('falta_credito', 0)} falta crédito"
           + (f" · {k['sin_revisar']} sin revisar" if k.get("sin_revisar") else ""))
ui.tarjeta(f2[1], "alarm", "Seguimientos vencidos", len(k["vencidos"]), f"{len(k['por_vencer'])} vencen en 15 días",
           ESTADO["critico"] if len(k["vencidos"]) else ESTADO["bien"])
ui.tarjeta(f2[2], "rule", "Pendientes en Revision", k["pendientes"], "por resolver",
           ESTADO["aviso"] if k["pendientes"] else ESTADO["bien"])

st.write("")
c1, c2 = st.columns([3, 2])
with c1.container(border=True, height="stretch"):
    st.markdown(":material/lightbulb: **Prioridades**")
    prio = reporte_ejecutivo.prioridades(d)
    for p in prio or ["Todo al día."]:
        st.markdown("- " + re.sub(r"</?b>", "**", p))
with c2.container(border=True, height="stretch"):
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
        st.page_link("paginas/actualizar.py", label="Ir a Actualizar", icon=":material/arrow_forward:", width="content")
    except Exception:  # fuera de la navegación (pruebas) no hay a dónde ligar
        pass

tablero.mostrar(reporte_ejecutivo.tablero(d), anio)
