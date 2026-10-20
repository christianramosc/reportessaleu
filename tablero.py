# -*- coding: utf-8 -*-
"""
Tablero del Inicio: gráficas interactivas (Altair) con los datos de reporte_ejecutivo.tablero(d). Sin nombres de
clientes ni de asesores: los seguimientos se identifican por su ID de «MG Colima Rechazos».

  · Atraso de los vencidos por rango (clic en un rango filtra la gráfica de al lado) y días de atraso por ID.
  · Agenda: vencidos y lo que toca cada semana.
  · Embudo de contacto, resultado del contacto y canal.
  · Seguimiento al día, registro en Sale-U y rechazos por financiera (donas).
"""
import altair as alt
import pandas as pd
import streamlit as st

from cuadres.reporte_ejecutivo import ESTADO, NAVY, NAVY_2, SERIE, RANGOS, COLOR_RANGO

TINTA, TINTA_2 = "#1B2A41", "#5B6B82"
FUENTE = "Archivo, sans-serif"


def _estilo(c, alto):
    return (c.properties(height=alto)
            .configure_view(stroke=None)
            .configure_axis(labelColor=TINTA_2, titleColor=TINTA_2, gridColor="#E3E8EF", domainColor="#E3E8EF",
                            labelFont=FUENTE, titleFont=FUENTE, labelFontSize=11, titleFontSize=11, tickColor="#E3E8EF")
            .configure_legend(labelColor=TINTA_2, titleColor=TINTA_2, labelFont=FUENTE, titleFont=FUENTE,
                              labelFontSize=11, orient="bottom", symbolType="circle")
            .configure_text(font=FUENTE))


def _caja(titulo, subtitulo=""):
    caja = st.container(border=True)
    caja.markdown(f"**{titulo}**" + (f"  \n<span style='color:{TINTA_2};font-size:.85rem'>{subtitulo}</span>"
                                     if subtitulo else ""), unsafe_allow_html=True)
    return caja


def _vacio(caja, texto):
    caja.markdown(f"<div style='color:{ESTADO['bien']};padding:24px 0;text-align:center'>{texto}</div>",
                  unsafe_allow_html=True)


def _dona(df, campo, alto=230):
    df = df[df["Clientes"] > 0].copy()
    total = int(df["Clientes"].sum())
    df["Porcentaje"] = df["Clientes"] / total if total else 0
    hover = alt.selection_point(fields=[campo], on="pointerover", clear="pointerout")
    escala = alt.Scale(domain=list(df[campo]), range=list(df["Color"]))
    base = alt.Chart(df)
    arco = base.mark_arc(innerRadius=62, outerRadius=95, stroke="white", strokeWidth=2, cursor="pointer").encode(
        theta=alt.Theta("Clientes:Q", stack=True),
        color=alt.Color(f"{campo}:N", scale=escala, legend=alt.Legend(title=None, columns=2)),
        opacity=alt.condition(hover, alt.value(1), alt.value(0.45)),
        tooltip=[alt.Tooltip(f"{campo}:N"), alt.Tooltip("Clientes:Q"), alt.Tooltip("Porcentaje:Q", format=".0%")],
    ).add_params(hover)
    centro = alt.Chart(pd.DataFrame({"t": [str(total)]})).mark_text(fontSize=26, fontWeight="bold", color=TINTA,
                                                                     dy=-6).encode(text="t:N")
    pie = alt.Chart(pd.DataFrame({"t": ["clientes"]})).mark_text(fontSize=11, color=TINTA_2, dy=16).encode(text="t:N")
    return _estilo(alt.layer(arco, centro, pie), alto)


def mostrar(t, anio):
    st.markdown("#### :material/monitoring: Tablero")

    # ------------------------------------------------------------ vencidos (filtro cruzado por rango)
    venc, antig = t["vencidos"], t["antiguedad"]
    caja = _caja("Seguimientos vencidos",
                 f"{len(venc)} sin primer contacto después de los 180 días · haz clic en un rango para filtrar")
    if venc.empty:
        _vacio(caja, ":material/check_circle: Sin seguimientos vencidos")
    else:
        escala = alt.Scale(domain=[r for _, _, r in RANGOS], range=COLOR_RANGO)
        rango = alt.selection_point(fields=["Atraso"], name="rango")
        i, d = caja.columns([2, 3])
        izq = (alt.Chart(antig).mark_bar(cornerRadiusEnd=4, cursor="pointer").encode(
            y=alt.Y("Atraso:N", sort=[r for _, _, r in RANGOS], title=None),
            x=alt.X("Seguimientos:Q", title="Seguimientos", axis=alt.Axis(tickMinStep=1)),
            color=alt.Color("Atraso:N", scale=escala, legend=None),
            opacity=alt.condition(rango, alt.value(1), alt.value(0.35)),
            tooltip=["Atraso:N", "Seguimientos:Q"]).add_params(rango))
        izq = alt.layer(izq, izq.mark_text(align="left", dx=4).encode(text="Seguimientos:Q", opacity=alt.value(1),
                                                                     color=alt.value(TINTA)))
        ev = i.altair_chart(_estilo(izq, 240), width="stretch", on_select="rerun", key="sel_rango")
        elegidos = [x.get("Atraso") for x in (ev.selection.get("rango") or [])] if ev else []
        v = venc[venc["Atraso"].isin(elegidos)] if elegidos else venc
        v = v.head(25)
        der = alt.Chart(v).mark_bar(cornerRadiusEnd=4).encode(
            y=alt.Y("Seguimiento:N", sort=list(v["Seguimiento"]), title=None),
            x=alt.X("Días:Q", title="Días de atraso"),
            color=alt.Color("Atraso:N", scale=escala, legend=None),
            tooltip=["Seguimiento:N", "Días:Q", "Atraso:N", "Financiera:N", "Tocaba el:N"])
        txt = der.mark_text(align="left", dx=4, color=TINTA).encode(text="Días:Q", color=alt.value(TINTA))
        d.altair_chart(_estilo(alt.layer(der, txt), max(240, 24 * len(v))), width="stretch")
        if elegidos:
            d.caption(f"Filtrado: {', '.join(elegidos)}. Vuelve a hacer clic en la barra para quitar el filtro.")
        elif len(venc) > 25:
            d.caption(f"Se muestran los 25 más atrasados de {len(venc)}.")

    # ------------------------------------------------------------ agenda
    c1, c2 = st.columns([3, 2])
    with c1:
        ag = t["agenda"]
        caja = _caja("Agenda de seguimientos", "Vencidos y lo que toca cada semana (de lunes a domingo)")
        orden = list(ag["Semana"])
        graf = alt.Chart(ag).mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4).encode(
            x=alt.X("Semana:N", sort=orden, title=None, axis=alt.Axis(labelAngle=0)),
            y=alt.Y("Seguimientos:Q", title=None, axis=alt.Axis(tickMinStep=1)),
            color=alt.Color("Tipo:N", scale=alt.Scale(domain=["Vencido", "Esta semana", "Próximo"],
                                                      range=[ESTADO["critico"], NAVY_2, SERIE[0]]),
                            legend=alt.Legend(title=None)),
            tooltip=["Semana:N", "Seguimientos:Q"])
        txt = alt.Chart(ag[ag["Seguimientos"] > 0]).mark_text(dy=-8, color=TINTA, fontWeight="bold").encode(
            x=alt.X("Semana:N", sort=orden), y="Seguimientos:Q", text="Seguimientos:Q")
        caja.altair_chart(_estilo(alt.layer(graf, txt), 260), width="stretch")
    with c2:
        caja = _caja("Seguimiento al día", "De los rechazos que ya cumplieron 180 días")
        if t["al_dia"]["Clientes"].sum():
            caja.altair_chart(_dona(t["al_dia"], "Estado"), width="stretch")
        else:
            _vacio(caja, "Todavía no le toca a ninguno")

    # ------------------------------------------------------------ contacto
    c1, c2, c3 = st.columns([5, 4, 3])
    with c1:
        e = t["embudo"].copy()
        e["Porcentaje"] = e["Clientes"] / e["Clientes"].iloc[0] if e["Clientes"].iloc[0] else 0
        tope = max(int(e["Clientes"].max()), 1)
        e["ini"], e["fin"] = (tope - e["Clientes"]) / 2, (tope + e["Clientes"]) / 2
        caja = _caja("Embudo de contacto", f"Rechazos de {anio}")
        pasos = list(e["Paso"])
        barras = alt.Chart(e).mark_bar(cornerRadius=4, height=34).encode(
            y=alt.Y("Paso:N", sort=pasos, title=None, axis=alt.Axis(labelLimit=200)),
            x=alt.X("ini:Q", axis=None, scale=alt.Scale(domain=[0, tope])), x2="fin:Q",
            color=alt.Color("Paso:N", sort=pasos, scale=alt.Scale(range=[NAVY, NAVY_2, SERIE[0], ESTADO["bien"]]),
                            legend=None),
            tooltip=["Paso:N", "Clientes:Q", alt.Tooltip("Porcentaje:Q", format=".0%")])
        txt = alt.Chart(e).transform_calculate(
            et="datum.Clientes + '  (' + format(datum.Porcentaje, '.0%') + ')'", x=f"{tope} / 2").mark_text(
            color=TINTA, fontWeight="bold", dy=-24, fontSize=11).encode(
            y=alt.Y("Paso:N", sort=pasos), x=alt.X("x:Q", scale=alt.Scale(domain=[0, tope])), text="et:N")
        caja.altair_chart(_estilo(alt.layer(barras, txt), 260), width="stretch")
    with c2:
        caja = _caja("Resultado del contacto", "Primer contacto: ¿contestó?")
        if t["contacto"]["Clientes"].sum():
            caja.altair_chart(_dona(t["contacto"], "Resultado"), width="stretch")
        else:
            _vacio(caja, "Sin contactos capturados")
    with c3:
        caja = _caja("Canal del contacto", "Primer contacto")
        ca = t["canal"]
        if ca.empty:
            _vacio(caja, "Sin contactos capturados")
        else:
            graf = alt.Chart(ca).mark_bar(cornerRadiusEnd=4, color=NAVY_2).encode(
                y=alt.Y("Canal:N", sort="-x", title=None), x=alt.X("Clientes:Q", title=None),
                tooltip=["Canal:N", "Clientes:Q"])
            txt = graf.mark_text(align="left", dx=4, color=TINTA).encode(text="Clientes:Q")
            caja.altair_chart(_estilo(alt.layer(graf, txt), 230), width="stretch")

    # ------------------------------------------------------------ Sale-U y financieras
    c1, c2 = st.columns(2)
    with c1:
        caja = _caja("Registro en Sale-U", "Rechazados y aprobados del año")
        if t["saleu"]["Clientes"].sum():
            caja.altair_chart(_dona(t["saleu"], "Estado"), width="stretch")
        else:
            _vacio(caja, "Sin datos de Sale-U")
    with c2:
        caja = _caja("Rechazos por financiera", f"{int(t['financiera']['Rechazos'].sum())} en {anio}")
        fi = t["financiera"].rename(columns={"Rechazos": "Clientes"})
        if fi.empty:
            _vacio(caja, "Sin rechazos en el año")
        else:
            fi["Color"] = fi["Financiera"].map(dict(zip(["CETELEM", "BBVA", "Otras"], SERIE))).fillna("#8a8984")
            caja.altair_chart(_dona(fi, "Financiera"), width="stretch")
