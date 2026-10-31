# -*- coding: utf-8 -*-
"""Página: KPIs de capacitación Ventas y Posventa, MG Colima."""
import html

import pandas as pd
import streamlit as st

import kpis_core as kc
import ui


@st.cache_data(show_spinner="Leyendo el archivo y generando gráficas…")
def procesar_kpis(contenido: bytes):
    return kc.generar(contenido)


def tarjetas(d):
    g = d["general"]
    celdas = [f"""<div class="kpi total" style="--c:{kc.semaforo(g)}">
        <div class="n">Avance total</div><div class="v">{kc.pct(g)}</div>
        <div class="bar"><div style="width:{min(g,1)*100:.0f}%"></div></div>
        <div class="s">{d['colaboradores']} colaboradores</div></div>"""]
    for k in kc.ORDEN:
        nombre, col, _, _ = kc.MODULOS[k]
        r = d["modulos"][k]
        celdas.append(f"""<div class="kpi" style="--c:{col}">
            <div class="n">{nombre}</div><div class="v">{kc.pct(r['avance'])}</div>
            <div class="bar"><div style="width:{min(r['avance'],1)*100:.0f}%"></div></div>
            <div class="s">{r['concluidas']:.0f} de {r['total']:.0f} concluidas<br>
            {r['pendientes']:.0f} pendientes, {r['noshow']:.0f} no show</div></div>""")
    st.markdown(f'<div class="kpis">{"".join(celdas)}</div>', unsafe_allow_html=True)


# ---------------------------------------------------------------- ARCHIVO Y DESCARGAS
guardado = st.session_state.get("_archivo_kpis_archivo")
ui.encabezado("Avance de capacitación", "Ventas y Posventa, MG Colima",
              guardado[0] if guardado else "Sin archivo cargado")

c_arch, c_desc = st.columns([3, 2])
with c_arch:
    archivo = ui.archivo_persistente(
        "Plan de acción (.xlsx)", "kpis_archivo", ["xlsx", "xlsm"],
        ayuda=f"Debe tener las hojas {kc.HOJAS['VENTAS']} y {kc.HOJAS['POSVENTA']}.")

out = None
if archivo is not None:
    try:
        out = procesar_kpis(archivo[1])
    except Exception as e:  # noqa: BLE001
        ui.error("No se pudo leer el archivo", e)

if out is None:
    st.markdown("""<div class="vacio"><b>Sube el plan de acción</b><br>
    Verás el avance de DOT, Cápsulas, LOL y Presencial y podrás descargar PDF, imágenes y Excel.</div>""",
                unsafe_allow_html=True)
    st.stop()

areas = out["areas"]
n_alertas = sum(len(d["alertas"]) for d in areas.values())

with c_desc.container(border=True):
    st.markdown(":material/download: **Descargas**")
    fila = ui.barra(alinear="left")
    fila.download_button("PDF completo", out["pdf"], file_name=kc.SALIDA, mime="application/pdf",
                         icon=":material/picture_as_pdf:", type="primary")
    fila.download_button("Excel", out["excel"], file_name="KPIs_Capacitacion_detalle.xlsx",
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         icon=":material/table_view:")
    fila.download_button("ZIP", out["zip"], file_name="KPIs_Capacitacion.zip", mime="application/zip",
                         icon=":material/folder_zip:")
    for area, d in areas.items():
        fila.download_button(f"Imagen {area.title()}", out["imagenes"][d["png"]], file_name=d["png"],
                             mime="image/png", icon=":material/image:")
    st.caption(f"Semáforo: verde ≥ {kc.META_VERDE:.0%}, ámbar ≥ {kc.META_AMBAR:.0%}, rojo menor.")

etiqueta_rev = f"Revisión de datos ({n_alertas})" if n_alertas else "Revisión de datos"
tabs = st.tabs([a.title() for a in areas] + ["Comparativo", etiqueta_rev])

# ---------------------------------------------------------------- ÁREAS
for (area, d), tab in zip(areas.items(), tabs):
    with tab:
        if d["alertas"]:
            n = len(d["alertas"])
            st.warning(f"{n} diferencia(s) con el Excel. Ver *Revisión de datos*.", icon=":material/warning:")
        tarjetas(d)

        with st.container(border=True):
            st.image(out["imagenes"][d["png"]], width="stretch")

        ui.seccion("Detalle por colaborador")
        df = d["detalle"]
        c1, c2 = st.columns([1, 2])
        letras = sorted(df["Letra"].unique(), key=lambda l: (l != "P", l))
        sel = c1.multiselect("Letra", letras, default=letras, key=f"l_{area}")
        buscar = c2.text_input("Buscar colaborador", key=f"b_{area}", placeholder="Nombre o apellido")
        vista = df[df["Letra"].isin(sel)]
        if buscar:
            vista = vista[vista["Colaborador"].str.contains(buscar, case=False, na=False)]
        modulos = [kc.MODULOS[k][0] for k in kc.ORDEN]
        tabla = vista[["Colaborador", "Letra"] + modulos + ["Promedio"]].copy()
        tabla["Colaborador"] = tabla["Colaborador"].str.title()
        tabla[modulos + ["Promedio"]] = tabla[modulos + ["Promedio"]] * 100
        cfg = {m: st.column_config.ProgressColumn(m, format="%.0f%%", min_value=0, max_value=100)
               for m in modulos + ["Promedio"]}
        cfg["Colaborador"] = st.column_config.TextColumn("Colaborador", width="medium")
        cfg["Letra"] = st.column_config.TextColumn("Letra", width="small")
        st.dataframe(tabla.sort_values("Promedio"), column_config=cfg, hide_index=True,
                     width="stretch", height=min(36 * (len(tabla) + 1) + 4, 620))
        st.caption(f"{len(tabla)} de {len(df)} colaboradores, de menor a mayor avance. El detalle está en el Excel.")

# ---------------------------------------------------------------- COMPARATIVO
with tabs[len(areas)]:
    import altair as alt

    nombres_area = [a.title() for a in areas]
    etiquetas = [kc.MODULOS[k][0] for k in kc.ORDEN] + ["Total"]
    filas = []
    for area, d in areas.items():
        for k in kc.ORDEN:
            filas.append({"Módulo": kc.MODULOS[k][0], "Área": area.title(),
                          "Avance": d["modulos"][k]["avance"]})
        filas.append({"Módulo": "Total", "Área": area.title(), "Avance": d["general"]})
    comp = pd.DataFrame(filas)

    ui.seccion("Avance por módulo")
    base = alt.Chart(comp).encode(
        x=alt.X("Módulo:N", sort=etiquetas, title=None, axis=alt.Axis(labelAngle=0, labelFontSize=13)),
        xOffset=alt.XOffset("Área:N", sort=nombres_area),
        y=alt.Y("Avance:Q", title=None, scale=alt.Scale(domain=[0, 1.08]),
                axis=alt.Axis(format="%", tickCount=5)),
        color=alt.Color("Área:N", sort=nombres_area, title=None,
                        scale=alt.Scale(domain=nombres_area, range=["#3987e5", "#86b6ef"]),
                        legend=alt.Legend(orient="top")),
        tooltip=["Área", "Módulo", alt.Tooltip("Avance:Q", format=".0%")],
    )
    barras = base.mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
    textos = base.mark_text(dy=-8, fontSize=12, fontWeight="bold", color="#27272a").encode(
        text=alt.Text("Avance:Q", format=".0%"))
    grafica = (
        (barras + textos)
        .properties(height=380, background="white",
                    padding={"left": 16, "right": 16, "top": 16, "bottom": 16})
        .configure_view(fill="white", stroke=None)
        .configure_axis(labelColor="#3f3f46", titleColor="#3f3f46", gridColor="#e4e4e7",
                        domainColor="#d4d4d8", tickColor="#d4d4d8")
        .configure_legend(labelColor="#27272a", titleColor="#27272a")
    )
    st.altair_chart(grafica, width="stretch", theme=None)

    tabla = comp.pivot(index="Módulo", columns="Área", values="Avance").reindex(etiquetas)[nombres_area]
    if len(nombres_area) == 2:
        a, b = nombres_area
        tabla[f"{a} vs {b} (pts)"] = (tabla[a] - tabla[b]) * 100
    fmt = {c: "{:.0%}" for c in nombres_area}
    fmt.update({c: "{:+.0f}" for c in tabla.columns if c not in nombres_area})
    st.dataframe(tabla.style.format(fmt), width="stretch")

# ---------------------------------------------------------------- REVISIÓN
with tabs[-1]:
    st.write("Diferencias entre las fórmulas del Excel y el recálculo por colaborador. El reporte usa el recálculo.")
    for area, d in areas.items():
        ui.seccion(area.title())
        if not d["alertas"]:
            st.success("Sin diferencias.", icon=":material/check_circle:")
        for a in d["alertas"]:
            st.markdown(f'<div class="alerta">{html.escape(a)}</div>', unsafe_allow_html=True)
