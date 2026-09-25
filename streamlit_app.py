# -*- coding: utf-8 -*-
"""App de Streamlit: KPIs de capacitación Ventas y Posventa, MG Colima."""
import html

import pandas as pd
import streamlit as st

import kpis_core as kc

st.set_page_config(page_title="KPIs de capacitación | MG Colima", page_icon="📊", layout="wide")

# ---------------------------------------------------------------- ESTILO
st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700;800&display=swap');
.hero, .kpi, .alerta, .vacio {{ font-family: 'Archivo', sans-serif; }}
.block-container {{ padding-top: 1.6rem; max-width: 1280px; }}

.hero {{ background: {kc.AZUL_OSCURO}; color: #fff; border-radius: 14px; padding: 22px 28px;
        display: flex; justify-content: space-between; align-items: flex-end; gap: 16px;
        flex-wrap: wrap; margin-bottom: 18px; }}
.hero h1 {{ font-size: 1.9rem; font-weight: 800; margin: 0; color: #fff; line-height: 1.1; }}
.hero p {{ margin: 6px 0 0; color: #C9D6EA; font-size: .95rem; }}
.hero .file {{ color: #C9D6EA; font-size: .85rem; text-align: right; }}

.kpis {{ display: grid; grid-template-columns: 1.25fr repeat(4, 1fr); gap: 12px; margin: 4px 0 18px; }}
@media (max-width: 900px) {{ .kpis {{ grid-template-columns: repeat(2, 1fr); }} }}
.kpi {{ background: #fff; border: 1px solid #E3E8EF; border-radius: 12px; padding: 14px 16px;
       border-top: 5px solid var(--c); }}
.kpi .n {{ font-size: .9rem; color: {kc.TEXTO_2}; font-weight: 600; }}
.kpi .v {{ font-size: 2rem; font-weight: 800; color: {kc.TEXTO}; line-height: 1.15; }}
.kpi .s {{ font-size: .8rem; color: {kc.TEXTO_2}; }}
.kpi .bar {{ height: 6px; background: #E7EBF1; border-radius: 4px; margin: 8px 0 6px; overflow: hidden; }}
.kpi .bar > div {{ height: 100%; background: var(--c); border-radius: 4px; }}
.kpi.total {{ background: {kc.AZUL_OSCURO}; border-color: {kc.AZUL_OSCURO}; }}
.kpi.total .n, .kpi.total .s {{ color: #C9D6EA; }}
.kpi.total .v {{ color: #fff; font-size: 2.4rem; }}
.kpi.total .bar {{ background: rgba(255,255,255,.18); }}

.alerta {{ border-left: 4px solid {kc.AMBAR}; background: #FFF9EC; padding: 10px 14px;
          border-radius: 6px; margin-bottom: 8px; color: {kc.TEXTO}; font-size: .93rem; }}
.vacio {{ border: 2px dashed #CBD4E1; border-radius: 14px; padding: 36px; text-align: center;
         color: {kc.TEXTO_2}; }}
.vacio b {{ color: {kc.TEXTO}; font-size: 1.1rem; }}
</style>
""", unsafe_allow_html=True)


@st.cache_data(show_spinner="Leyendo el archivo y generando gráficas…")
def procesar(contenido: bytes):
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


# ---------------------------------------------------------------- BARRA LATERAL
with st.sidebar:
    st.subheader("Archivo")
    archivo = st.file_uploader("Plan de acción (.xlsx)", type=["xlsx", "xlsm"],
                               help=f"Debe tener las hojas {kc.HOJAS['VENTAS']} y {kc.HOJAS['POSVENTA']}.")

out = None
if archivo is not None:
    try:
        out = procesar(archivo.getvalue())
    except Exception as e:  # noqa: BLE001
        st.sidebar.error(f"No se pudo leer el archivo. {e}")

sub = archivo.name if archivo is not None else "Sin archivo cargado"
st.markdown(f"""<div class="hero"><div><h1>Avance de capacitación</h1>
<p>Ventas y Posventa, MG Colima</p></div><div class="file">{html.escape(sub)}</div></div>""",
            unsafe_allow_html=True)

if out is None:
    st.markdown("""<div class="vacio"><b>Sube el archivo del plan de acción en la barra lateral</b><br>
    Se leerán las tablas de DOT, Cápsulas, LOL y Presencial de cada área y podrás descargar
    el PDF, la imagen del resumen y el detalle en Excel.</div>""", unsafe_allow_html=True)
    st.stop()

areas = out["areas"]
n_alertas = sum(len(d["alertas"]) for d in areas.values())

with st.sidebar:
    st.subheader("Descargas")
    st.download_button("PDF completo", out["pdf"], file_name=kc.SALIDA, mime="application/pdf",
                       icon=":material/picture_as_pdf:", type="primary", width="stretch")
    for area, d in areas.items():
        st.download_button(f"Imagen {area.title()}", out["imagenes"][d["png"]], file_name=d["png"],
                           mime="image/png", icon=":material/image:", width="stretch")
    st.download_button("Detalle en Excel", out["excel"], file_name="KPIs_Capacitacion_detalle.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       icon=":material/table_view:", width="stretch")
    st.download_button("Todo en un ZIP", out["zip"], file_name="KPIs_Capacitacion.zip",
                       mime="application/zip", icon=":material/folder_zip:", width="stretch")
    st.divider()
    st.caption(f"Semáforo: verde ≥ {kc.META_VERDE:.0%}, ámbar ≥ {kc.META_AMBAR:.0%}, rojo menor. "
               "Los totales se recalculan desde las filas de cada colaborador.")

etiqueta_rev = f"Revisión de datos ({n_alertas})" if n_alertas else "Revisión de datos"
tabs = st.tabs([a.title() for a in areas] + ["Comparativo", etiqueta_rev])

# ---------------------------------------------------------------- ÁREAS
for (area, d), tab in zip(areas.items(), tabs):
    with tab:
        if d["alertas"]:
            n = len(d["alertas"])
            txt = "1 diferencia" if n == 1 else f"{n} diferencias"
            verbo = "Se encontró" if n == 1 else "Se encontraron"
            st.warning(f"{verbo} {txt} entre el Excel y los datos recalculados. "
                       "Consulta la pestaña de revisión de datos.", icon=":material/warning:")
        tarjetas(d)

        st.image(out["imagenes"][d["png"]], width="stretch")

        st.subheader("Detalle por colaborador")
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
        st.caption(f"{len(tabla)} de {len(df)} colaboradores, de menor a mayor avance promedio. "
                   "El detalle de concluidas y total por módulo está en el Excel de descarga.")

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

    st.subheader("Avance por módulo")
    base = alt.Chart(comp).encode(
        x=alt.X("Módulo:N", sort=etiquetas, title=None, axis=alt.Axis(labelAngle=0, labelFontSize=13)),
        xOffset=alt.XOffset("Área:N", sort=nombres_area),
        y=alt.Y("Avance:Q", title=None, scale=alt.Scale(domain=[0, 1.08]),
                axis=alt.Axis(format="%", tickCount=5)),
        color=alt.Color("Área:N", sort=nombres_area, title=None,
                        scale=alt.Scale(domain=nombres_area, range=[kc.AZUL_OSCURO, "#7FA7D9"]),
                        legend=alt.Legend(orient="top")),
        tooltip=["Área", "Módulo", alt.Tooltip("Avance:Q", format=".0%")],
    )
    barras = base.mark_bar(cornerRadiusTopLeft=4, cornerRadiusTopRight=4)
    textos = base.mark_text(dy=-8, fontSize=12, fontWeight="bold", color=kc.TEXTO).encode(
        text=alt.Text("Avance:Q", format=".0%"))
    st.altair_chart((barras + textos).properties(height=380), width="stretch")

    tabla = comp.pivot(index="Módulo", columns="Área", values="Avance").reindex(etiquetas)[nombres_area]
    if len(nombres_area) == 2:
        a, b = nombres_area
        tabla[f"{a} vs {b} (pts)"] = (tabla[a] - tabla[b]) * 100
    fmt = {c: "{:.0%}" for c in nombres_area}
    fmt.update({c: "{:+.0f}" for c in tabla.columns if c not in nombres_area})
    st.dataframe(tabla.style.format(fmt), width="stretch")

# ---------------------------------------------------------------- REVISIÓN
with tabs[-1]:
    st.write("Diferencias encontradas entre las fórmulas del Excel y el recálculo desde las filas, "
             "además de datos inconsistentes por colaborador. El reporte usa siempre los valores recalculados.")
    for area, d in areas.items():
        st.subheader(area.title())
        if not d["alertas"]:
            st.success("Sin diferencias.", icon=":material/check_circle:")
        for a in d["alertas"]:
            st.markdown(f'<div class="alerta">{html.escape(a)}</div>', unsafe_allow_html=True)
