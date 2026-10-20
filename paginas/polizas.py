# -*- coding: utf-8 -*-
"""
Página: Pólizas

Seguimiento de pólizas de seguro de los créditos: a quién llamar hoy, oportunidades por mes y registro del contacto.
Se suben los PDFs / ZIP de las pólizas (Crédito Click) o el Excel Polizas_extraidas.xlsx; los datos de clientes solo
viven en la sesión (nunca se guardan en el repo ni en el servidor). El seguimiento se descarga en el mismo Excel.
"""
import datetime as dt

import altair as alt
import pandas as pd
import streamlit as st

import ui
from cuadres import polizas_seguimiento as ps
from cuadres.salidas import MESES_CORTOS

MESES = ["enero", "febrero", "marzo", "abril", "mayo", "junio", "julio", "agosto", "septiembre", "octubre",
         "noviembre", "diciembre"]
hoy = dt.date.today()
ui.hero("Pólizas de seguro", "A quién llamar para póliza externa, pago de contado o fin de crédito.",
        f"Corte {hoy:%d/%m/%Y}")

# ----------------------------------------------------------------------------- carga
with st.container(border=True):
    c1, c2 = st.columns([3, 2])
    pdfs = c1.file_uploader("Pólizas (PDF o ZIP)", type=["pdf", "zip"], accept_multiple_files=True, key="pol_pdfs",
                            help="Los PDF que se descargan de Crédito Click (columna Impresión). Puedes subir varios "
                                 "o un ZIP. Si solo vas a revisar el seguimiento, basta con el Excel.")
    xls = c2.file_uploader("Excel anterior (opcional)", type=["xlsx"], key="pol_xlsx",
                           help="Polizas_extraidas.xlsx: conserva lo que ya capturaste (Contactado, Resultado, Notas) "
                                "y, si también subes PDFs nuevos, los junta con los clientes que ya tenías.")
    if st.button("Cargar", type="primary", disabled=not (pdfs or xls), icon=":material/upload:"):
        with st.spinner("Leyendo pólizas…" if pdfs else "Leyendo el Excel…"):
            try:
                base, excel, errores = ps.cargar([(f.name, f.getvalue()) for f in pdfs or []],
                                                 xls.getvalue() if xls else None)
                st.session_state.update(pol_base=base, pol_excel=excel, pol_err=errores,
                                        pol_origen=f"{len(pdfs)} archivo(s) PDF/ZIP" if pdfs else xls.name)
                st.session_state.pop("pol_editor", None)
            except Exception as e:
                st.error(f"No se pudo cargar: {e}")

base = st.session_state.get("pol_base")
if base is None:
    st.markdown('<div class="vacio"><b>Sube las pólizas o tu Excel de pólizas</b><br>'
                'La primera vez sube los PDF o el ZIP; después basta con el Excel que descargas de aquí.</div>',
                unsafe_allow_html=True)
    st.stop()

if st.session_state.get("pol_err"):
    with st.expander(f":material/warning: {len(st.session_state['pol_err'])} PDF no se leyeron completos"):
        st.dataframe(pd.DataFrame(st.session_state["pol_err"], columns=["Archivo", "Motivo"]), hide_index=True)

with st.expander("Reglas", icon=":material/tune:"):
    r1, r2, r3 = st.columns(3)
    reglas = dict(
        externa=r1.number_input("Último día póliza externa (días antes del aniversario)", 1, 120, ps.REGLAS["externa"]),
        contado=r2.number_input("Abre ventana de contado (días antes)", 1, 120, ps.REGLAS["contado"]),
        aviso=r3.number_input("Llamar N días antes del límite de externa", 0, 60, ps.REGLAS["aviso"]))
    st.caption("Póliza externa hasta 45 días antes del aniversario (después se cierra) · contado desde 30 días antes "
               "hasta el día anterior · al terminar el crédito el auto es del cliente y puede contratar el seguro que "
               "quiera.")

df = ps.calcular(base, hoy, reglas)
proy = ps.proyeccion(base, hoy, 12, reglas)
n = lambda e: int((df["ESTATUS"] == e).sum())
contactados = df["CONTACTADO"].astype(str).str.lower().isin(["sí", "si"]).sum()
cierres = df["RESULTADO"].isin(ps.CIERRES).sum()

k = st.columns(5)
ui.tarjeta(k[0], "call", "Llamar ya", n(ps.LLAMAR_EXT) + n(ps.FIN_LLAMAR),
           f"{n(ps.LLAMAR_EXT)} póliza externa · {n(ps.FIN_LLAMAR)} fin de crédito", "#9C0006")
ui.tarjeta(k[1], "payments", "Contado abierto", n(ps.CONTADO), "Que paguen antes del día anterior al aniversario",
           "#9C5700")
ui.tarjeta(k[2], "lock_clock", "Externa cerrada", n(ps.EXT_CERRADA), "Esperan su ventana de contado", "#C55A11")
mes_sig = proy[proy["MES"] == dt.date(hoy.year, hoy.month, 1)] if not proy.empty else proy
ui.tarjeta(k[3], "event_upcoming", f"Este mes ({MESES[hoy.month - 1]})", len(mes_sig),
           f"Prima anual en juego ${mes_sig['PRIMA ANUAL ACTUAL'].sum():,.0f}" if len(mes_sig) else "Sin oportunidades")
ui.tarjeta(k[4], "verified", "Contactados", int(contactados),
           f"{cierres} cerrados ({cierres / contactados:.0%})" if contactados else "Aún sin registrar")
st.caption(f"{len(df)} créditos · fuente: {st.session_state.get('pol_origen', '')}")

COLS_HOY = ["ESTATUS", "CLIENTE", "TELÉFONO", "CIERRA EN", "PRÓXIMA FECHA", "OPORTUNIDAD", "LLAMAR DESDE",
            "ÚLTIMO DÍA EXTERNA", "ABRE CONTADO", "ÚLTIMO DÍA CONTADO", "QUÉ HACER", "ASEGURADORA",
            "PRIMA ANUAL ACTUAL", "CORREO", "VEHÍCULO", "FOLIO CRÉDITO", "MUNICIPIO"]
CONF = {c: st.column_config.DateColumn(c, format="DD/MM/YYYY") for c in
        ["PRÓXIMA FECHA", "LLAMAR DESDE", "ÚLTIMO DÍA EXTERNA", "ABRE CONTADO", "ÚLTIMO DÍA CONTADO", "FECHA",
         "FECHA CONTACTO", "FIN DEL CRÉDITO"]}
CONF["PRIMA ANUAL ACTUAL"] = st.column_config.NumberColumn("PRIMA ANUAL", format="dollar")


def con_seguimiento(base, editado):
    """Pasa lo capturado en la tabla de Seguimiento a la base (si marcó «Sí» sin fecha, se toma hoy)."""
    b = base.copy()
    if editado is None or editado.empty:
        return b
    for _, r in editado.iterrows():
        i = b.index[b["SERIE"] == r["SERIE"]]
        fecha = r["FECHA CONTACTO"]
        sin_fecha = fecha is None or pd.isna(fecha)
        if str(r["CONTACTADO"]).strip().lower() in ("sí", "si") and sin_fecha:
            fecha, sin_fecha = pd.Timestamp(hoy), False
        b.loc[i, "CONTACTADO"] = r["CONTACTADO"] if isinstance(r["CONTACTADO"], str) else ""
        b.loc[i, "FECHA CONTACTO"] = None if sin_fecha else pd.Timestamp(fecha).date()
        b.loc[i, "RESULTADO"] = r["RESULTADO"] if isinstance(r["RESULTADO"], str) else ""
        b.loc[i, "NOTAS"] = r["NOTAS"] if isinstance(r["NOTAS"], str) else ""
    return b


def colorear(v):
    if v in ps.ESTATUS:
        _, fondo, letra = ps.ESTATUS[v]
        return f"background-color: #{fondo}; color: #{letra}; font-weight: 600"
    return ""


def ver(d, cols, alto=None):
    if d.empty:
        st.info("Sin clientes en esta vista.")
        return
    cols = [c for c in cols if c in d.columns]
    estilo = d[cols].style.map(colorear, subset=[c for c in ["ESTATUS"] if c in cols])
    st.dataframe(estilo, hide_index=True, width="stretch", column_config=CONF, **({"height": alto} if alto else {}))


t_hoy, t_mes, t_todos, t_seg = st.tabs([":material/today: Hoy", ":material/calendar_month: Por mes",
                                        ":material/groups: Todos", ":material/edit_note: Seguimiento"])

with t_hoy:
    urg = df[df["ESTATUS"].isin(ps.URGENTES)]
    st.markdown(f"**Para contactar ya ({len(urg)})** · ordenados por lo que vence primero")
    ver(urg.sort_values("DÍAS QUE QUEDAN", na_position="last"), COLS_HOY)
    sem = df[(df["ESTATUS"].isin([ps.PROXIMO, ps.FIN_PROX])) &
             (df["LLAMAR DESDE"] <= hoy + dt.timedelta(days=7))]
    st.markdown(f"**Empiezan a llamarse en los próximos 7 días ({len(sem)})**")
    ver(sem, COLS_HOY)
    esp = df[df["ESTATUS"] == ps.EXT_CERRADA]
    st.markdown(f"**Externa cerrada, esperando contado ({len(esp)})**")
    ver(esp, COLS_HOY)

with t_mes:
    if proy.empty:
        st.info("Sin oportunidades en los próximos 12 meses.")
    else:
        g = proy.assign(Mes=proy["MES"].map(lambda d: f"{MESES_CORTOS[d.month - 1]} {d.year}"),
                        orden=proy["MES"]).groupby(["orden", "Mes", "TIPO"]).size().reset_index(name="Oportunidades")
        graf = alt.Chart(g).mark_bar().encode(
            x=alt.X("Mes:N", sort=alt.SortField("orden"), title=None),
            y=alt.Y("Oportunidades:Q", title="Oportunidades"),
            color=alt.Color("TIPO:N", title=None, scale=alt.Scale(domain=["Aniversario", "Fin de crédito"],
                                                                   range=["#1D4A86", "#8E5BB5"])),
            tooltip=["Mes", "TIPO", "Oportunidades"]).properties(height=240)
        st.altair_chart(graf, width="stretch")
        st.caption("Mes en que toca empezar a llamar (7 días antes del límite de póliza externa). "
                   "Las que ya están abiertas cuentan en el mes actual.")
        meses = sorted(proy["MES"].unique())
        etiqueta = lambda d: f"{MESES[d.month - 1].capitalize()} {d.year}"
        m = st.selectbox("Mes", meses, format_func=etiqueta, key="pol_mes")
        del_mes = proy[proy["MES"] == m]
        a, b, c = st.columns(3)
        a.metric("Oportunidades", len(del_mes))
        b.metric("Fin de crédito", int((del_mes["TIPO"] == "Fin de crédito").sum()))
        c.metric("Prima anual en juego", f"${del_mes['PRIMA ANUAL ACTUAL'].sum():,.0f}")
        ver(del_mes, ["LLAMAR DESDE", "FECHA", "TIPO", "OPORTUNIDAD", "CLIENTE", "TELÉFONO", "CORREO", "ASEGURADORA",
                      "PRIMA ANUAL ACTUAL", "VEHÍCULO"])
        st.session_state["pol_pdf_mes"] = (del_mes, etiqueta(m))

with t_todos:
    f1, f2, f3 = st.columns(3)
    est = f1.multiselect("Estatus", list(ps.ESTATUS), placeholder="Todos", key="pol_f_est")
    ase = f2.multiselect("Aseguradora", sorted(df["ASEGURADORA"].dropna().unique()), placeholder="Todas",
                         key="pol_f_ase")
    txt = f3.text_input("Buscar cliente, serie o folio", key="pol_f_txt")
    vista = df
    if est:
        vista = vista[vista["ESTATUS"].isin(est)]
    if ase:
        vista = vista[vista["ASEGURADORA"].isin(ase)]
    if txt:
        t = txt.strip().upper()
        vista = vista[vista["CLIENTE"].str.upper().str.contains(t, regex=False) |
                      vista["SERIE"].str.upper().str.contains(t, regex=False) |
                      vista["FOLIO CRÉDITO"].astype(str).str.contains(t, regex=False)]
    st.caption(f"{len(vista)} de {len(df)} créditos")
    ver(vista, COLS_HOY + ["FIN DEL CRÉDITO", "DIRECCIÓN", "SERIE"], alto=520)

with t_seg:
    st.caption("Registra aquí el contacto. Para no perderlo, **descarga el Excel** al terminar (la app no guarda "
               "datos de clientes) y la próxima vez súbelo en «Excel anterior».")
    solo = st.toggle("Solo los que hay que contactar ahora", value=True, key="pol_seg_solo")
    ed = df if not solo else df[df["ESTATUS"].isin(ps.URGENTES + [ps.EXT_CERRADA, ps.CONTACTADO])]
    cols_ed = ["ESTATUS", "CLIENTE", "TELÉFONO", "PRÓXIMA FECHA", "CIERRA EN"] + ps.COLS_SEG + ["SERIE"]
    datos = ed[cols_ed].copy()
    datos["FECHA CONTACTO"] = pd.to_datetime(datos["FECHA CONTACTO"], errors="coerce")
    editado = st.data_editor(
        datos, hide_index=True, width="stretch", key="pol_editor",
        disabled=["ESTATUS", "CLIENTE", "TELÉFONO", "PRÓXIMA FECHA", "CIERRA EN", "SERIE"],
        column_config={**CONF,
                       "CONTACTADO": st.column_config.SelectboxColumn("CONTACTADO", options=["Sí", "No"]),
                       "FECHA CONTACTO": st.column_config.DateColumn("FECHA CONTACTO", format="DD/MM/YYYY"),
                       "RESULTADO": st.column_config.SelectboxColumn("RESULTADO", options=ps.RESULTADOS, width="medium"),
                       "NOTAS": st.column_config.TextColumn("NOTAS", width="large")})
    if st.button("Actualizar estatus", icon=":material/refresh:",
                 help="Pasa lo capturado a los indicadores y estatus. La descarga del Excel ya lo incluye."):
        st.session_state["pol_base"] = con_seguimiento(st.session_state["pol_base"], editado)
        st.session_state.pop("pol_editor", None)
        st.rerun()
    res = df[df["RESULTADO"].astype(str).str.strip() != ""]
    if not res.empty:
        st.bar_chart(res["RESULTADO"].value_counts().rename("Clientes"), horizontal=True, height=200,
                     color="#1D4A86")

# ----------------------------------------------------------------------------- descargas
st.divider()
d1, d2, _ = st.columns([1, 1, 2])
excel = ps.aplicar_seguimiento(st.session_state["pol_excel"], con_seguimiento(st.session_state["pol_base"], editado))
d1.download_button("Excel con seguimiento", excel, file_name=f"Polizas_extraidas_{hoy:%Y%m%d}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", width="stretch",
                   icon=":material/table_view:", type="primary")
mes_pdf, titulo = st.session_state.get("pol_pdf_mes", (mes_sig, f"{MESES[hoy.month - 1]} {hoy.year}"))
d2.download_button("PDF para contactar", ps.pdf(df, mes_pdf, titulo, hoy),
                   file_name=f"Polizas_contactar_{hoy:%Y%m%d}.pdf", mime="application/pdf", width="stretch",
                   icon=":material/picture_as_pdf:")
st.caption("El PDF trae a quién contactar ya y las oportunidades del mes elegido en «Por mes».")
