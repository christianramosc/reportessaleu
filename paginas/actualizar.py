# -*- coding: utf-8 -*-
"""
Página: Actualizar

Un solo paso para todo lo que llega en el día: capturas de Cetelem, Bitácora de F&I, reportes de Sale-U (Crédito y
Fuente) y el Excel de pólizas de seguro. Se suben juntos (la app reconoce cada archivo), se revisan los cambios y se
aplican con un solo botón: primero las bitácoras (Rechazos, Finalizados y Bitacora), después el Registro en Sale-U
sobre el Sheet ya actualizado y al final la pestaña Polizas (cuadres/polizas_excel.py).
"""
import pandas as pd
import streamlit as st

import ui
import conexion
from comunes import (HOJA_CORTA, aviso_excel, aplicar_saleu, descargas, filtro_vendedor, preparar_saleu, tabla,
                     valor_resumen)
from cuadres import actualizar, cargas, detectar, fuentes, polizas_excel, saleu_sheet
from cuadres import saleu as cu
from cuadres.base import TAB_FINALIZADOS

ui.encabezado("Actualizar", "Sube los archivos del día: la app reconoce cada uno, revisa los cambios y los aplica juntos.")
ui.guia_mes()

base = conexion.panel_base()
if base is None:
    st.stop()

for clave, tipo in (("msg_todo", "success"), ("aviso_todo", "warning")):
    if st.session_state.get(clave):
        getattr(st, tipo)(st.session_state.pop(clave))

# ---------------------------------------------------------------- lo último que se subió
try:
    estado = cargas.estado(base)
except Exception as e:
    estado = []
    ui.error("No se pudo leer qué archivos se han subido", e)
if estado:
    ui.seccion("Lo último que se subió")
    cols = st.columns(len(estado))
    for c, e in zip(cols, estado):
        icono = ":green[:material/check_circle:]" if e["Nivel"] == "ok" else ":orange[:material/warning:]"
        caja = c.container(border=True, height="stretch")
        caja.markdown(f"{icono} **{e['Archivo']}**  \n{e['Último']}  \n<span class='tenue'>{e['Detalle']}</span>",
                      unsafe_allow_html=True)
        if e["Aviso"]:
            caja.markdown(f":orange[:material/notification_important: {e['Aviso']}]")

# ---------------------------------------------------------------- archivos
ui.seccion("Archivos del día", "Capturas de Cetelem, Bitácora de F&I, reportes de Sale-U (Crédito y Fuente) y tu "
           "Excel de pólizas de seguro, tal cual se descargan. Puedes subir uno o todos.")
caja_archivos = st.container(border=True)
archivos = caja_archivos.file_uploader(
    "Archivos del día", type=["xlsx", "csv"], accept_multiple_files=True, key="up_todo", label_visibility="collapsed",
    help="Capturas de Cetelem (JULIO_2026_FINAL.xlsx…), Bitácora de F&I, reporteLeadsconCredito y ReporteFuentes de "
         "Sale-U y Polizas_extraidas.xlsx (pólizas de seguro), tal cual se descargan. Puedes subir uno o todos; la app "
         "reconoce cada uno.")
clasif = detectar.clasificar([(f.name, f.getvalue()) for f in archivos or []])
if archivos:
    filas = [(n, detectar.NOMBRES[t]) for t, lista in clasif.items() for n, _ in lista]
    caja_archivos.dataframe(pd.DataFrame(filas, columns=["Archivo", "Se reconoció como"]), hide_index=True,
                            width="stretch")
    if clasif.get(None):
        caja_archivos.warning("No reconocí: " + ", ".join(n for n, _ in clasif[None]) + ". Súbelos tal cual se descargan.")
    if clasif.get("base"):
        caja_archivos.info("El Excel de seguimiento no se sube aquí" + (" (se trabaja directo en el Google Sheet)."
                if conexion.google_configurado() else "; súbelo arriba, en «Modo Excel»."))
    if len(clasif.get("credito", [])) > 1:
        caja_archivos.warning(f"Subiste {len(clasif['credito'])} reportes de Crédito; se usa «{clasif['credito'][0][0]}».")
    if len(clasif.get("polizas", [])) > 1:
        caja_archivos.warning(f"Subiste {len(clasif['polizas'])} Excel de pólizas; se usa «{clasif['polizas'][0][0]}».")

tablas_fi, titulo_fi = None, None
try:
    tablas_fi, titulo_fi = conexion.tablas_financiamiento(base)
except Exception as e:
    ui.error("No se pudo leer la Bitácora de F&I de Google", e)
if tablas_fi is not None:
    caja_archivos.caption(f":material/link: La Bitácora de F&I se lee sola de Google: **{titulo_fi}**.")
elif clasif.get("fi"):
    tablas_fi = {}
    for nombre, datos in clasif["fi"]:
        tablas_fi.update({f"{nombre} · {h}": v for h, v in fuentes.tablas_de_excel(datos).items()})
    titulo_fi = ", ".join(n for n, _ in clasif["fi"])

guardadas = conexion.capturas_guardadas(base)
fi_guardados, _ = conexion.registros_fi_guardados(base)
hay_bitacoras = bool(clasif.get("cetelem")) or not guardadas.empty or tablas_fi is not None or bool(fi_guardados)
puede = hay_bitacoras or bool(clasif.get("credito")) or bool(clasif.get("fuente")) or bool(clasif.get("polizas"))


def revisar():
    todo = dict(cargas=[])
    if hay_bitacoras:
        nuevas = fuentes.leer_capturas(clasif.get("cetelem", []))
        capturas = fuentes.combinar_capturas(guardadas, nuevas) if not nuevas.empty else guardadas
        regs = fuentes.registros_capturas(capturas)
        hojas_fi, regs_fi = [], []
        if tablas_fi is not None:
            regs_fi, hojas_fi = fuentes.registros_financiamiento(tablas_fi)
            regs_fi = fuentes.combinar_fi(fi_guardados, regs_fi)          # se acumula con lo ya guardado
        elif fi_guardados:
            regs_fi, hojas_fi = fi_guardados, ["copia guardada en el Sheet"]
        sol = fuentes.consolidar(regs + regs_fi)
        prop, ctx = actualizar.proponer(base, sol)
        todo["act"] = dict(prop=prop, ctx=ctx, sol=sol, capturas=capturas, cambia_capturas=not nuevas.empty,
                           meses=list(dict.fromkeys(capturas["Hoja"])) if not capturas.empty else [],
                           hojas_fi=hojas_fi, regs_fi=regs_fi if tablas_fi is not None else None)
        for n, _ in clasif.get("cetelem", []):
            meses = sorted(set(nuevas[nuevas["Archivo"] == n]["Hoja"])) if not nuevas.empty else []
            todo["cargas"].append(dict(Tipo="cetelem", Archivo=n, Periodo=", ".join(meses),
                                       Registros=int((nuevas["Archivo"] == n).sum()) if not nuevas.empty else 0))
        if tablas_fi is not None:
            todo["cargas"].append(dict(Tipo="fi", Archivo=titulo_fi, Periodo=", ".join(hojas_fi), Registros=len(regs_fi)))
    if clasif.get("credito"):
        su = preparar_saleu(base, clasif["credito"][0], clasif.get("fuente", []), conexion.clientes_saleu(base))
        todo["su"] = su
        todo["cargas"].append(dict(Tipo="credito", Archivo=clasif["credito"][0][0], Periodo=su["periodo_saleu"],
                                   Registros=len(su["credito_df"])))
    if clasif.get("fuente"):
        from cuadres import saleu_fuente
        nuevas_fu = [saleu_fuente.leer_fuente(f) for f in clasif["fuente"]]
        if "su" not in todo:         # sin reporte de Crédito: solo se guardan los leads
            todo["solo_fuente"] = saleu_fuente.combinar(saleu_fuente.guardada(base), *nuevas_fu)
        for (n, _), f in zip(clasif["fuente"], nuevas_fu):
            todo["cargas"].append(dict(Tipo="fuente", Archivo=n, Periodo=saleu_fuente.periodo_txt(f), Registros=len(f)))
    if clasif.get("polizas"):
        nombre, datos = clasif["polizas"][0]
        todo["pol"] = polizas_excel.proponer(base, nombre, datos)
        todo["cargas"].append(dict(Tipo="polizas", Archivo=nombre, Periodo="", Registros=len(todo["pol"]["df"])))
    return todo


fila = ui.barra(caja_archivos, alinear="left")
if fila.button("Revisar cambios", type="primary", disabled=not puede, key="btn_revisar_todo", icon=":material/search:"):
    with st.spinner("Comparando con el Sheet…"):
        try:
            st.session_state["todo"] = revisar()
        except Exception as e:
            st.session_state.pop("todo", None)
            ui.error("No se pudo revisar", e)
if not puede:
    fila.caption("Sube al menos un archivo.")

todo = st.session_state.get("todo")
if todo:
    act, su = todo.get("act"), todo.get("su")
    prop = act["prop"] if act else pd.DataFrame()

    # ------------------------------------------------ 1. bitácoras
    if act:
        ui.seccion("Bitácoras de Cetelem y F&I", "Cetelem: " + (", ".join(act["meses"]) or "—")
                   + "  ·  F&I: " + (", ".join(act["hojas_fi"]) or "—"), numero=1)
        m = st.columns(5)
        m[0].metric("Rechazos nuevos", int((prop["Tipo"] == "Nuevo rechazo").sum()))
        m[1].metric("Correcciones", int((prop["Tipo"] == "Corrección").sum()))
        m[2].metric("Datos faltantes", int((prop["Tipo"] == "Dato faltante").sum()),
                    help="Celdas vacías que se llenan con la bitácora (cuadre por teléfono o nombre).")
        m[3].metric("Comentarios", int((prop["Tipo"] == "Comentario").sum()))
        m[4].metric("Solicitudes", len(act["sol"]))
        if prop.empty:
            st.success("Rechazos y Finalizados ya están al día con las bitácoras.")
        else:
            st.markdown("**Desmarca** lo que no quieras aplicar; no se volverá a proponer.")
            vista = st.data_editor(
                prop[["Aplicar", "Tipo", "Hoja", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo"]].assign(
                    Hoja=prop["Hoja"].map(lambda h: "Finalizados" if h == TAB_FINALIZADOS else "Rechazos"),
                    Fila=prop["Fila"].map(lambda f: "nueva" if f is None or f != f else str(int(f)))),
                hide_index=True, width="stretch", height=min(420, 60 + 35 * len(prop)), key="editor_todo",
                disabled=["Tipo", "Hoja", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo"],
                column_config={"Aplicar": st.column_config.CheckboxColumn(width="small"),
                               "Fila": st.column_config.TextColumn(width="small"),
                               "Motivo": st.column_config.TextColumn(width="large")})
            prop = prop.assign(Aplicar=vista["Aplicar"].values)

    # ------------------------------------------------ 2. Sale-U
    n_su = 0
    if su:
        reg, datos = su["registro"], su["datos"]
        cambian = reg[reg["Cambia"]]
        n_su = len(cambian) + len(datos)
        ui.seccion("Sale-U", f"Crédito {su['periodo_saleu']}"
                   + (f" · Fuente {su['periodo_fuente']}" if su.get("con_fuente") else " · sin reporte Fuente"),
                   numero=2)
        k = st.columns(5)
        k[0].metric("Registrados", int((reg["Después"] == saleu_sheet.REGISTRADO).sum()))
        k[1].metric("Falta crédito", int((reg["Después"] == saleu_sheet.FALTA_CREDITO).sum()),
                    help="Están en el reporte Fuente, pero no en el de Crédito.")
        k[2].metric("No Registrado", int((reg["Después"] == saleu_sheet.NO_REGISTRADO).sum()),
                    help="No están en Sale-U: falta darlos de alta.")
        k[3].metric("Registro por cambiar", len(cambian),
                    help="Registrado / Falta crédito / No Registrado. Nunca baja de nivel lo que ya dice el Sheet.")
        k[4].metric("Datos por llenar", len(datos), help="Celdas vacías que se llenan con Sale-U o la bitácora.")
        if len(cambian):
            with st.expander(f"Registro en Sale-U que cambia ({len(cambian)})"):
                tabla(cambian.assign(Hoja=cambian["Hoja"].map(HOJA_CORTA)),
                      ["Hoja", "Fila", "Cliente", "Antes", "Después", "Como"], ["Después"])
        if len(datos):
            with st.expander(f"Datos por llenar ({len(datos)})"):
                tabla(datos.assign(Hoja=datos["Hoja"].map(HOJA_CORTA), Después=datos["Después"].astype(str)),
                      ["Hoja", "Fila", "Cliente", "Campo", "Después", "Motivo"])
        rev = reg[reg["Revisar"] != ""] if "Revisar" in reg.columns else reg.iloc[0:0]
        if len(rev):
            with st.expander(f"Marcados como Registrado que no salen en los reportes ({len(rev)})"):
                st.caption("Se dejan igual; confírmalos en Sale-U.")
                tabla(rev.assign(Hoja=rev["Hoja"].map(HOJA_CORTA)), ["Hoja", "Fila", "Cliente", "Antes", "Revisar"])
        if "resumen" in su:
            with st.container(border=True):
                t, d = st.columns([3, 2], vertical_alignment="center")
                t.markdown(":material/description: **Reporte del cuadre Sale-U vs bitácora**  \n"
                           "<span class='tenue'>Todos los estatus y financieras, con quién falta en Sale-U.</span>",
                           unsafe_allow_html=True)
                descargas(su, "su_todo", donde=d)
                r = su["resumen"]
                m = st.columns(4)
                m[0].metric("En bitácora", valor_resumen(r, "Clientes únicos"))
                m[1].metric("En Sale-U", valor_resumen(r, "Clientes de la bitácora encontrados"))
                m[2].metric("Faltan", valor_resumen(r, "Clientes de la bitácora NO"))
                m[3].metric("Estatus distinto", valor_resumen(r, "Estatus que no cuadra"))
                df = filtro_vendedor(su["faltantes"], "fv_todo")
                vista = cu._faltantes_con_vendedor(df)
                vista["Fecha"] = vista["Fecha"].map(cu.fecha_txt)
                tabla(vista, (["Falta"] if "Falta" in vista else []) + ["Fecha", "Nombre", "Vendedor", "Celular",
                      "Financiera", "Estatus", cu.COL_SALEU, "Fuente"], ["Falta"])
        rech = su.get("rechazos")
        if rech is not None:
            with st.container(border=True):
                t, d = st.columns([3, 2], vertical_alignment="center")
                t.markdown(":material/fact_check: **Registro en Sale-U · MG Colima Rechazos**  \n"
                           "<span class='tenue'>Cómo está cada rechazo en Sale-U: créditos por financiera, estatus y "
                           "con qué nombre y teléfono está capturado."
                           + (" Sus hojas («Rechazos · …») van en el Excel del reporte de arriba."
                              if "xlsx_bytes" in su else "") + "</span>", unsafe_allow_html=True)
                if "xlsx_bytes" in rech:
                    descargas(rech, "su_rech", donde=d)
                else:
                    ui.barra(d).download_button("PDF", rech["pdf_bytes"], file_name=rech["pdf_nombre"],
                                                mime="application/pdf", key="su_rech_p",
                                                icon=":material/picture_as_pdf:")
                det = rech["detalle"]
                cuenta = lambda k: int((det["_clave"] == k).sum()) if len(det) else 0
                m = st.columns(5)
                m[0].metric("Rechazos", len(det))
                m[1].metric("Registrado", int((det["Registro_cuadre"] == saleu_sheet.REGISTRADO).sum()) if len(det) else 0)
                m[2].metric("No Registrado", cuenta("alta"))
                m[3].metric("Otra financiera / estatus", cuenta("otra_fin") + cuenta("estatus"),
                            help="En Sale-U el crédito está con otra financiera, o con la misma pero no rechazado.")
                m[4].metric("Nombre o teléfono distinto", cuenta("dato"))
                pr = rech["por_revisar"]
                if len(pr):
                    pr = pr.assign(_o=pr["_clave"].map(rech["orden"])).sort_values(["_o", "Fila"])
                    tabla(pr.assign(Fecha_ingreso=pr["Fecha_ingreso"].map(cu.fecha_txt)),
                          ["Situacion", "Fila", "Fecha_ingreso", "Cliente", "Celular", "Financiera", "Asesor",
                           "Creditos_SaleU", "Como_en_SaleU", "Observaciones"])
                else:
                    st.success("Todos los rechazos del periodo cuadran con Sale-U.")
    elif todo.get("solo_fuente") is not None:
        ui.seccion("Sale-U", numero=2)
        st.info(f"Solo subiste el reporte Fuente: sus leads se guardan al aplicar "
                f"({len(todo['solo_fuente'])} acumulados). Para revisar el Registro en Sale-U sube también el de Crédito.")

    # ------------------------------------------------ 3. pólizas de seguro
    pol = todo.get("pol")
    conservar = True
    if pol:
        ui.seccion("Pólizas de seguro", f"«{pol['nombre']}» → pestaña Polizas del Sheet", numero=3)
        m = st.columns(4)
        m[0].metric("Pólizas en el Excel", len(pol["df"]), help="Una fila por crédito (hoja Contacto).")
        m[1].metric("En la pestaña hoy", len(pol["actual"]))
        m[2].metric("Nuevas", len(pol["altas"]), help="Ligadas por SERIE (VIN) o, si no hay, por nombre.")
        m[3].metric("Con teléfono", pol["con_tel"])
        if len(pol["altas"]):
            with st.expander(f"Pólizas nuevas ({len(pol['altas'])})"):
                st.dataframe(pol["altas"][[c for c in ("CLIENTE", "TELÉFONO", "VEHÍCULO", "SERIE", "ASEGURADORA")
                                           if c in pol["altas"].columns]].astype(str),
                             hide_index=True, width="stretch")
        quedan = pol["solo_pestana"]
        if len(quedan):
            conservar = st.checkbox(f"Conservar {len(quedan)} póliza(s) que están en la pestaña y no vienen en este "
                                    "Excel", value=True, key="pol_conservar",
                                    help="Desmárcalo si tu Excel ya trae a todos y quieres que la pestaña quede igual "
                                         "al Excel (como lo deja la app de seguros).")
            with st.expander(f"Están en la pestaña, pero no en el Excel ({len(quedan)})"):
                st.dataframe(quedan[[c for c in ("CLIENTE", "TELÉFONO", "VEHÍCULO", "SERIE") if c in quedan.columns]]
                             .astype(str), hide_index=True, width="stretch")
        st.caption("Al aplicar, la pestaña Polizas se reescribe (antes se guarda un respaldo oculto). Después ve a "
                   "**Mantenimiento → Revisar contra Pólizas** para pasar teléfonos, correos y vehículos a "
                   "Créditos Finalizados.")

    # ------------------------------------------------ aplicar todo
    st.write("")
    n_act = int(prop["Aplicar"].sum()) if not prop.empty else 0
    hay_algo = (n_act or n_su or (act and (act["cambia_capturas"] or act.get("regs_fi") is not None))
                or todo.get("solo_fuente") is not None or todo["cargas"] or pol)
    etiqueta = (f"Aplicar todo ({n_act + n_su} cambio(s))" if n_act + n_su
                else "Actualizar pólizas" if pol and not (act or su) else "Guardar archivos")
    pie = st.container(border=True)
    t, b = pie.columns([3, 1], vertical_alignment="center")
    t.markdown(":material/done_all: **Aplicar al Sheet**  \n<span class='tenue'>Primero las bitácoras, después el "
               "Registro en Sale-U sobre el Sheet ya actualizado y al final las pólizas. Queda en el Historial.</span>",
               unsafe_allow_html=True)
    if b.button(etiqueta, type="primary", disabled=not hay_algo, key="btn_aplicar_todo", icon=":material/done_all:",
                width="stretch"):
        with st.spinner("Escribiendo en el Sheet…"):
            partes, aviso = [], None
            try:
                hubo_act = False
                if act and (n_act or act["cambia_capturas"] or act.get("regs_fi") is not None):
                    r = actualizar.aplicar(base, prop, act["sol"], act["ctx"], act["capturas"], act.get("regs_fi"))
                    hubo_act, aviso = True, r.get("aviso")
                    partes.append(f"bitácoras: {r['nuevos']} nuevo(s), {r['correcciones']} corrección(es), "
                                  f"{r.get('faltantes', 0)} dato(s)")
                if su:
                    escritos, datos_escritos = aplicar_saleu(base, su, volver_a_proponer=hubo_act)
                    partes.append(f"Sale-U: {escritos} registro(s) y {datos_escritos} dato(s)")
                elif todo.get("solo_fuente") is not None:
                    from cuadres import saleu_fuente
                    n = saleu_fuente.guardar(base, todo["solo_fuente"])
                    partes.append(f"{n} leads de Fuente guardados")
                if pol:
                    rp = polizas_excel.aplicar(base, pol, conservar=conservar)
                    partes.append(f"Pólizas: {rp['filas']} en la pestaña ({rp['altas']} nueva(s)"
                                  + (f", {rp['conservadas']} conservada(s)" if rp["conservadas"] else "") + ")")
                    if rp["aviso"]:
                        aviso = f"{aviso}\n\n{rp['aviso']}" if aviso else rp["aviso"]
                cargas.anotar(base, todo["cargas"])
                conexion.recargar()
                st.session_state.pop("todo", None)
                st.session_state["msg_todo"] = "Listo — " + " · ".join(partes or ["archivos guardados"]) + "." \
                    + aviso_excel(base)
                if aviso:
                    st.session_state["aviso_todo"] = aviso
                st.rerun()
            except Exception as e:
                ui.error("No se pudo aplicar", e)
