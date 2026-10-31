# -*- coding: utf-8 -*-
"""Utilidades compartidas por las páginas (tablas con color, descargas y el cuadre de Sale-U)."""
import pandas as pd
import streamlit as st

import conexion
import ui
from cuadres import saleu as cu
from cuadres import saleu_sheet, saleu_fuente
from cuadres.salidas import etiqueta

ESTADO = ui.ESTADO
HOJA_CORTA = lambda h: "Bitacora" if h == "Bitacora" else ("Finalizados" if "Finalizados" in str(h) else "Rechazos")


def aviso_excel(base):
    return " Descarga el Excel arriba." if isinstance(base, conexion.BaseExcel) else ""


# =============================================================================
# Tablas
# =============================================================================
# Fondos translúcidos con texto claro: se leen bien sobre el tema oscuro (como las insignias de shadcn).
VERDE = "background-color: rgba(34,197,94,.16); color: #86efac"
AMARILLO = "background-color: rgba(234,179,8,.16); color: #fde047"
ROJO = "background-color: rgba(239,68,68,.16); color: #fca5a5"
GRIS = "background-color: rgba(161,161,170,.12); color: #d4d4d8"


def colorear(valor):
    v = str(valor)
    if v in ("Falta crédito", "Registrar crédito"):
        return AMARILLO
    if v == "Dar de alta":
        return ROJO
    if v in ("OK", "Coincide por folio", "Coincide (teléfono y nombre)", "Cuadra"):
        return VERDE
    if (v.startswith("NO ") or v == "No cuadra" or "pero NO" in v or v.startswith("Falta")
            or v.startswith("En BlueService, no")):
        return ROJO
    if (v == "Revisar" or v.startswith("Coincide por") or v.startswith("Posible")
            or v.startswith("Cuadra –")):
        return AMARILLO
    return ""


def tabla(df, columnas, columnas_color=(), alto=None):
    cols = [c for c in columnas if c in df.columns]
    if df.empty:
        st.info("Sin registros.")
        return
    vista = df[cols].copy()
    vista.columns = [etiqueta(c) for c in cols]
    estilo = vista.style
    color_cols = [etiqueta(c) for c in columnas_color if c in cols]
    if color_cols:
        estilo = estilo.map(colorear, subset=color_cols)
    extra = {"height": alto} if alto else {}
    st.dataframe(estilo, width="stretch", hide_index=True, **extra)


def filtro_vendedor(df, clave, columna="Vendedor"):
    if df.empty or columna not in df.columns:
        return df
    opciones = sorted(v for v in df[columna].dropna().unique() if str(v).strip())
    elegidos = st.multiselect("Vendedor", opciones, key=clave, placeholder="Todos")
    return df[df[columna].isin(elegidos)] if elegidos else df


def valor_resumen(resumen, texto):
    fila = resumen[resumen["Concepto"].str.strip().str.startswith(texto)]
    return fila["Valor"].iloc[0] if not fila.empty else "—"


def descargas(res, prefijo, donde=st):
    """Botones de Excel y PDF juntos, alineados a la derecha."""
    fila = ui.barra(donde)
    fila.download_button("Excel", res["xlsx_bytes"], file_name=res["xlsx_nombre"],
                         mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                         key=f"{prefijo}_x", icon=":material/table_view:")
    fila.download_button("PDF", res["pdf_bytes"], file_name=res["pdf_nombre"], mime="application/pdf",
                         key=f"{prefijo}_p", icon=":material/picture_as_pdf:")


# =============================================================================
# Cuadre de Sale-U (revisar sin escribir; aplicar después)
# =============================================================================
def preparar_saleu(base, credito, fuentes_nuevas, clientes):
    """credito: (nombre, bytes) del reporte de Crédito; fuentes_nuevas: [(nombre, bytes)] reportes Fuente (opcional).
    Cuadra Sale-U contra el Sheet y propone el Registro en Sale-U y los datos faltantes. No escribe nada.
    Cada reporte se lee una sola vez y el mismo DataFrame sirve para el reporte del cuadre y para lo que se escribe."""
    su = cu.leer_saleu(credito)
    nuevas = [saleu_fuente.leer_fuente(f) for f in fuentes_nuevas]
    fu_nuevo = saleu_fuente.combinar(*nuevas) if nuevas else None
    # se suma a los leads ya guardados en el Sheet: un reporte del mes no pierde los leads de meses anteriores
    fu = saleu_fuente.combinar(saleu_fuente.guardada(base), fu_nuevo)
    res = cu.procesar(su, clientes, fuente="la bitácora", fuente_df=fu) if not clientes.empty else {}
    res.update(_propuesta_saleu(base, su, fu))
    res.update(con_fuente=fu is not None, fuente_df=fu, fuente_subida=fu_nuevo is not None, credito_df=su,
               periodo_fuente=saleu_fuente.periodo_txt(fu), archivo_credito=credito[0])
    return res


def _propuesta_saleu(base, su, fu):
    """Registro en Sale-U y datos por llenar que se escribirían en el Sheet con estos reportes."""
    cambios, (ini, fin) = saleu_sheet.proponer(base, su, fuente=fu)
    datos = pd.DataFrame(saleu_sheet.proponer_datos(base, su, fuente=fu),
                         columns=["Hoja", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo", "_valor"])
    return dict(registro=cambios, datos=datos, periodo_saleu=f"{ini:%d/%m/%Y} – {fin:%d/%m/%Y}")


def aplicar_saleu(base, res, volver_a_proponer=False):
    """Guarda los leads de Fuente y escribe el Registro en Sale-U y los datos faltantes. Con volver_a_proponer
    (porque antes se rehízo la Bitacora y cambiaron sus filas) recalcula los cambios sobre el Sheet ya actualizado.
    Devuelve (registros escritos, datos escritos)."""
    fu, su = res.get("fuente_df"), res["credito_df"]
    if res.get("fuente_subida") and fu is not None:
        saleu_fuente.guardar(base, fu)
    reg, datos = res["registro"], res["datos"]
    if volver_a_proponer:
        base.refrescar()
        nuevo = _propuesta_saleu(base, su, fu)
        reg, datos = nuevo["registro"], nuevo["datos"]
    escritos = saleu_sheet.aplicar(base, reg, res["periodo_saleu"])
    datos_escritos = saleu_sheet.aplicar_datos(base, datos.to_dict("records"))
    return escritos, datos_escritos
