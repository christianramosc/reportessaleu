# -*- coding: utf-8 -*-
"""
Registro en Sale-U → columna "Cliente con Registro en SalesU" de MG Colima Rechazos.

Con el reporte de Sale-U (reporteLeadsconCredito) se revisa cada rechazo de la base cuya fecha
de ingreso cae dentro del periodo del reporte:
  · encontrado (mismo teléfono y nombre parecido, o nombre muy parecido) → "Registrado"
  · no encontrado                                                        → "No Registrado"
Lo que ya dice que SÍ está registrado (p. ej. marcado a mano) nunca se baja a "No Registrado":
si no aparece en el reporte se deja como registrado y se lista para revisar.
Los rechazos fuera del periodo del reporte no se tocan (no hay con qué compararlos).
"""

import pandas as pd

from cuadres import saleu as cu
from cuadres.base import TAB_RECHAZOS, letra, norm
from cuadres.actualizar import leer_rechazos, C

REGISTRADO, NO_REGISTRADO = "Registrado", "No Registrado"


def ya_registrado(valor):
    """'Registrado', 'Sí', 'SI', 'Si registrado'… (no 'No Registrado')."""
    n = norm(valor)
    return n in ("REGISTRADO", "REGISTRADA") or n == "SI" or n.startswith("SI ")


def periodo(su):
    fechas = su["_fecha"].dropna()
    if fechas.empty:
        return None, None
    return fechas.min().date(), fechas.max().date()


def proponer(base, su, ini=None, fin=None):
    """su: DataFrame de cu.leer_saleu(). Devuelve DataFrame con lo que se escribiría."""
    p_ini, p_fin = periodo(su)
    ini, fin = ini or p_ini, fin or p_fin
    if not ini:
        raise ValueError("El reporte de Sale-U no trae fechas de alta de crédito.")
    por_tel = {}
    for _, s in su.iterrows():
        if s["Tel_norm"]:
            por_tel.setdefault(s["Tel_norm"], []).append(s)
    filas = []
    for row in leer_rechazos(base):
        if not row["fecha"] or not (ini <= row["fecha"] <= fin):
            continue
        nom = cu.normalizar_nombre(row["nombre"])
        tel = str(row["cel"]) if row["cel"] else ""
        hallado, como = None, ""
        for s in por_tel.get(tel, []):
            r, _ = cu.similitud_nombre(nom, s["Nombre_norm"])
            if r >= cu.UMBRAL_NOMBRE_CON_TEL:
                hallado, como = s, "teléfono y nombre"
                break
        if hallado is None:
            mejor = (0, 0, None)
            for _, s in su.iterrows():
                r, k = cu.similitud_nombre(nom, s["Nombre_norm"])
                if (r, k) > mejor[:2]:
                    mejor = (r, k, s)
            if mejor[0] >= cu.UMBRAL_NOMBRE_FUERTE and mejor[1] >= 2:
                hallado, como = mejor[2], "nombre (teléfono distinto)"
        actual = row["saleu"] or ""
        revisar = ""
        if hallado is not None:
            nuevo = REGISTRADO
        elif ya_registrado(actual):
            nuevo = REGISTRADO          # lo marcaron como registrado: se respeta (solo se uniforma la palabra)
            revisar = "Marcado como registrado, pero no aparece en este reporte de Sale-U"
            como = "se respetó lo que ya decía"
        else:
            nuevo = NO_REGISTRADO
        filas.append(dict(
            Fila=row["fila"], Cliente=row["nombre"], Fecha_ingreso=row["fecha"], Antes=actual, Después=nuevo,
            Cambia=actual != nuevo, Como=como, Revisar=revisar,
            Cliente_SaleU=hallado["Nombre"] if hallado is not None else "",
            Estatus_SaleU=hallado["Estatus_SaleU"] if hallado is not None else "",
            Financiera_SaleU=hallado["Financiera"] if hallado is not None else ""))
    df = pd.DataFrame(filas, columns=["Fila", "Cliente", "Fecha_ingreso", "Antes", "Después", "Cambia", "Como",
                                      "Revisar", "Cliente_SaleU", "Estatus_SaleU", "Financiera_SaleU"])
    return df, (ini, fin)


def aplicar(base, cambios, periodo_txt):
    """Escribe solo las filas cuyo valor cambia y las anota en el Historial."""
    cambios = cambios[cambios["Cambia"]]
    if cambios.empty:
        return 0
    col = base.col(TAB_RECHAZOS, C["saleu"])
    if not col:
        raise ValueError(f"No encuentro la columna '{C['saleu']}' en '{TAB_RECHAZOS}'.")
    base.respaldar(TAB_RECHAZOS)
    base.escribir(TAB_RECHAZOS, [(int(r["Fila"]), col, r["Después"]) for _, r in cambios.iterrows()])
    base.anotar_historial([dict(
        Origen="App · cuadre Sale-U", Hoja=TAB_RECHAZOS, Celda=f"{letra(col)}{int(r['Fila'])}", Cliente=r["Cliente"],
        Campo=C["saleu"], Antes=r["Antes"], Después=r["Después"],
        Motivo=(f"Ya decía que estaba registrado; se uniformó la palabra (no aparece en el reporte {periodo_txt})"
                if r["Revisar"] else
                f"Encontrado en Sale-U por {r['Como']}: {r['Cliente_SaleU']} ({r['Estatus_SaleU']})" if r["Después"] == REGISTRADO
                else f"No aparece en el reporte de Sale-U ({periodo_txt})"))
        for _, r in cambios.iterrows()])
    return len(cambios)
