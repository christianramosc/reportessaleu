# -*- coding: utf-8 -*-
"""
ID fijo de cada rechazo y marca de reingreso en «MG Colima Rechazos».

  · ID (R-0001, R-0002…): no cambia aunque alguien ordene la hoja o inserte filas. La Bitacora y la pestaña
    Revision guardan el ID (no el número de fila), así siguen apuntando al mismo cliente.
  · Reingreso: fórmula que dice si el celular aparece en otra solicitud («1 de 2», «2 de 2»), para no
    contactar dos veces a la misma persona.
Las dos columnas van al final de la pestaña para no mover las columnas que ya existen.
"""
import re

from cuadres.base import TAB_RECHAZOS, letra, norm

COL_ID = "ID"
COL_REINGRESO = "Reingreso"
PREFIJO = "R-"
COLOR_REINGRESO = "#FFF2CC"
FILAS_FORMULA = 1000            # igual que las fórmulas que ya trae la hoja (Fecha de Seguimiento, Efectividad)
_RE_ID = re.compile(r"^R-(\d+)$")


def _numero(v):
    m = _RE_ID.match(str(v or "").strip().upper())
    return int(m.group(1)) if m else None


def formato(n):
    return f"{PREFIJO}{n:04d}"


def _filas(base):
    """[(fila, nombre, id)] de las filas con cliente."""
    df = base.leer(TAB_RECHAZOS)
    if df.empty:
        return []
    c_nom = next((c for c in df.columns if norm(c) == norm("Nombre de Cliente")), None)
    c_id = next((c for c in df.columns if norm(c) == norm(COL_ID)), None)
    out = []
    for _, r in df.iterrows():
        nombre = r[c_nom] if c_nom else None
        if nombre is None or not str(nombre).strip():
            continue
        out.append((int(r["_fila"]), str(nombre).strip(), str(r[c_id]).strip() if c_id and r[c_id] else ""))
    return out


def mapa(base):
    """{fila: ID} de los rechazos que ya tienen ID."""
    return {f: i for f, _, i in _filas(base) if i}


def siguiente(usados):
    return max([_numero(i) or 0 for i in usados] + [0]) + 1


def asegurar(base):
    """Crea la columna ID si falta y le pone ID a cada rechazo que no lo tenga (en orden de fila).
    Devuelve {fila: ID} de todos los rechazos."""
    if TAB_RECHAZOS not in base.pestanas():
        return {}
    filas = _filas(base)
    if not filas:
        return {}
    if not base.col(TAB_RECHAZOS, COL_ID) and not base.col(TAB_RECHAZOS, COL_REINGRESO):
        asegurar_reingreso(base)          # mismo orden siempre: … COMENTARIOS | Reingreso | ID
    c = base.agregar_columna(TAB_RECHAZOS, COL_ID)
    usados = {i for _, _, i in filas if i}
    n = siguiente(usados)
    celdas, out = [], {}
    for fila, _, i in filas:
        if not i:
            i = formato(n)
            n += 1
            celdas.append((fila, c, i))
        out[fila] = i
    base.escribir(TAB_RECHAZOS, celdas)
    return out


def formula_reingreso(fila, l_nombre, l_tel, hasta=FILAS_FORMULA):
    """«1 de 2», «2 de 2»… si el celular está en más de una fila; vacío si es único."""
    tel = f"${l_tel}{fila}"
    todos = f"COUNTIF(${l_tel}$2:${l_tel}${hasta},{tel})"
    hasta_aqui = f"COUNTIF(${l_tel}$2:${l_tel}{fila},{tel})"
    return (f'=IF(OR(${l_nombre}{fila}="",{tel}=""),"",IF({todos}<2,"",{hasta_aqui}&" de "&{todos}))')


def letras(base):
    c_nom, c_tel = base.col(TAB_RECHAZOS, "Nombre de Cliente"), base.col(TAB_RECHAZOS, "Celular")
    return (letra(c_nom), letra(c_tel)) if c_nom and c_tel else (None, None)


def asegurar_reingreso(base, hasta=FILAS_FORMULA):
    """Columna Reingreso con su fórmula en las filas 2..hasta y el color de aviso. Devuelve la columna."""
    l_nom, l_tel = letras(base)
    if not l_nom:
        return None
    c = base.agregar_columna(TAB_RECHAZOS, COL_REINGRESO)
    base.escribir(TAB_RECHAZOS, [(f, c, formula_reingreso(f, l_nom, l_tel, hasta)) for f in range(2, hasta + 1)])
    base.regla_color(TAB_RECHAZOS, [c], f'=${letra(c)}2<>""', COLOR_REINGRESO, hasta)
    return c
