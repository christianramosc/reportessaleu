# -*- coding: utf-8 -*-
"""
Mantenimiento del Sheet de seguimiento (lo llama formato_sheet.aplicar y, en parte, Actualizar):

  · limpiar_respaldos: deja solo el respaldo más reciente de cada pestaña y borra las copias ocultas sueltas
    («Hoja 1», «Hoja 2»… con los mismos encabezados que una pestaña de datos).
  · llenar_fijos: en MG Colima Rechazos, las columnas que siempre traen lo mismo (Tipo de Plan, Tipo de Subvención,
    Accesorios, Status 1) se llenan solas donde estén vacías, con su valor de siempre.
  · ligar_finalizados: columna «ID en 'MG Colima Rechazos'» en Créditos Finalizados (el rechazo previo del cliente,
    por teléfono o nombre).
  · proteger: columnas calculadas de Rechazos y la Bitacora completa con protección «solo advertencia».
  · vistas_por_asesor: una vista filtrada de MG Colima Rechazos por asesor y encabezado congelado.
"""
import re
from collections import Counter

from cuadres.base import TAB_RECHAZOS, TAB_FINALIZADOS, TAB_BITACORA, canonico, norm
from cuadres.fuentes import name_key, similar, to_phone, COL_ID_RECHAZOS

COLUMNAS_FIJAS = ["Tipo de Plan", "Tipo de Subvención", "Accesorios", "Status 1"]
_RESP = re.compile(r"^_resp (.+?) (\d{4}-\d{2}-\d{2} \d{6}|\d{4} \d{6}|\d{4}-\d)")


# =============================================================================
# Respaldos
# =============================================================================
def _encabezado(base, pestana):
    try:
        v = base._valores(pestana, True)
        return [norm(canonico(str(x))) for x in (v[0] if v else []) if x not in (None, "")]
    except Exception:
        return []


def limpiar_respaldos(base):
    """Borra respaldos viejos (deja el más reciente por pestaña) y copias ocultas «Hoja N». Devuelve los nombres."""
    pestanas = base.pestanas()
    grupos = {}
    for p in pestanas:
        m = _RESP.match(p)
        if m or p.startswith("_resp "):
            grupos.setdefault(norm(m.group(1) if m else p[:18]), []).append(p)
    borrar = []
    for nombres in grupos.values():
        borrar += sorted(nombres)[:-1]               # el nombre lleva la fecha: el último es el más reciente
    datos = [p for p in pestanas if not p.startswith("_resp ") and not re.fullmatch(r"Hoja \d+", p)]
    encabezados = [_encabezado(base, p)[:6] for p in datos]
    for p in pestanas:
        if re.fullmatch(r"Hoja \d+", p):
            enc = _encabezado(base, p)[:6]
            if enc and any(enc[:5] == e[:5] for e in encabezados if e):   # copia de una pestaña de datos
                borrar.append(p)
    for p in borrar:
        base.borrar(p)
    return borrar


# =============================================================================
# Columnas fijas
# =============================================================================
def llenar_fijos(base):
    """Llena las columnas fijas vacías de MG Colima Rechazos con su valor más común. Devuelve cuántas celdas."""
    if TAB_RECHAZOS not in base.pestanas():
        return 0
    df = base.leer(TAB_RECHAZOS)
    c_nom = next((c for c in df.columns if norm(c) == norm("Nombre de Cliente")), None)
    if df.empty or not c_nom:
        return 0
    celdas = []
    for nombre in COLUMNAS_FIJAS:
        c = next((x for x in df.columns if norm(x) == norm(nombre)), None)
        col = base.col(TAB_RECHAZOS, nombre)
        if not c or not col:
            continue
        valores = Counter(str(v).strip() for v in df[c] if v not in (None, "") and str(v).strip())
        if not valores:
            continue
        valor = valores.most_common(1)[0][0]
        for _, r in df.iterrows():
            if r[c_nom] not in (None, "") and r[c] in (None, ""):
                celdas.append((int(r["_fila"]), col, valor))
    if celdas:
        base.escribir(TAB_RECHAZOS, celdas)
    return len(celdas)


# =============================================================================
# Finalizados ↔ Rechazos
# =============================================================================
def ligar_finalizados(base):
    """Escribe en Créditos Finalizados el ID del rechazo previo del cliente (mismo teléfono o nombre). Devuelve
    cuántas filas cambiaron."""
    from cuadres import ids
    if TAB_FINALIZADOS not in base.pestanas() or TAB_RECHAZOS not in base.pestanas():
        return 0
    fin = base.leer(TAB_FINALIZADOS)
    if fin.empty:
        return 0
    c_nom = next((c for c in fin.columns if norm(c) == "NOMBRE"), None)
    c_tel = next((c for c in fin.columns if norm(c) in ("TELEFONO", "CELULAR")), None)
    if not c_nom:
        return 0
    rech = base.leer(TAB_RECHAZOS)
    r_nom = next((c for c in rech.columns if norm(c) == norm("Nombre de Cliente")), None)
    r_tel = next((c for c in rech.columns if norm(c) == "CELULAR"), None)
    id_de = ids.mapa(base)
    rechazos = [(id_de.get(int(r["_fila"])), name_key(r[r_nom]), to_phone(r[r_tel]) if r_tel else None)
                for _, r in rech.iterrows() if r_nom and r[r_nom]]
    rechazos = [x for x in rechazos if x[0]]
    col = base.agregar_columna(TAB_FINALIZADOS, COL_ID_RECHAZOS)
    c_id = next((c for c in fin.columns if norm(c) == norm(COL_ID_RECHAZOS)), None)
    celdas = []
    for _, r in fin.iterrows():
        if not r[c_nom]:
            continue
        k, t = name_key(r[c_nom]), to_phone(r[c_tel]) if c_tel else None
        ligados = [i for i, kr, tr in rechazos if (t and tr == t) or similar(k, kr)]
        valor = ", ".join(dict.fromkeys(ligados))
        antes = str(r[c_id] or "").strip() if c_id else ""
        if valor and valor != antes:
            celdas.append((int(r["_fila"]), col, valor))
    if celdas:
        base.escribir(TAB_FINALIZADOS, celdas)
    return len(celdas)


# =============================================================================
# Protección, vistas y orden
# =============================================================================
CALCULADAS = ["Fecha de Seguimiento", "EFECTIVIDAD 1", "EFECTIVIDAD 2", "EFECTIVIDAD 3", "Reingreso", "ID",
              "Registro en Sale-U"]


def proteger(base):
    hecho = []
    if TAB_RECHAZOS in base.pestanas():
        cols = [c for c in (base.col(TAB_RECHAZOS, n) for n in CALCULADAS) if c]
        if cols and base.proteger(TAB_RECHAZOS, cols, "App: columnas calculadas (fórmulas, ID y Registro en Sale-U)"):
            hecho.append("Columnas calculadas de MG Colima Rechazos protegidas (solo advertencia)")
    if TAB_BITACORA in base.pestanas() and base.proteger(TAB_BITACORA, None,
                                                         "App: la Bitacora se rehace sola con Actualizar"):
        hecho.append("Bitacora protegida (solo advertencia): se corrige en MG Colima Rechazos o en las bitácoras")
    return hecho


def vistas_por_asesor(base):
    """Encabezado y nombre congelados en Rechazos y Finalizados; una vista filtrada por asesor en Rechazos."""
    from cuadres.actualizar import _catalogo_asesores
    base.congelar(TAB_RECHAZOS, 1, 3)
    base.congelar(TAB_FINALIZADOS, 1, 2)
    col = base.col(TAB_RECHAZOS, "Asesor de Ventas")
    if not col:
        return 0
    df = base.leer(TAB_RECHAZOS)
    c = next((x for x in df.columns if norm(x) == norm("Asesor de Ventas")), None)
    usados = {str(v).strip() for v in df[c] if v not in (None, "")} if c else set()
    asesores = [a for a in _catalogo_asesores(base) if a in usados] or sorted(usados)
    return sum(1 for a in asesores if base.vista_filtrada(TAB_RECHAZOS, f"Asesor: {a}", col, a))
