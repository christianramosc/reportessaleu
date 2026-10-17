# -*- coding: utf-8 -*-
"""
Limpiar pendientes resueltos.

Lee la pestaña "Revision" y, para cada fila con Estado = "Pendiente", revisa si el dato ya se
capturó en la hoja. Si está resuelto: quita el relleno morado y la nota de la celda, cambia el
Estado a "Corregido" y lo anota en el Historial.

También quita la marca de los pendientes que alguien ya cambió a mano a "Corregido" o "No aplica".
"""
import re
import datetime as dt

import pandas as pd

from cuadres.base import TAB_RECHAZOS, TAB_FINALIZADOS, norm, letra, canonico
from cuadres.fuentes import name_key, similar

TAB_REVISION = "Revision"
_COL_NOMBRE = {TAB_RECHAZOS: "Nombre de Cliente", TAB_FINALIZADOS: "NOMBRE"}


def _celda(ref):
    m = re.fullmatch(r"\s*([A-Z]+)(\d+)\s*", str(ref or ""))
    if not m:
        return None
    col = 0
    for ch in m.group(1):
        col = col * 26 + ord(ch) - 64
    return int(m.group(2)), col


def _digitos(v):
    if isinstance(v, float) and v == v and v.is_integer():
        v = int(v)
    return re.sub(r"\D", "", str(v if v is not None else ""))


def _vacio(v):
    return v is None or (isinstance(v, float) and v != v) or str(v).strip() == ""


def _col_rev(base, nombre):
    c = base.col(TAB_REVISION, nombre)
    if not c:
        raise ValueError(f"La pestaña '{TAB_REVISION}' no tiene la columna '{nombre}'.")
    return c


def revisar(base):
    """DataFrame con cada pendiente y si ya se puede dar por resuelto (no escribe nada)."""
    pestanas = set(base.pestanas())
    if TAB_REVISION not in pestanas:
        return pd.DataFrame()
    rev = base.leer(TAB_REVISION)
    if rev.empty:
        return pd.DataFrame()
    col = lambda n: next((c for c in rev.columns if norm(c) == norm(n)), None)
    c_est, c_hoja, c_celda, c_cli, c_campo = col("Estado"), col("Hoja"), col("Celda"), col("Cliente"), col("Campo")
    c_ant, c_des, c_id = col("Antes"), col("Después"), col("ID")
    c_mot = next((c for c in rev.columns if norm(c).startswith("MOTIVO")), None)
    hojas, cache_cat = {}, {}

    def get_cat():
        if "v" not in cache_cat:
            cache_cat["v"] = _catalogo(base)
        return cache_cat["v"]

    filas = []
    for _, r in rev.iterrows():
        estado = str(r[c_est] or "").strip()
        sin_valores = _vacio(r[c_ant]) and _vacio(r[c_des])
        manual = estado in ("Corregido", "No aplica") and sin_valores and _celda(r[c_celda])
        if estado != "Pendiente" and not manual:
            continue
        hoja = str(r[c_hoja] or "").strip()
        out = dict(_fila_rev=int(r["_fila"]), Hoja=hoja, Celda=str(r[c_celda] or ""), Cliente=str(r[c_cli] or ""),
                   Campo=str(r[c_campo] or ""), Problema=str(r[c_mot] or ""), Estado=estado,
                   ID=str(r[c_id] or "").strip() if c_id else "",
                   Resultado="", Detalle="", _marca=None)
        if manual:
            out.update(Resultado="Quitar marca", Detalle=f"Marcado a mano como «{estado}»")
            out["_marca"] = _celda(r[c_celda])
            filas.append(out)
            continue
        if hoja not in pestanas:
            out.update(Resultado="Revisar a mano", Detalle=f"No existe la pestaña «{hoja}»")
            filas.append(out)
            continue
        if hoja not in hojas:
            hojas[hoja] = base.leer(hoja)
        df = hojas[hoja]
        hcol = lambda n: next((c for c in df.columns if norm(c) == norm(canonico(n))), None)
        resuelto, detalle = _evaluar(base, hoja, df, hcol, r[c_celda], out, get_cat)
        out["Resultado"] = {True: "Resuelto", False: "Sigue pendiente", None: "Revisar a mano"}[resuelto]
        out["Detalle"] = detalle
        filas.append(out)
    return pd.DataFrame(filas)


def _catalogo(base):
    if "Catalogos" not in base.pestanas():
        return set()
    cat = base.leer("Catalogos")
    return {norm(v) for v in cat.iloc[:, 0].tolist() if not _vacio(v)} if not cat.empty else set()


def _evaluar(base, hoja, df, hcol, celda_txt, out, catalogo):
    """Devuelve (True resuelto / False pendiente / None no se puede saber, detalle)."""
    campo, problema = out["Campo"], norm(out["Problema"])
    ref = _celda(celda_txt)

    # Pendiente de toda una columna (p. ej. "L2:L33" sin dato de Sale-U)
    if ref is None:
        m = re.fullmatch(r"\s*([A-Z]+)\d+:([A-Z]+)\d+\s*", str(celda_txt or ""))
        cc = hcol(campo)
        nom = hcol(_COL_NOMBRE.get(hoja, ""))
        if m and cc and nom:
            faltan = int(sum(1 for _, x in df.iterrows() if not _vacio(x[nom]) and _vacio(x[cc])))
            return (faltan == 0), (f"Faltan {faltan}" if faltan else "Ya no falta ninguno")
        return None, "Celda no reconocida"

    fila, col_marca = ref
    nom = hcol(_COL_NOMBRE.get(hoja, ""))
    c_id = hcol("ID")
    if out.get("ID") and c_id:                     # el ID no cambia aunque se ordene la hoja
        por_id = df[df[c_id].astype(str).str.strip().str.upper() == out["ID"].upper()]
        if len(por_id) == 1:
            fila = int(por_id.iloc[0]["_fila"])
    fila_df = df[df["_fila"] == fila]
    if nom and (fila_df.empty or not similar(name_key(fila_df.iloc[0][nom]), name_key(out["Cliente"]))):
        # la fila se movió: buscar al cliente por nombre
        cand = df[df[nom].map(lambda v: similar(name_key(v), name_key(out["Cliente"])))] if nom else pd.DataFrame()
        if len(cand) != 1:
            return None, "No encontré al cliente en la fila indicada"
        fila_df = cand
        fila = int(cand.iloc[0]["_fila"])
    out["_marca"] = (fila, col_marca)
    x = fila_df.iloc[0]

    if "TELEFONO / CORREO" in norm(campo):
        tel, cor = hcol("TELEFONO"), hcol("CORREO")
        if "REPETIDO" in problema:
            t = _digitos(x[tel])
            if len(t) != 10:
                return False, "Teléfono vacío o incompleto"
            otros = df[(df[tel].map(_digitos) == t)
                       & (df[nom].map(name_key) != name_key(x[nom]))]
            return (otros.empty, "Teléfono ya no repetido" if otros.empty else "El teléfono sigue repetido con " + ", ".join(otros[nom].astype(str)))
        ok = not _vacio(x[tel]) or not _vacio(x[cor])
        return ok, "Contacto capturado" if ok else "Sigue sin teléfono ni correo"

    campos = [c.strip() for c in campo.split(",") if c.strip()]
    cols = [hcol(c) for c in campos]
    if not campos or any(c is None for c in cols):
        return None, "Campo no reconocido"
    faltan = [c for c, k in zip(campos, cols) if _vacio(x[k])]
    if faltan:
        return False, "Falta: " + ", ".join(faltan)
    if "DIGITOS" in problema:
        t = _digitos(x[cols[0]])
        return (len(t) == 10), (f"Ahora tiene {len(t)} dígitos")
    if "CATALOGO" in problema:
        cat = catalogo()
        ok = norm(x[cols[0]]) in cat
        return ok, ("Está en el catálogo" if ok else f"«{x[cols[0]]}» sigue sin estar en Catalogos")
    if "NO APARECE COMO RECHAZO" in problema:
        return None, "Hay que confirmarlo a mano (cambia el Estado en Revision cuando lo confirmes)"
    return True, "Dato capturado"


def aplicar(base, revision):
    """Quita marcas y actualiza Revision/Historial para las filas Resuelto / Quitar marca."""
    hacer = revision[revision["Resultado"].isin(["Resuelto", "Quitar marca"])]
    if hacer.empty:
        return 0
    c_est, c_des = _col_rev(base, "Estado"), _col_rev(base, "Después")
    hoy = dt.date.today().strftime("%d/%m/%Y")
    marcas, celdas_rev, hist = {}, [], []
    pestanas = set(base.pestanas())
    for _, r in hacer.iterrows():
        if r["_marca"] and r["Hoja"] in pestanas:
            marcas.setdefault(r["Hoja"], []).append(tuple(r["_marca"]))
        if r["Resultado"] == "Resuelto":
            celdas_rev.append((r["_fila_rev"], c_est, "Corregido"))
        celdas_rev.append((r["_fila_rev"], c_des, f"Resuelto, marca quitada por la app ({hoy})"))
        f, c = r["_marca"] if r["_marca"] else (None, None)
        hist.append(dict(Origen="App · limpiar pendientes", Hoja=r["Hoja"],
                         Celda=f"{letra(c)}{f}" if f else r["Celda"], Cliente=r["Cliente"], Campo=r["Campo"],
                         Antes="Pendiente" if r["Resultado"] == "Resuelto" else r["Estado"], Después="Corregido",
                         Motivo=r["Detalle"]))
    for hoja, celdas in marcas.items():
        base.quitar_marca(hoja, celdas)
    base.escribir(TAB_REVISION, celdas_rev)
    base.anotar_historial(hist)
    try:
        archivar(base)          # los resueltos pasan a «Revision resueltos»; en Revision quedan los pendientes
    except Exception:
        pass
    return len(hacer)


# =============================================================================
# ID de cada pendiente y archivo de los resueltos
# =============================================================================
TAB_ARCHIVO = "Revision resueltos"


def ligar_ids(base):
    """Pone en la columna ID de Revision el ID del rechazo de cada pendiente de «MG Colima Rechazos» (así sigue
    apuntando al cliente aunque se ordene la hoja). Solo cuando la fila de la celda es del mismo cliente.
    Devuelve cuántos ligó."""
    from cuadres.ids import COL_ID
    if TAB_REVISION not in base.pestanas() or TAB_RECHAZOS not in base.pestanas():
        return 0
    rev = base.leer(TAB_REVISION)
    if rev.empty:
        return 0
    df = base.leer(TAB_RECHAZOS)
    c_id_r = next((c for c in df.columns if norm(c) == norm(COL_ID)), None)
    c_nom = next((c for c in df.columns if norm(c) == norm(_COL_NOMBRE[TAB_RECHAZOS])), None)
    if not c_id_r or not c_nom:
        return 0
    por_fila = {int(r["_fila"]): r for _, r in df.iterrows()}
    col = lambda n: next((c for c in rev.columns if norm(c) == norm(n)), None)
    c_hoja, c_celda, c_cli, c_id = col("Hoja"), col("Celda"), col("Cliente"), col("ID")
    n_id = base.agregar_columna(TAB_REVISION, COL_ID)
    celdas = []
    for _, r in rev.iterrows():
        if str(r[c_hoja] or "").strip() != TAB_RECHAZOS or (c_id and str(r[c_id] or "").strip()):
            continue
        ref = _celda(r[c_celda])
        x = por_fila.get(ref[0]) if ref else None
        if x is not None and x[c_id_r] and similar(name_key(x[c_nom]), name_key(r[c_cli])):
            celdas.append((int(r["_fila"]), n_id, str(x[c_id_r])))
    base.escribir(TAB_REVISION, celdas)
    return len(celdas)


def archivar(base):
    """Pasa a la pestaña «Revision resueltos» los renglones de Revision que ya no están pendientes (Corregido /
    No aplica con la marca ya quitada). En Revision quedan solo los pendientes. Devuelve cuántos movió."""
    if TAB_REVISION not in base.pestanas():
        return 0
    valores = base._valores(TAB_REVISION)
    if len(valores) < 2:
        return 0
    enc = [str(h or "").strip() for h in valores[0]]
    while enc and not enc[-1]:
        enc.pop()
    idx = lambda n: next((i for i, h in enumerate(enc) if norm(h) == norm(n)), None)
    i_est, i_celda, i_ant, i_des = idx("Estado"), idx("Celda"), idx("Antes"), idx("Después")
    if i_est is None:
        return 0
    quedan, salen = [], []
    for fila in valores[1:]:
        fila = (list(fila) + [None] * len(enc))[:len(enc)]
        if not any(v not in (None, "") for v in fila):
            continue
        estado = str(fila[i_est] or "").strip()
        sin_valores = all(_vacio(fila[i]) for i in (i_ant, i_des) if i is not None)
        falta_quitar_marca = sin_valores and i_celda is not None and _celda(fila[i_celda])
        (salen if estado in ("Corregido", "No aplica") and not falta_quitar_marca else quedan).append(fila)
    if not salen:
        return 0
    previos = []
    if TAB_ARCHIVO in base.pestanas():
        previos = [list(f)[:len(enc)] for f in base._valores(TAB_ARCHIVO)[1:]
                   if any(v not in (None, "") for v in f)]
    base.respaldar(TAB_REVISION)
    base.reemplazar(TAB_ARCHIVO, enc, previos + salen)
    base.dar_formato(TAB_ARCHIVO)
    base.reemplazar(TAB_REVISION, enc, quedan)
    return len(salen)
