# -*- coding: utf-8 -*-
"""
Qué archivos se han subido y cuándo, para avisar qué falta.

  anotar(base, cargas)  → agrega renglones a la pestaña oculta «Cargas» (Fecha de carga, Tipo, Archivo, Periodo,
                          Registros); conserva los últimos 300.
  estado(base, hoy)     → una fila por tipo de archivo con lo último que hay y un aviso si hace falta algo:
      · Cetelem: meses guardados (pestaña Capturas Cetelem); avisa si falta la captura del mes.
      · F&I: fecha de la última carga (Bitacora F&I).
      · Sale-U Crédito: última carga y su periodo (pestaña Cargas); avisa si tiene más de 7 días.
      · Sale-U Fuente: leads guardados y la última alta (Fuente Sale-U); avisa si tiene más de 7 días.
"""
import datetime as dt

import pandas as pd

from cuadres.base import TAB_CAPTURAS, TAB_FI, norm

TAB_CARGAS = "Cargas"
COLS = ["Fecha de carga", "Tipo", "Archivo", "Periodo", "Registros"]
MAX_FILAS = 300
DIAS_VIEJO = 7


def _ahora():
    return dt.datetime.now().replace(microsecond=0)


def anotar(base, cargas):
    """cargas: [dict(Tipo, Archivo, Periodo, Registros)]."""
    if not cargas:
        return
    previas = []
    if TAB_CARGAS in base.pestanas():
        df = base.leer(TAB_CARGAS).drop(columns="_fila", errors="ignore")
        previas = [[r.get(c) for c in COLS] for _, r in df.iterrows()]
    nuevas = [[_ahora(), c.get("Tipo", ""), c.get("Archivo", ""), c.get("Periodo", ""), c.get("Registros", "")]
              for c in cargas]
    base.reemplazar(TAB_CARGAS, COLS, (previas + nuevas)[-MAX_FILAS:], oculta=True)


def _a_fecha(v):
    if isinstance(v, dt.datetime):
        return v
    if isinstance(v, dt.date):
        return dt.datetime(v.year, v.month, v.day)
    f = pd.to_datetime(v, errors="coerce", dayfirst=False)
    return None if f is pd.NaT or f != f else f.to_pydatetime()


def _ultima(base, tipo):
    if TAB_CARGAS not in base.pestanas():
        return None
    df = base.leer(TAB_CARGAS)
    if df.empty or "Tipo" not in df.columns:
        return None
    df = df[df["Tipo"].map(norm) == norm(tipo)]
    return df.iloc[-1] if len(df) else None


def _hace(f, hoy):
    if not f:
        return None
    d = (hoy - f.date()).days
    return "hoy" if d <= 0 else ("ayer" if d == 1 else f"hace {d} días")


def estado(base, hoy=None):
    """Lista de dicts: Archivo, Último, Detalle, Nivel («ok» / «aviso»), Aviso."""
    from cuadres.fuentes import ETIQUETA_MES, etiqueta_hoja
    hoy = hoy or dt.date.today()
    out = []
    pestanas = base.pestanas()

    # Cetelem
    meses, f_cet = [], None
    if TAB_CAPTURAS in pestanas:
        cap = base.leer(TAB_CAPTURAS)
        if not cap.empty and "Hoja" in cap.columns:
            meses = list(dict.fromkeys(etiqueta_hoja(v) for v in cap["Hoja"] if etiqueta_hoja(v)))
        if "Fecha de carga" in cap.columns:
            fechas = [f for f in (_a_fecha(v) for v in cap["Fecha de carga"]) if f]
            f_cet = max(fechas) if fechas else None
    del_mes = f"{ETIQUETA_MES[hoy.month]} {hoy.year}"
    falta = norm(del_mes) not in {norm(m) for m in meses}
    out.append(dict(Archivo="Captura de Cetelem", Último=meses[-1] if meses else "—",
                    Detalle=(f"{len(meses)} mes(es) guardados" + (f" · subida {_hace(f_cet, hoy)}" if f_cet else "")),
                    Nivel="aviso" if falta and hoy.day > 5 else "ok",
                    Aviso=f"Falta la captura de {del_mes.capitalize()}" if falta and hoy.day > 5 else ""))

    # F&I
    f_fi = None
    if TAB_FI in pestanas:
        fi = base.leer(TAB_FI)
        if not fi.empty and "Fecha de carga" in fi.columns:
            fechas = [f for f in (_a_fecha(v) for v in fi["Fecha de carga"]) if f]
            f_fi = max(fechas) if fechas else None
    viejo = not f_fi or (hoy - f_fi.date()).days > DIAS_VIEJO
    out.append(dict(Archivo="Bitácora de F&I", Último=f"{f_fi:%d/%m/%Y}" if f_fi else "—",
                    Detalle=f"subida {_hace(f_fi, hoy)}" if f_fi else "sin cargas",
                    Nivel="aviso" if viejo else "ok",
                    Aviso=("La Bitácora de F&I tiene más de una semana" if f_fi else "No hay Bitácora de F&I")
                    if viejo else ""))

    # Sale-U Crédito
    u = _ultima(base, "credito")
    f_cr = _a_fecha(u["Fecha de carga"]) if u is not None else None
    viejo = not f_cr or (hoy - f_cr.date()).days > DIAS_VIEJO
    out.append(dict(Archivo="Sale-U · Crédito", Último=f"{f_cr:%d/%m/%Y}" if f_cr else "—",
                    Detalle=(f"periodo {u['Periodo']} · subido {_hace(f_cr, hoy)}" if u is not None else "sin cargas"),
                    Nivel="aviso" if viejo else "ok",
                    Aviso=(f"El reporte de Crédito tiene {(hoy - f_cr.date()).days} días" if f_cr
                           else "Sube el reporte de Crédito de Sale-U") if viejo else ""))

    # Sale-U Fuente
    from cuadres.saleu_fuente import guardada
    fu = guardada(base)
    u = _ultima(base, "fuente")
    f_fu = _a_fecha(u["Fecha de carga"]) if u is not None else None
    if f_fu is None and fu is not None and fu.attrs.get("rango"):
        f_fu = fu.attrs["rango"][1]
    viejo = not f_fu or (hoy - f_fu.date()).days > DIAS_VIEJO
    out.append(dict(Archivo="Sale-U · Fuente", Último=f"{f_fu:%d/%m/%Y}" if f_fu else "—",
                    Detalle=(f"{len(fu)} leads guardados" if fu is not None else "sin leads guardados")
                    + (f" · subido {_hace(f_fu, hoy)}" if f_fu else ""),
                    Nivel="aviso" if viejo else "ok",
                    Aviso=(f"El reporte Fuente tiene {(hoy - f_fu.date()).days} días" if f_fu
                           else "Sube el reporte Fuente de Sale-U (de enero a la fecha la primera vez)")
                    if viejo else ""))
    return out
