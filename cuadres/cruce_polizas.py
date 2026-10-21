# -*- coding: utf-8 -*-
"""
Cruce de «MG Colima Créditos Finalizados» con la pestaña «Polizas» (la escribe la app de seguros:
una fila por crédito con el nombre de la póliza, teléfono —del PDF o de BlueService— y su nota, correo y VIN).

  proponer(base)            → DataFrame de cambios propuestos (no escribe nada)
  aplicar(base, propuestas) → respalda Finalizados, escribe lo marcado y anota el Historial

Cómo se liga cada fila de Finalizados con una póliza (de lo más a lo menos seguro):
  1. VIN igual (si Finalizados tiene columna VIN / SERIE).
  2. Mismo teléfono y nombre parecido (≥ 50 % de las palabras), para no mezclar a un familiar.
  3. Solo nombre: ≥ 80 % de las palabras y una sola póliza con ese parecido.

Qué se propone:
  · TELEFONO vacío o inválido → el de la póliza.                                   (marcado)
  · TELEFONO distinto, ambos válidos → el de la póliza, para que tú decidas.        (sin marcar)
  · NOMBRE incompleto (la póliza trae las mismas palabras y más) → el de la póliza. (marcado)
  · NOMBRE escrito distinto, ligado por VIN o teléfono → el de la póliza.           (marcado solo si fue por VIN)
  · CORREO vacío → el de la póliza.                                                (marcado)
  · VIN vacío (si existe la columna) y ligado por teléfono → el de la póliza.      (marcado)
Un teléfono de la póliza con NOTA TELÉFONO (incompleto, a nombre de otra persona…) nunca se propone.
Lo que desmarques queda como «Ignorado» en el Historial y no se vuelve a proponer.
"""
import re

import pandas as pd

from cuadres.base import TAB_FINALIZADOS, norm, letra
from cuadres.fuentes import to_phone, vacio
from cuadres import saleu as cu

TAB_POLIZAS = "Polizas"
ORIGEN = "App · pólizas"
COLS_VIN = ("VIN", "SERIE", "NUMERO DE SERIE", "NO DE SERIE", "NO. DE SERIE", "NIV")
UMBRAL_SOLO_NOMBRE = 0.80


def _col(df, *nombres):
    objetivos = {norm(n) for n in nombres}
    return next((c for c in df.columns if norm(c) in objetivos), None)


def _vin(v):
    s = re.sub(r"[^A-Z0-9]", "", norm(v))
    return s if len(s) == 17 else ""


def _correo(v):
    s = str(v or "").strip().lower()
    return s if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", s) else ""


def _parecido(a, b):
    return cu.similitud_nombre(cu.normalizar_nombre(a), cu.normalizar_nombre(b))[0]


def leer_polizas(base):
    """Pólizas como lista de dicts (nombre, tel, nota, correo, vin, vehiculo)."""
    if TAB_POLIZAS not in base.pestanas():
        return []
    df = base.leer(TAB_POLIZAS)
    c_nom, c_tel, c_nota = _col(df, "CLIENTE"), _col(df, "TELÉFONO", "TELEFONO"), _col(df, "NOTA TELÉFONO")
    c_cor, c_vin, c_veh = _col(df, "CORREO"), _col(df, "SERIE", "VIN"), _col(df, "VEHÍCULO", "VEHICULO")
    out = []
    for _, r in df.iterrows() if c_nom else []:
        if vacio(r[c_nom]) or not str(r[c_nom]).strip():
            continue
        nota = str(r[c_nota]).strip() if c_nota and not vacio(r[c_nota]) else ""
        out.append(dict(nombre=" ".join(str(r[c_nom]).split()).upper(),
                        tel=None if nota else to_phone(r[c_tel]) if c_tel else None,   # con nota no se usa
                        tel_txt=str(r[c_tel] or "").strip() if c_tel else "", nota=nota,
                        correo=_correo(r[c_cor]) if c_cor else "", vin=_vin(r[c_vin]) if c_vin else "",
                        vehiculo=str(r[c_veh] or "").strip() if c_veh else ""))
    return out


def _ligar(nombre, tel, vin, polizas, por_vin, por_tel):
    """(póliza, cómo) o (None, motivo por el que no se ligó)."""
    if vin and vin in por_vin:
        return por_vin[vin], "VIN"
    if tel:
        cands = [p for p in por_tel.get(tel, []) if _parecido(nombre, p["nombre"]) >= cu.UMBRAL_NOMBRE_CON_TEL]
        if len(cands) == 1:
            return cands[0], "teléfono"
    puntos = sorted(((_parecido(nombre, p["nombre"]), i) for i, p in enumerate(polizas)), reverse=True)
    if not puntos or puntos[0][0] < UMBRAL_SOLO_NOMBRE:
        return None, ""
    if len(puntos) > 1 and puntos[1][0] == puntos[0][0]:
        return None, "más de una póliza con ese nombre"
    if len(cu.normalizar_nombre(nombre).split()) < 2:
        return None, ""
    return polizas[puntos[0][1]], "nombre"


def _ignorados(base):
    from cuadres.actualizar import _ignorados as ign
    try:
        return ign(base)
    except Exception:
        return set()


def proponer(base):
    """DataFrame (Aplicar, Fila, Cliente, Campo, Antes, Después, Ligado por, Motivo, _valor). No escribe nada."""
    cols = ["Aplicar", "Fila", "Cliente", "Campo", "Antes", "Después", "Ligado por", "Motivo", "_valor"]
    if TAB_FINALIZADOS not in base.pestanas():
        raise ValueError(f"No existe la pestaña «{TAB_FINALIZADOS}» en el Sheet.")
    polizas = leer_polizas(base)
    if not polizas:
        raise ValueError(f"No existe la pestaña «{TAB_POLIZAS}» (o está vacía). Carga el Excel en la app de seguros "
                         "con los secrets del Google Sheet para que la cree.")
    por_vin = {p["vin"]: p for p in polizas if p["vin"]}
    por_tel = {}
    for p in polizas:
        if p["tel"]:
            por_tel.setdefault(p["tel"], []).append(p)

    df = base.leer(TAB_FINALIZADOS)
    c_nom, c_tel, c_cor = _col(df, "NOMBRE"), _col(df, "TELEFONO"), _col(df, "CORREO")
    c_vin = _col(df, *COLS_VIN)
    if not c_nom:
        raise ValueError(f"No encuentro la columna NOMBRE en «{TAB_FINALIZADOS}».")
    ignorados = _ignorados(base)
    out = []

    def add(r, campo, antes, despues, motivo, marcado, como):
        if (norm(r[c_nom]), norm(campo), norm(despues)) in ignorados:
            return
        out.append(dict(Aplicar=marcado, Fila=int(r["_fila"]), Cliente=str(r[c_nom]).strip(), Campo=campo,
                        Antes="" if vacio(antes) else str(antes), Después=str(despues), **{"Ligado por": como},
                        Motivo=motivo, _valor=despues))

    for _, r in df.iterrows():
        nombre = r[c_nom]
        if vacio(nombre) or not str(nombre).strip():
            continue
        tel_fin = to_phone(r[c_tel]) if c_tel else None
        vin_fin = _vin(r[c_vin]) if c_vin else ""
        p, como = _ligar(str(nombre), tel_fin, vin_fin, polizas, por_vin, por_tel)
        if p is None:
            continue
        fuerte = como in ("VIN", "teléfono")
        detalle = f"póliza de {p['nombre']}" + (f" · {p['vehiculo']}" if p["vehiculo"] else "")

        # teléfono
        if c_tel and p["tel"]:
            antes = r[c_tel]
            if not tel_fin:
                add(r, c_tel, antes, p["tel"], ("Faltaba el teléfono" if vacio(antes) else
                                                f"Teléfono inválido ({antes})") + f"; {detalle}", True, como)
            elif tel_fin != p["tel"]:
                add(r, c_tel, antes, p["tel"], f"Teléfono distinto al de la {detalle}; confirma cuál es el bueno",
                    False, como)
        # nombre
        n_fin, n_pol = cu.normalizar_nombre(nombre), cu.normalizar_nombre(p["nombre"])
        if n_fin != n_pol and norm(nombre) != norm(p["nombre"]):
            pal_fin, pal_pol = set(n_fin.split()), set(n_pol.split())
            if pal_fin < pal_pol:
                add(r, c_nom, nombre, p["nombre"], f"Nombre incompleto; completo en la {detalle}", True, como)
            elif fuerte and pal_fin != pal_pol:
                add(r, c_nom, nombre, p["nombre"], f"Nombre escrito distinto en la {detalle}", como == "VIN", como)
        # correo
        if c_cor and p["correo"] and not _correo(r[c_cor]):
            add(r, c_cor, r[c_cor], p["correo"], f"Faltaba el correo; {detalle}", vacio(r[c_cor]), como)
        # VIN
        if c_vin and p["vin"] and not vin_fin and como == "teléfono":
            add(r, c_vin, r[c_vin], p["vin"], f"Faltaba el VIN; {detalle}", True, como)

    prop = pd.DataFrame(out, columns=cols)
    return prop.sort_values(["Aplicar", "Fila"], ascending=[False, True]).reset_index(drop=True)


def aplicar(base, propuestas):
    """Escribe lo marcado. Devuelve (escritos, ignorados)."""
    if propuestas is None or propuestas.empty:
        return 0, 0
    marca = propuestas["Aplicar"].astype(bool).values
    elegidas, ignoradas = propuestas[marca], propuestas[~marca]
    celdas, historial = [], []
    if len(elegidas):
        base.respaldar(TAB_FINALIZADOS)
    for _, p in elegidas.iterrows():
        c = base.col(TAB_FINALIZADOS, p["Campo"])
        if not c:
            continue
        celdas.append((int(p["Fila"]), c, p["_valor"]))
        historial.append(dict(Origen=ORIGEN, Hoja=TAB_FINALIZADOS, Celda=f"{letra(c)}{int(p['Fila'])}",
                              Cliente=p["Cliente"], Campo=p["Campo"], Antes=p["Antes"], Después=p["Después"],
                              Motivo=f"{p['Motivo']} (ligado por {p['Ligado por']})"))
    for _, p in ignoradas.iterrows():
        historial.append(dict(Origen="Ignorado", Hoja=TAB_FINALIZADOS, Celda="", Cliente=p["Cliente"],
                              Campo=p["Campo"], Antes=p["Antes"], Después=p["Después"],
                              Motivo="No se aplicó: " + p["Motivo"]))
    base.escribir(TAB_FINALIZADOS, celdas)
    base.anotar_historial(historial)
    return len(celdas), len(ignoradas)
