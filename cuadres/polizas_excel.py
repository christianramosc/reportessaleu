# -*- coding: utf-8 -*-
"""
Actualiza la pestaña «Polizas» del Sheet con tu Excel de pólizas (Polizas_extraidas.xlsx de la app de seguros),
desde la página Actualizar.

  es_excel_polizas(wb)          → True si el libro es el Excel de pólizas (hoja «Contacto» con CLIENTE y SERIE)
  leer(datos)                   → DataFrame: una fila por crédito (hoja Contacto + RFC/MOTOR/CP de «Base de datos»)
  proponer(base, nombre, datos) → dict con las pólizas del Excel, lo que hay hoy en la pestaña y qué cambia
  aplicar(base, prop, conservar)→ respalda la pestaña, la reescribe y (en Google) guarda el Excel para la app de seguros

Cómo se ligan las filas del Excel con las de la pestaña: por SERIE (VIN de 17 caracteres) o, si no hay, por nombre.
  · conservar=True  → las pólizas que ya estaban en la pestaña y no vienen en el Excel se quedan (no se pierde nada).
  · conservar=False → la pestaña queda igual al Excel (como lo hace la app de seguros).

Las columnas calculadas del Excel (ESTATUS, DÍAS QUE QUEDAN…) son fórmulas: se copian con el último valor que guardó
Excel; si el archivo nunca se abrió en Excel vienen vacías y se omiten. La app de seguros las vuelve a calcular.
El cruce con Créditos Finalizados (Mantenimiento) solo usa CLIENTE, TELÉFONO, NOTA TELÉFONO, CORREO, SERIE y VEHÍCULO.
"""
import base64
import datetime as dt
import io
import re

import pandas as pd

from cuadres.base import norm
from cuadres.cruce_polizas import TAB_POLIZAS

HOJA_CONTACTO = "Contacto"
HOJA_BASE = "Base de datos"
RESPALDO_EXCEL = "_Polizas_excel"     # la misma pestaña oculta de donde la app de seguros abre el último Excel
TROZO = 45000                         # caracteres por celda (Google admite hasta 50 000)
# nombres del Excel → nombres que usa la app de seguros en la pestaña
RENOMBRAR = {"ESTATUS HOY": "ESTATUS", "ÚLTIMO DÍA PÓLIZA EXTERNA": "ÚLTIMO DÍA EXTERNA",
             "TIPO DE OPORTUNIDAD": "TIPO"}
EXTRA_BASE = ("RFC", "MOTOR", "CP")
TEXTO = {"TELÉFONO", "SERIE", "MOTOR", "RFC", "FOLIO CRÉDITO", "CP", "PÓLIZA ACTUAL"}   # que Google no los haga número
COL_ACTUALIZADO = "ACTUALIZADO"


def _encabezado(filas):
    """Índice de la fila de encabezados (la que trae CLIENTE y SERIE) o None."""
    for i, f in enumerate(filas):
        t = {norm(v) for v in f if v is not None}
        if "CLIENTE" in t and "SERIE" in t:
            return i
    return None


def es_excel_polizas(wb):
    if HOJA_CONTACTO not in wb.sheetnames:
        return False
    filas = [list(f) for f, _ in zip(wb[HOJA_CONTACTO].iter_rows(values_only=True), range(15))]
    return _encabezado(filas) is not None


def _limpio(v):
    if v is None or (isinstance(v, float) and v != v):
        return ""
    if isinstance(v, dt.datetime):
        return v.date() if v.time() == dt.time() else v
    if isinstance(v, dt.date):
        return v
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    s = str(v).strip()
    return "" if s.startswith("=") else s.removesuffix(".0") if isinstance(v, str) else v


def _tabla(ws):
    filas = [list(f) for f in ws.iter_rows(values_only=True)]
    h = _encabezado(filas[:15]) if ws.title == HOJA_CONTACTO else 0
    if h is None or not filas:
        return pd.DataFrame()
    enc = [str(c).strip() if c is not None else "" for c in filas[h]]
    datos = [f + [None] * (len(enc) - len(f)) for f in filas[h + 1:]]
    df = pd.DataFrame([f[:len(enc)] for f in datos], columns=enc, dtype=object)
    return df.loc[:, [c for c in df.columns if c]]


def leer(datos):
    """Excel de pólizas → DataFrame con una fila por crédito (valores, no fórmulas)."""
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(datos), read_only=True, data_only=True)
    try:
        if not es_excel_polizas(wb):
            raise ValueError("El Excel no es el de pólizas: falta la hoja «Contacto» con las columnas CLIENTE y SERIE.")
        df = _tabla(wb[HOJA_CONTACTO])
        bd = _tabla(wb[HOJA_BASE]) if HOJA_BASE in wb.sheetnames else pd.DataFrame()
    finally:
        wb.close()
    df = df.loc[:, ~df.columns.duplicated()].rename(columns=RENOMBRAR)
    df = df.apply(lambda c: c.map(_limpio))
    df = df[df["CLIENTE"].astype(str).str.strip() != ""].reset_index(drop=True)
    df["CLIENTE"] = df["CLIENTE"].map(lambda v: " ".join(str(v).split()).upper())
    # RFC, motor y CP solo vienen en «Base de datos»: se toman por SERIE
    if not bd.empty and "SERIE" in bd.columns:
        bd = bd.apply(lambda c: c.map(_limpio))
        bd = bd[bd["SERIE"].astype(str) != ""].drop_duplicates("SERIE").set_index("SERIE")
        for c in EXTRA_BASE:
            if c in bd.columns and c not in df.columns:
                df[c] = df["SERIE"].map(bd[c]).fillna("")
    # columnas calculadas sin valor guardado (el Excel nunca se abrió en Excel): fuera
    vacias = [c for c in df.columns if c != "CLIENTE" and (df[c].astype(str).str.strip() == "").all()]
    return df.drop(columns=vacias)


def _vin(v):
    s = re.sub(r"[^A-Z0-9]", "", norm(v))
    return s if len(s) == 17 else ""


def _clave(r):
    vin = _vin(r.get("SERIE", ""))
    return f"VIN {vin}" if vin else f"NOMBRE {norm(r.get('CLIENTE', ''))}"


def _actual(base):
    if TAB_POLIZAS not in base.pestanas():
        return pd.DataFrame()
    df = base.leer(TAB_POLIZAS).drop(columns="_fila", errors="ignore")
    if df.empty or "CLIENTE" not in df.columns:
        return pd.DataFrame()
    df = df[df["CLIENTE"].astype(str).str.strip().isin(["", "nan", "None"]) == False]   # noqa: E712
    return df.reset_index(drop=True)


def proponer(base, nombre, datos):
    """Qué pasaría al cargar el Excel. No escribe nada."""
    nuevo = leer(datos)
    if nuevo.empty:
        raise ValueError(f"«{nombre}» no trae pólizas en la hoja Contacto.")
    actual = _actual(base)
    k_nuevo = nuevo.apply(_clave, axis=1)
    k_actual = actual.apply(_clave, axis=1) if not actual.empty else pd.Series(dtype=str)
    ya = set(k_actual)
    altas = nuevo[~k_nuevo.isin(ya)]
    solo_pestana = actual[~k_actual.isin(set(k_nuevo))] if not actual.empty else actual
    con_tel = int((nuevo.get("TELÉFONO", pd.Series([""] * len(nuevo))).astype(str).str.strip() != "").sum())
    return dict(nombre=nombre, datos=datos, df=nuevo, actual=actual, altas=altas, solo_pestana=solo_pestana,
                existia=TAB_POLIZAS in base.pestanas(), con_tel=con_tel)


def _celda(col, v, google):
    if v is None or (not isinstance(v, str) and pd.isna(v)):
        return ""
    if isinstance(v, (dt.date, dt.datetime, pd.Timestamp)):
        return pd.Timestamp(v).strftime("%Y-%m-%d")      # ISO: Google lo toma como fecha en cualquier idioma
    if col in TEXTO:
        s = str(v).strip().removesuffix(".0")
        return ("'" + s) if (s and google) else s
    if hasattr(v, "item"):
        return v.item()
    return v


def _guardar_excel_para_seguros(base, datos):
    """Deja el Excel en la pestaña oculta que lee la app de seguros al abrir (mismo formato: base64 en trozos)."""
    from gspread.utils import ValueInputOption
    texto = base64.b64encode(datos).decode()
    filas = [[texto[i:i + TROZO]] for i in range(0, len(texto), TROZO)]
    try:
        ws = base.libro.worksheet(RESPALDO_EXCEL)
        ws.clear()
        ws.resize(rows=len(filas), cols=1)
    except base._gs.WorksheetNotFound:
        ws = base.libro.add_worksheet(RESPALDO_EXCEL, rows=len(filas), cols=1)
        ws.hide()
    ws.update(filas, "A1", value_input_option=ValueInputOption.raw)    # raw: que Google no lo tome como fórmula
    base._hojas = None


def aplicar(base, prop, conservar=True):
    """Reescribe la pestaña Polizas. Devuelve dict(filas, altas, conservadas, excel_seguros, aviso)."""
    from cuadres.base import BaseGoogle
    google = isinstance(base, BaseGoogle)
    nuevo, quedan = prop["df"], prop["solo_pestana"] if conservar else pd.DataFrame()
    cols = list(nuevo.columns)
    cols += [c for c in quedan.columns if c not in cols and c != COL_ACTUALIZADO]
    ahora = dt.datetime.now().strftime("%Y-%m-%d %H:%M")
    filas = [[_celda(c, r.get(c), google) for c in cols] + [ahora] for r in nuevo.to_dict("records")]
    filas += [[_celda(c, r.get(c), google) for c in cols] + [r.get(COL_ACTUALIZADO, "")]
              for r in quedan.to_dict("records")]
    if prop["existia"] and google:
        base.respaldar(TAB_POLIZAS)
    base.reemplazar(TAB_POLIZAS, cols + [COL_ACTUALIZADO], filas)
    try:
        base.congelar(TAB_POLIZAS, filas=1)
    except Exception:
        pass
    # la app de seguros abre el último Excel guardado y vuelve a llenar la pestaña: que abra este
    excel_seguros, aviso = False, None
    if google:
        if len(quedan):
            aviso = (f"Se conservaron {len(quedan)} póliza(s) que no vienen en «{prop['nombre']}». La app de seguros "
                     "no se actualizó con este Excel (al abrirla volvería a dejar la pestaña sin esas pólizas): "
                     "súbelo también allá cuando tenga a todos.")
        else:
            try:
                _guardar_excel_para_seguros(base, prop["datos"])
                excel_seguros = True
            except Exception as e:
                aviso = f"La pestaña Polizas quedó actualizada, pero no se pudo guardar el Excel para la app de seguros ({e})."
    base.anotar_historial([dict(Origen="App · pólizas", Hoja=TAB_POLIZAS, Celda="", Cliente="",
                                Campo="Pestaña completa", Antes=f"{len(prop['actual'])} póliza(s)",
                                Después=f"{len(filas)} póliza(s)",
                                Motivo=f"Excel de pólizas «{prop['nombre']}»: {len(prop['altas'])} nueva(s)"
                                       + (f", {len(quedan)} conservada(s)" if len(quedan) else ""))])
    return dict(filas=len(filas), altas=len(prop["altas"]), conservadas=len(quedan), excel_seguros=excel_seguros,
                aviso=aviso)
