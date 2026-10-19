# -*- coding: utf-8 -*-
"""
Reporte «Fuente» de Sale-U (ReporteFuentes_Agencia_MG-Colima_…csv): todos los leads de MG, con cualquier
estatus. Es el soporte del reporte de créditos (reporteLeadsconCredito):

  · el cliente está en el reporte de Crédito            → «Registrado»
  · está en Fuente pero no en Crédito                    → «Falta crédito» (existe como lead; falta registrar el crédito)
  · no está en ninguno de los dos                        → «No Registrado» (falta darlo de alta en Sale-U)

Búsqueda de un cliente en Fuente (igual que en el de Crédito): por teléfono o por nombre; basta con que coincida
uno de los dos. Por nombre se pide un nombre muy parecido con al menos 3 palabras iguales (Fuente trae miles de
leads con nombres cortos: con 2 palabras salen falsos parecidos).
"""
import csv
import datetime as dt
import io
import re

import pandas as pd

from cuadres import saleu as cu

MIN_PALABRAS_SOLO_NOMBRE = 3


def _texto(ruta_o_bytes):
    datos = ruta_o_bytes if isinstance(ruta_o_bytes, bytes) else open(ruta_o_bytes, "rb").read()
    for enc in ("utf-8-sig", "latin-1"):
        try:
            return datos.decode(enc)
        except UnicodeDecodeError:
            continue
    return datos.decode("utf-8", "ignore")


def _filas(ruta):
    if str(ruta).lower().endswith((".xlsx", ".xls")):
        crudo = pd.read_excel(ruta, header=None, dtype=str).fillna("")
        return [[str(x) for x in f] for f in crudo.values.tolist()]
    return list(csv.reader(io.StringIO(_texto(ruta))))


def _encabezado(filas):
    """Índice de la fila de títulos del reporte Fuente (Fecha de Alta, Nombre, Telefono…) o None."""
    for i, f in enumerate(filas[:15]):
        t = [cu.quitar_acentos(x).strip().upper() for x in f]
        if "FECHA DE ALTA" in t and "NOMBRE" in t and "TELEFONO" in t:
            return i
    return None


def es_fuente(ruta):
    """True si el archivo es el reporte Fuente (no el de Leads con Crédito)."""
    try:
        return _encabezado(_filas(ruta)) is not None
    except Exception:
        return False


def leer_fuente(ruta):
    """DataFrame del reporte Fuente: Nombre, Telefono, Asesor, Estatus_Lead, Fuente, Fecha_Alta (texto), _fecha,
    Tel_norm, Nombre_norm. attrs["rango"] = (primera, última fecha de alta)."""
    filas = _filas(ruta)
    h = _encabezado(filas)
    if h is None:
        raise ValueError("No es el reporte Fuente de Sale-U (faltan los títulos «Fecha de Alta», «Nombre», "
                         "«Telefono»).")
    enc = [x.strip() for x in filas[h]]
    i_com = next((i for i, x in enumerate(enc) if cu.quitar_acentos(x).upper() == "COMENTARIO"), None)
    datos = []
    for f in filas[h + 1:]:
        if not "".join(f).strip():
            continue
        if len(f) > len(enc) and i_com is not None:      # comas sueltas dentro del comentario
            extra = len(f) - len(enc)
            f = f[:i_com] + [",".join(f[i_com:i_com + extra + 1])] + f[i_com + extra + 1:]
        f = (f + [""] * len(enc))[:len(enc)]
        datos.append(f)
    df = pd.DataFrame(datos, columns=enc)
    col = lambda n: next((c for c in enc if cu.quitar_acentos(c).strip().upper() == n), None)
    out = pd.DataFrame({
        "Nombre": df[col("NOMBRE")].str.strip(),
        "Telefono": df[col("TELEFONO")].str.strip(),
        "Asesor": df[col("ASESOR")].str.strip() if col("ASESOR") else "",
        "Estatus_Lead": df[col("ESTATUS")].str.strip() if col("ESTATUS") else "",
        "Fuente": df[col("FUENTE")].str.strip() if col("FUENTE") else "",
        "Fecha_Alta": df[col("FECHA DE ALTA")].str.strip(),
    })
    return _completar(out)


def _completar(out):
    """Columnas de trabajo (_fecha, Tel_norm, Nombre_norm) y el rango de fechas de alta."""
    out = out.copy()
    fecha = pd.to_datetime(out["Fecha_Alta"].map(lambda v: v if isinstance(v, (pd.Timestamp, dt.datetime)) else
                                                  str(v or "").strip()), errors="coerce")
    out["_fecha"] = fecha
    out["Fecha_Alta"] = [f.strftime("%Y-%m-%d %H:%M:%S") if f == f and f is not pd.NaT else str(v or "")
                         for f, v in zip(fecha, out["Fecha_Alta"])]
    for c in ("Nombre", "Telefono", "Asesor", "Estatus_Lead", "Fuente"):
        out[c] = out[c].map(lambda v: "" if v is None or (isinstance(v, float) and v != v) else str(v).strip())
    out["Tel_norm"] = out["Telefono"].map(cu.normalizar_tel)
    out["Nombre_norm"] = out["Nombre"].map(cu.normalizar_nombre)
    out = out[(out["Tel_norm"] != "") | (out["Nombre_norm"] != "")].reset_index(drop=True)
    f = out["_fecha"].dropna()
    out.attrs["rango"] = (f.min().to_pydatetime(), f.max().to_pydatetime()) if len(f) else None
    return out


# =============================================================================
# Leads guardados en el Sheet (pestaña oculta «Fuente Sale-U»)
# =============================================================================
TAB_FUENTE = "Fuente Sale-U"
COLS_TAB = ["Fecha de Alta", "Nombre", "Telefono", "Asesor", "Fuente", "Estatus"]
_DE_TAB = {"Fecha de Alta": "Fecha_Alta", "Nombre": "Nombre", "Telefono": "Telefono", "Asesor": "Asesor",
           "Fuente": "Fuente", "Estatus": "Estatus_Lead"}


def guardada(base):
    """Leads del reporte Fuente que ya se guardaron en el Sheet (o None)."""
    if base is None or TAB_FUENTE not in base.pestanas():
        return None
    df = base.leer(TAB_FUENTE)
    if df.empty:
        return None
    col = lambda n: next((c for c in df.columns if cu.quitar_acentos(c).strip().upper() == cu.quitar_acentos(n).upper()),
                         None)
    out = pd.DataFrame({v: (df[col(k)] if col(k) else "") for k, v in _DE_TAB.items()})
    return _completar(out)


def combinar(*partes):
    """Une reportes Fuente (lo guardado + lo que se subió): un lead por teléfono, nombre y fecha de alta; si se repite
    se queda el más reciente del último reporte."""
    partes = [p for p in partes if p is not None and len(p)]
    if not partes:
        return None
    todo = pd.concat(partes, ignore_index=True)
    todo["_k"] = todo["Tel_norm"] + "|" + todo["Nombre_norm"] + "|" + todo["Fecha_Alta"].str[:16]
    todo = todo.drop_duplicates("_k", keep="last").drop(columns="_k")
    return _completar(todo.sort_values("_fecha", ascending=False, na_position="last")
                      [["Nombre", "Telefono", "Asesor", "Estatus_Lead", "Fuente", "Fecha_Alta"]])


def guardar(base, fu):
    """Guarda los leads en la pestaña oculta «Fuente Sale-U» (la reescribe con todo lo acumulado). Devuelve cuántos."""
    if fu is None or not len(fu):
        return 0
    filas = [[r["Fecha_Alta"], r["Nombre"], r["Telefono"], r["Asesor"], r["Fuente"], r["Estatus_Lead"]]
             for _, r in fu.iterrows()]
    base.reemplazar(TAB_FUENTE, COLS_TAB, filas, oculta=True)
    return len(filas)


def periodo_txt(fu):
    r = fu.attrs.get("rango") if fu is not None else None
    return f"{r[0]:%d/%m/%Y} – {r[1]:%d/%m/%Y}" if r else ""


class Buscador:
    """Busca clientes en el reporte Fuente (índices por teléfono y por palabra para que sea rápido)."""

    def __init__(self, fu):
        self.fu = fu
        self.por_tel, self.por_palabra = {}, {}
        for i, r in fu.iterrows():
            if r["Tel_norm"]:
                self.por_tel.setdefault(r["Tel_norm"], []).append(i)
            for p in set(r["Nombre_norm"].split()):
                self.por_palabra.setdefault(p[:4], set()).add(i)

    def buscar(self, nombre, tel=""):
        """Lead por teléfono o por nombre: (fila del reporte Fuente, cómo) o (None, "")."""
        nom = cu.normalizar_nombre(nombre)
        tel = cu.normalizar_tel(tel) if tel else ""
        mismos = self.por_tel.get(tel, []) if tel else []
        if mismos:
            i = max(mismos, key=lambda j: cu.similitud_nombre(nom, self.fu.at[j, "Nombre_norm"]))
            r = cu.similitud_nombre(nom, self.fu.at[i, "Nombre_norm"])[0]
            return self.fu.loc[i], "teléfono" if r >= cu.UMBRAL_NOMBRE_CON_TEL else "teléfono (nombre distinto)"
        i = self._por_nombre(nom)
        return (self.fu.loc[i], "nombre (teléfono distinto)") if i is not None else (None, "")

    def _por_nombre(self, nom):
        palabras = nom.split()
        if len(palabras) < MIN_PALABRAS_SOLO_NOMBRE:
            return None
        cand = {}
        for p in palabras:
            for i in self.por_palabra.get(p[:4], ()):
                cand[i] = cand.get(i, 0) + 1
        mejor = (0, 0, None)
        for i, n in cand.items():
            if n < MIN_PALABRAS_SOLO_NOMBRE:
                continue
            r, k = cu.similitud_nombre(nom, self.fu.at[i, "Nombre_norm"])
            if (r, k) > mejor[:2]:
                mejor = (r, k, i)
        if mejor[0] >= cu.UMBRAL_NOMBRE_FUERTE and mejor[1] >= MIN_PALABRAS_SOLO_NOMBRE:
            return mejor[2]
        return None


def alta_txt(fila):
    """Fecha de alta del lead como dd/mm/aaaa."""
    f = fila.get("_fecha")
    return f.strftime("%d/%m/%Y") if f is not None and f == f else str(fila.get("Fecha_Alta", ""))[:10]


COL_SALEU = "Cómo está en Sale-U"
NO_APARECE = "No se encontró por nombre ni por teléfono"


def _limpio(v):
    return " ".join(str(v or "").split())


def por_que(como):
    """Texto corto de la coincidencia: «nombre y teléfono», «teléfono» o «nombre»."""
    c = cu.quitar_acentos(str(como or "")).lower()
    if not c:
        return ""
    if "distinto" in c or "otro" in c:
        return "teléfono" if c.startswith("telefono") else "nombre"
    if "telefono" in c:
        return "nombre y teléfono"
    return "nombre"


def describir(nombre, como="", tel=""):
    """Lo que va en «Cómo está en Sale-U»: Coincidió por … · Se encuentra como «…». Si coincidió solo por uno de
    los dos datos, también el teléfono con que está en Sale-U (para corregir el que no cuadra)."""
    motivo = por_que(como)
    partes = []
    if motivo:
        partes.append(f"Coincidió por: {motivo}")
    texto = f"Se encuentra como: «{_limpio(nombre)}»"
    if motivo in ("nombre", "teléfono"):
        texto += f" con teléfono: {cu.normalizar_tel(tel) or _limpio(tel) or 'sin teléfono'}"
    partes.append(texto)
    return " · ".join(partes)


def describir_lead(fila, como=""):
    """Cómo está el lead en el reporte Fuente (por qué coincidió, con qué nombre y teléfono está)."""
    return describir(fila.get("Nombre"), como, fila.get("Telefono"))


def describir_credito(nombre, tel="", asesor="", estatus="", financiera="", alta="", como=""):
    """Cómo está el cliente en el reporte de Crédito de Sale-U (por qué coincidió, con qué nombre y teléfono está)."""
    return describir(nombre, como, tel)


def tel_como_texto(v):
    """Teléfono del Sheet (puede venir como número 3121234567.0) a texto de dígitos."""
    if v is None or (isinstance(v, float) and v != v):
        return ""
    s = str(v)
    if re.fullmatch(r"\d+\.0", s):
        s = s[:-2]
    return s
