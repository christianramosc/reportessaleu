"""
CUADRE DE ARCHIVOS v2 — App de Streamlit
=========================================
Cuadra (reconcilia) dos archivos Excel o CSV por una llave (simple o compuesta),
compara las columnas que elijas y genera observaciones y un Excel de trabajo.

Novedades v2
  - Llave compuesta (varias columnas) y normalización de llave (solo letras y
    números, solo dígitos, ignorar ceros a la izquierda, últimos N caracteres).
  - Diagnóstico de la llave antes de cuadrar (vacías, repetidas, cobertura).
  - Duplicados emparejados por orden de aparición (1.ª con 1.ª, 2.ª con 2.ª…).
  - Posibles errores de captura: llaves sin cruce que se parecen a otra
    (ej. un VIN con un carácter cambiado), indicando la posición que difiere.
  - Modos de comparación: Igual, Contiene, Empieza con y Similar (% mínimo).
    Tolerancia para números (absoluta) y fechas (días).
  - Guardar / cargar la configuración en JSON para cuadres recurrentes.
  - Excel con fórmulas vivas: Estatus, Resultado por columna, Diferencias,
    Veces que aparece la llave, Prioridad, Resumen e indicadores se recalculan
    si corriges un valor o cambias una tolerancia. Colores por estatus,
    columna de Revisión con lista desplegable, hojas de Observaciones,
    Posibles coincidencias, Duplicados, Datos de origen, Configuración y Leyenda.

Ejecutar sola:        streamlit run cuadre_app.py
Como página:          copiar a pages/Cuadre_de_archivos.py (funciona tal cual)
Desde un menú propio: from cuadre_app import main as cuadre_main; cuadre_main()

Requisitos: streamlit, pandas, openpyxl  (xlrd solo para archivos .xls)
"""

import datetime
import difflib
import hashlib
import io
import json
import re
import unicodedata

import pandas as pd
import streamlit as st
from openpyxl import Workbook
from openpyxl.cell.cell import ILLEGAL_CHARACTERS_RE
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.chart.label import DataLabelList
from openpyxl.chart.series import DataPoint
from openpyxl.comments import Comment
from openpyxl.formatting.rule import CellIsRule, FormulaRule
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
from openpyxl.workbook.properties import CalcProperties
from openpyxl.worksheet.datavalidation import DataValidation

P = "cq_"  # prefijo de session_state / widgets para no chocar con otras apps

TIPOS = ["Texto flexible", "Texto exacto", "Número", "Fecha"]
MODOS = ["Igual", "Contiene", "Empieza con", "Similar"]
NORMS_LLAVE = ["Estándar", "Solo letras y números", "Solo dígitos"]
REVISION = ["Pendiente", "En revisión", "Corregido", "Justificado", "No requiere"]
R_OK, R_NO, R_VACIO, R_AMBOS, R_NA = "✔", "✘", "Vacío", "—", "N/A"

ORDEN = ["ok", "dif", "inc", "dup", "solo_a", "solo_b", "sin"]
INFO = {  # color, prioridad, significado, acción sugerida
    "ok": ("C6EFCE", "Baja", "La llave existe en ambos archivos y todas las columnas comparadas coinciden.",
           "Ninguna."),
    "dif": ("FFD8B1", "Alta", "La llave cruza pero al menos una columna comparada no coincide.",
            "Revisar cuál archivo tiene el dato correcto y corregir el otro."),
    "inc": ("FFF2CC", "Media", "La llave cruza y no hay diferencias, pero algún dato está vacío en un archivo.",
            "Completar el dato faltante."),
    "dup": ("E4DFEC", "Media", "La llave aparece más de una vez en algún archivo; se emparejó por orden de aparición.",
            "Confirmar si es un registro repetido o dos registros distintos."),
    "solo_a": ("FFC7CE", "Alta", "Está en el archivo A y no se encontró en el B.",
               "Verificar si falta en B o si la llave tiene un error (ver Posibles coincidencias)."),
    "solo_b": ("E6B8B7", "Alta", "Está en el archivo B y no se encontró en el A.",
               "Verificar si falta en A o si la llave tiene un error (ver Posibles coincidencias)."),
    "sin": ("D9D9D9", "Media", "El registro no tiene valor en la(s) columna(s) llave.",
            "Capturar la llave para poder cruzarlo."),
}
COLORES_OBS = {
    "Diferencia": "FFD8B1", "Dato vacío": "FFF2CC", "No encontrado": "FFC7CE",
    "Duplicado sin par": "FFC7CE", "Duplicado": "E4DFEC", "Posible error de captura": "BDD7EE",
    "Sin llave": "D9D9D9",
}


def etiquetas(nombre_a, nombre_b):
    return {"ok": "Coincide", "dif": "Con diferencias", "inc": "Incompleto", "dup": "Duplicado",
            "solo_a": f"Solo en {nombre_a}", "solo_b": f"Solo en {nombre_b}", "sin": "Sin llave"}


# ---------------------------------------------------------------------------
# Normalización
# ---------------------------------------------------------------------------

def quitar_acentos(s):
    return "".join(c for c in unicodedata.normalize("NFKD", s) if not unicodedata.combining(c))


def es_vacio(v):
    if v is None:
        return True
    try:
        return bool(pd.isna(v))
    except (TypeError, ValueError):
        return False


def norm_llave(v, modo="Estándar", sin_ceros=False, ultimos=0):
    if es_vacio(v):
        return ""
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)):
        v = v.strftime("%Y-%m-%d")
    s = re.sub(r"\s+", " ", quitar_acentos(str(v)).upper()).strip()
    if modo == "Solo letras y números":
        s = re.sub(r"[^A-Z0-9]", "", s)
    elif modo == "Solo dígitos":
        s = re.sub(r"[^0-9]", "", s)
    if sin_ceros and s:
        s = s.lstrip("0") or "0"
    if ultimos and ultimos > 0:
        s = s[-int(ultimos):]
    return s


def construir_llave(df, cols, modo="Estándar", sin_ceros=False, ultimos=0):
    """Llave normalizada por fila. Compuesta = partes unidas con '|'; si falta una parte, queda vacía."""
    partes = [df[c].map(lambda v: norm_llave(v, modo, sin_ceros, ultimos)) for c in cols]
    if len(partes) == 1:
        return partes[0]
    llave = partes[0]
    for p in partes[1:]:
        llave = llave + "|" + p
    incompleta = pd.concat(partes, axis=1).eq("").any(axis=1)
    return llave.where(~incompleta, "")


def norm_valor(v, tipo):
    if es_vacio(v):
        return None
    if tipo == "Texto exacto":
        s = str(v).strip()
        return s or None
    if tipo == "Número":
        if isinstance(v, (int, float)) and not isinstance(v, bool):
            return float(v)
        try:
            return float(str(v).replace(",", "").replace("$", "").replace(" ", ""))
        except ValueError:
            return None
    if tipo == "Fecha":
        if isinstance(v, (pd.Timestamp, datetime.datetime)):
            return v.date()
        if isinstance(v, datetime.date):
            return v
        texto = str(v).strip()
        try:
            iso = re.match(r"^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}", texto)
            f = pd.to_datetime(texto, yearfirst=True) if iso else pd.to_datetime(texto, dayfirst=True)
            return None if pd.isna(f) else f.date()
        except (ValueError, TypeError, OverflowError):
            return None
    # Texto flexible
    if isinstance(v, float) and v.is_integer():
        v = int(v)
    if isinstance(v, (pd.Timestamp, datetime.datetime)):
        v = v.strftime("%Y-%m-%d")
    s = re.sub(r"[^A-Z0-9]", "", quitar_acentos(str(v)).upper())
    s = re.sub(r"(?<![0-9])0+(?=[0-9])", "", s)
    return s or None


def norm_nombre(c):
    return re.sub(r"[^a-z0-9]", "", quitar_acentos(str(c)).lower())


def inferir_tipo(serie):
    s = serie.dropna().head(200)
    if s.empty:
        return "Texto flexible"
    if all(isinstance(v, (pd.Timestamp, datetime.datetime, datetime.date)) for v in s):
        return "Fecha"
    if all(isinstance(v, (int, float)) and not isinstance(v, bool) for v in s):
        return "Número"
    return "Texto flexible"


def fmt(v):
    """Valor legible para observaciones."""
    if es_vacio(v):
        return "(vacío)"
    if isinstance(v, (pd.Timestamp, datetime.datetime)):
        return v.strftime("%d/%m/%Y") if (v.hour, v.minute, v.second) == (0, 0, 0) else v.strftime("%d/%m/%Y %H:%M")
    if isinstance(v, datetime.date):
        return v.strftime("%d/%m/%Y")
    if isinstance(v, float):
        return str(int(v)) if v.is_integer() else f"{v:,.2f}"
    return str(v).strip()


def veces(n):
    return f"{n} vez" if n == 1 else f"{n} veces"


def describir_diferencia(a, b, max_ops=4):
    """Explica en qué difieren dos llaves (posición y carácter)."""
    if len(a) == len(b):
        dif = [i for i in range(len(a)) if a[i] != b[i]]
        if len(dif) == 2 and dif[1] == dif[0] + 1 and a[dif[0]] == b[dif[1]] and a[dif[1]] == b[dif[0]]:
            i = dif[0]
            return f"intercambio en pos. {i + 1}-{i + 2}: '{a[i:i + 2]}'→'{b[i:i + 2]}'"
    ops = []
    for tag, i1, i2, j1, j2 in difflib.SequenceMatcher(None, a, b).get_opcodes():
        if tag == "replace":
            ops.append(f"pos. {i1 + 1}: '{a[i1:i2]}'→'{b[j1:j2]}'")
        elif tag == "delete":
            ops.append(f"sobra '{a[i1:i2]}' en pos. {i1 + 1}")
        elif tag == "insert":
            ops.append(f"falta '{b[j1:j2]}' en pos. {i1 + 1}")
    extra = f" (+{len(ops) - max_ops} más)" if len(ops) > max_ops else ""
    return "; ".join(ops[:max_ops]) + extra


# ---------------------------------------------------------------------------
# Lectura de archivos
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def hojas_de(contenido, nombre):
    if nombre.lower().endswith(".csv"):
        return ["(CSV)"]
    return pd.ExcelFile(io.BytesIO(contenido)).sheet_names


@st.cache_data(show_spinner=False)
def leer_crudo(contenido, nombre, hoja):
    if nombre.lower().endswith(".csv"):
        try:
            df = pd.read_csv(io.BytesIO(contenido), header=None, dtype=object, sep=None, engine="python")
        except UnicodeDecodeError:
            df = pd.read_csv(io.BytesIO(contenido), header=None, dtype=object, sep=None, engine="python",
                             encoding="latin-1")
    else:
        df = pd.read_excel(io.BytesIO(contenido), sheet_name=hoja, header=None)
    return df.dropna(axis=1, how="all").dropna(how="all")


def detectar_encabezado(df_raw, max_filas=30):
    """Fila con más celdas de texto entre las primeras filas."""
    mejor, mejor_score = 0, -1
    for pos in range(min(max_filas, len(df_raw))):
        textos = sum(1 for v in df_raw.iloc[pos] if isinstance(v, str) and v.strip())
        if textos > mejor_score:
            mejor, mejor_score = pos, textos
    return mejor


def armar_tabla(df_raw, pos_enc):
    """Tabla con encabezados únicos. El índice es la fila real en Excel."""
    enc = df_raw.iloc[pos_enc].tolist()
    df = df_raw.iloc[pos_enc + 1:].copy()
    nombres, vistos = [], {}
    for i, h in enumerate(enc):
        if isinstance(h, float) and not pd.isna(h) and h.is_integer():
            h = int(h)
        n = str(h).strip() if not es_vacio(h) and str(h).strip() else f"Columna_{i + 1}"
        while n in vistos:
            vistos[n] += 1
            n = f"{n}_{vistos[n]}"
        vistos.setdefault(n, 0)
        nombres.append(n)
    df.columns = nombres
    df = df.dropna(how="all").dropna(axis=1, how="all")
    df.index = df.index + 1
    df.index.name = "Fila"
    return df


def para_mostrar(df):
    return df.astype(str).replace({"nan": "", "None": "", "NaT": "", "<NA>": ""})


# ---------------------------------------------------------------------------
# Sugerencias automáticas
# ---------------------------------------------------------------------------

@st.cache_data(show_spinner=False)
def sugerir_llave(df_a, df_b, modo="Estándar", sin_ceros=False, ultimos=0):
    """Par de columnas casi únicas que más se cruzan entre A y B."""
    claves = ("VIN", "SERIE", "FOLIO", "ID", "CLAVE", "RFC", "CURP", "CUENTA", "CREDITO", "CONTRATO",
              "NUMERO", "NO.", "SOLICITUD", "TELEFONO", "CELULAR", "PLACA")
    nk = lambda v: norm_llave(v, modo, sin_ceros, ultimos)  # noqa: E731
    sets_b = {}
    for cb in df_b.columns:
        if inferir_tipo(df_b[cb]) != "Fecha":
            sets_b[cb] = set(df_b[cb].dropna().map(nk)) - {""}
    mejor, mejor_score = (None, None), 0.0
    for ca in df_a.columns:
        if inferir_tipo(df_a[ca]) == "Fecha":
            continue
        va = df_a[ca].dropna().map(nk)
        va = va[va != ""]
        if va.empty:
            continue
        unicidad = va.nunique() / len(va)
        if unicidad < 0.8:
            continue
        set_a = set(va)
        largo = min(1.0, va.str.len().mean() / 6)
        for cb, sb in sets_b.items():
            inter = len(set_a & sb)
            if not inter:
                continue
            score = inter / min(len(set_a), len(sb)) * unicidad * largo
            nombres = quitar_acentos(f"{ca} {cb}").upper()
            if any(p in nombres for p in claves):
                score += 0.15
            if score > mejor_score:
                mejor, mejor_score = (ca, cb), score
    return mejor


def _tasa_acuerdo(sa, sb):
    pares = [(x, y) for x, y in zip(sa, sb) if not es_vacio(x) and not es_vacio(y)]
    if len(pares) < 3:
        return 0.0, 0.0, "Texto flexible", len(pares)
    ta = inferir_tipo(pd.Series([p[0] for p in pares], dtype=object))
    tb = inferir_tipo(pd.Series([p[1] for p in pares], dtype=object))
    tipo = "Fecha" if "Fecha" in (ta, tb) else "Número" if ta == tb == "Número" else "Texto flexible"
    igual = contiene = 0
    for x, y in pares:
        a, b = norm_valor(x, tipo), norm_valor(y, tipo)
        if a is None or b is None:
            continue
        if tipo == "Número":
            ok = abs(a - b) <= 0.01
            igual += ok
            contiene += ok
            continue
        igual += a == b
        if isinstance(a, str) and min(len(a), len(b)) >= 2 and (a in b or b in a):
            contiene += 1
    return igual / len(pares), contiene / len(pares), tipo, len(pares)


@st.cache_data(show_spinner=False)
def sugerir_comparaciones(df_a, df_b, ka, kb, excluir_a, excluir_b, muestra=300):
    a = df_a.assign(_k=ka)
    b = df_b.assign(_k=kb)
    a = a[a["_k"] != ""].drop_duplicates("_k").set_index("_k")
    b = b[b["_k"] != ""].drop_duplicates("_k").set_index("_k")
    comunes = list(a.index.intersection(b.index))[:muestra]
    a_c, b_c = a.loc[comunes], b.loc[comunes]

    sugerencias, usadas_b = [], set(excluir_b)
    for ca in df_a.columns:
        if ca in excluir_a:
            continue
        mejor = None
        for cb in df_b.columns:
            if cb in usadas_b:
                continue
            sim = difflib.SequenceMatcher(None, norm_nombre(ca), norm_nombre(cb)).ratio()
            igual, cont, tipo, n_pares = _tasa_acuerdo(a_c[ca], b_c[cb])
            if igual >= 0.5:
                score, modo = igual + 1, "Igual"
            elif cont >= 0.7 and tipo == "Texto flexible":
                score, modo = cont + 0.5, "Contiene"
            elif sim >= 0.8 and n_pares < 3 and inferir_tipo(df_a[ca]) == inferir_tipo(df_b[cb]):
                score, modo, tipo = sim, "Igual", inferir_tipo(df_a[ca])
            else:
                continue
            if mejor is None or score > mejor[0]:
                mejor = (score, cb, tipo, modo)
        if mejor:
            usadas_b.add(mejor[1])
            tol = 0.01 if mejor[2] == "Número" else 0.0
            sugerencias.append({"Columna A": ca, "Columna B": mejor[1], "Tipo": mejor[2],
                                "Modo": mejor[3], "Tolerancia": tol})
    return sugerencias


@st.cache_data(show_spinner=False)
def diagnostico_llave(ka, kb, nombre_a, nombre_b):
    def lado(k, otro):
        validas = k[k != ""]
        cnt = validas.value_counts()
        rep = cnt[cnt > 1]
        cruzan = validas.isin(set(otro[otro != ""])).sum()
        return [len(k), int((k == "").sum()), int(validas.nunique()),
                f"{len(rep)} llaves ({int(rep.sum())} filas)", int(cruzan),
                f"{cruzan / len(validas):.1%}" if len(validas) else "0%"]
    filas = ["Registros", "Llaves vacías", "Llaves únicas", "Llaves repetidas",
             "Filas que cruzan con el otro archivo", "Cobertura"]
    return pd.DataFrame({nombre_a: lado(ka, kb), nombre_b: lado(kb, ka)}, index=filas).astype(str)


# ---------------------------------------------------------------------------
# Motor del cuadre
# ---------------------------------------------------------------------------

def preparar_comparaciones(lista):
    out, usadas, repetidas = [], set(), {}
    for c in lista:
        tipo = c.get("Tipo") if c.get("Tipo") in TIPOS else "Texto flexible"
        modo = c.get("Modo") if c.get("Modo") in MODOS else "Igual"
        tol = c.get("Tolerancia")
        tol = 0.0 if es_vacio(tol) else abs(float(tol))
        if tipo in ("Número", "Fecha"):
            modo = "Igual"
        elif modo == "Similar":
            tol = 85.0 if tol <= 0 else min(tol, 100.0)
        else:
            tol = 0.0
        etiqueta = str(c["Columna A"])
        if etiqueta.lower() in usadas:
            repetidas[etiqueta.lower()] = repetidas.get(etiqueta.lower(), 1) + 1
            etiqueta = f"{etiqueta} ({repetidas[etiqueta.lower()]})"
        usadas.add(etiqueta.lower())
        out.append({"Columna A": c["Columna A"], "Columna B": c["Columna B"], "Tipo": tipo, "Modo": modo,
                    "Tolerancia": tol, "etiqueta": etiqueta})
    return out


def evaluar(va, vb, comp, presente):
    tipo, modo, tol = comp["Tipo"], comp["Modo"], comp["Tolerancia"]
    na, nb = norm_valor(va, tipo), norm_valor(vb, tipo)
    e = {"res": R_NA, "na": na, "nb": nb, "dif": None, "sim": None, "va": va, "vb": vb}
    if not presente:
        return e
    if na is None and nb is None:
        e["res"] = R_AMBOS
        return e
    if na is None or nb is None:
        e["res"] = R_VACIO
        return e
    if tipo == "Número":
        e["dif"] = na - nb
        ok = abs(e["dif"]) <= tol + 1e-9
    elif tipo == "Fecha":
        e["dif"] = (na - nb).days
        ok = abs(e["dif"]) <= tol
    elif modo == "Similar":
        e["sim"] = round(difflib.SequenceMatcher(None, na, nb).ratio() * 100, 1)
        ok = e["sim"] >= tol
    elif modo == "Contiene":
        ok = na in nb or nb in na
    elif modo == "Empieza con":
        ok = na.startswith(nb) or nb.startswith(na)
    else:
        ok = na == nb
    e["res"] = R_OK if ok else R_NO
    return e


def distancia_osa(a, b, maximo):
    """Distancia de edición (cambio, alta, baja o intercambio de vecinos = 1). Corta si supera 'maximo'."""
    if abs(len(a) - len(b)) > maximo:
        return maximo + 1
    prev2, prev = None, list(range(len(b) + 1))
    for i in range(1, len(a) + 1):
        cur = [i] + [0] * len(b)
        for j in range(1, len(b) + 1):
            costo = 0 if a[i - 1] == b[j - 1] else 1
            cur[j] = min(prev[j] + 1, cur[j - 1] + 1, prev[j - 1] + costo)
            if i > 1 and j > 1 and a[i - 1] == b[j - 2] and a[i - 2] == b[j - 1]:
                cur[j] = min(cur[j], prev2[j - 2] + 1)
        if min(cur) > maximo:
            return maximo + 1
        prev2, prev = prev, cur
    return prev[-1]


def _borrados(k, n):
    """Variantes de k quitando hasta n caracteres (índice de 'borrado simétrico')."""
    vistas, frontera = {k}, {k}
    for _ in range(n):
        frontera = {x[:i] + x[i + 1:] for x in frontera for i in range(len(x))} - vistas
        vistas |= frontera
    return vistas


def buscar_parecidas(origen, destino, max_cambios=1, largo_min=4, limite=60_000):
    """Llaves de 'origen' a ≤ max_cambios caracteres de alguna llave de 'destino'."""
    origen = sorted({k for k in origen if len(k) >= largo_min})
    destino = sorted({k for k in destino if len(k) >= largo_min})
    if not origen or not destino:
        return {}, False
    if len(origen) + len(destino) > limite:
        return {}, True
    indice = {}
    for d in destino:
        for v in _borrados(d, max_cambios):
            indice.setdefault(v, set()).add(d)
    res = {}
    for k in origen:
        candidatos = set()
        for v in _borrados(k, max_cambios):
            candidatos |= indice.get(v, set())
        mejor = None
        for c in candidatos:
            dist = distancia_osa(k, c, max_cambios)
            if dist <= max_cambios and (mejor is None or dist < mejor[1]):
                mejor = (c, dist)
        if mejor:
            sim = round(difflib.SequenceMatcher(None, k, mejor[0]).ratio() * 100, 1)
            res[k] = (mejor[0], sim, describir_diferencia(k, mejor[0]), mejor[1])
    return res, False


def nombres_unicos(nombres):
    vistos, out = {}, []
    for n in nombres:
        base, i = n, 2
        while n in vistos:
            n = f"{base} [{i}]"
            i += 1
        vistos[n] = 1
        out.append(n)
    return out


def ejecutar_cuadre(df_a, df_b, cfg):
    A, B = cfg["nombre_a"], cfg["nombre_b"]
    L = etiquetas(A, B)
    opts = (cfg["norm_llave"], cfg["sin_ceros"], cfg["ultimos"])
    comps = preparar_comparaciones(cfg["comparaciones"])

    ka = construir_llave(df_a, cfg["llave_a"], *opts)
    kb = construir_llave(df_b, cfg["llave_b"], *opts)
    va_, vb_ = ka[ka != ""], kb[kb != ""]
    cnt_a, cnt_b = va_.value_counts().to_dict(), vb_.value_counts().to_dict()

    ta = pd.DataFrame({"_k": va_.values, "_n": va_.groupby(va_).cumcount().values, "fa": va_.index})
    tb = pd.DataFrame({"_k": vb_.values, "_n": vb_.groupby(vb_).cumcount().values, "fb": vb_.index})
    m = ta.merge(tb, on=["_k", "_n"], how="outer" if cfg["incluir_solo_b"] else "left")

    pares = [(k, int(n), int(fa) if pd.notna(fa) else None, int(fb) if pd.notna(fb) else None)
             for k, n, fa, fb in m[["_k", "_n", "fa", "fb"]].itertuples(index=False)]
    pares += [("", 0, int(i), None) for i in ka.index[ka == ""]]
    if cfg["incluir_solo_b"]:
        pares += [("", 0, None, int(i)) for i in kb.index[kb == ""]]
    pares.sort(key=lambda p: (0, p[2]) if p[2] is not None else (1, p[3]))

    # Posibles errores de captura (llaves que solo existen de un lado)
    parecidas_a, parecidas_b, omitida = {}, {}, False
    if cfg["fuzzy"]:
        solo_a_k = [k for k in cnt_a if k not in cnt_b]
        solo_b_k = [k for k in cnt_b if k not in cnt_a]
        parecidas_a, o1 = buscar_parecidas(solo_a_k, solo_b_k, cfg["max_cambios"])
        parecidas_b, o2 = (buscar_parecidas(solo_b_k, solo_a_k, cfg["max_cambios"]) if cfg["incluir_solo_b"]
                           else ({}, False))
        omitida = o1 or o2

    def llave_original(df, cols, idx):
        return " | ".join(fmt(df.at[idx, c]) if not es_vacio(df.at[idx, c]) else "" for c in cols)

    filas = []
    for k, n, fa, fb in pares:
        presente = fa is not None and fb is not None
        ca_n, cb_n = cnt_a.get(k, 0) if k else 0, cnt_b.get(k, 0) if k else 0
        obs, evals = [], []
        for c in comps:
            va = df_a.at[fa, c["Columna A"]] if fa is not None else None
            vb = df_b.at[fb, c["Columna B"]] if fb is not None else None
            e = evaluar(va, vb, c, presente)
            evals.append(e)
            if e["res"] == R_NO:
                extra = ""
                if c["Tipo"] == "Número":
                    extra = f" · dif. {e['dif']:+,.2f}"
                elif c["Tipo"] == "Fecha":
                    extra = f" · dif. {e['dif']:+d} días"
                elif e["sim"] is not None:
                    extra = f" · similitud {e['sim']}%"
                obs.append(("Diferencia", c["etiqueta"], va, vb, e["dif"],
                            f"{c['etiqueta']} no coincide ({A}: {fmt(va)} / {B}: {fmt(vb)}){extra}"))
            elif e["res"] == R_VACIO:
                lado = A if e["na"] is None else B
                col = c["Columna A"] if lado == A else c["Columna B"]
                obs.append(("Dato vacío", c["etiqueta"], va, vb, None, f"{col} vacío en {lado}"))
        n_dif = sum(e["res"] == R_NO for e in evals)
        n_vac = sum(e["res"] == R_VACIO for e in evals)

        posible = None
        if not k:
            est = "sin"
            lado, cols = (A, cfg["llave_a"]) if fa is not None else (B, cfg["llave_b"])
            obs.insert(0, ("Sin llave", ", ".join(cols), None, None, None,
                           f"Sin valor en la llave ({', '.join(cols)}) en {lado}"))
        elif fb is None:
            est = "solo_a"
            if cb_n:
                obs.insert(0, ("Duplicado sin par", "Llave", k, None, None,
                               f"La llave está {veces(ca_n)} en {A} y {veces(cb_n)} en {B}; esta ocurrencia (#{n + 1}) quedó sin par"))
            else:
                obs.insert(0, ("No encontrado", "Llave", k, None, None, f"No se encontró en {B}"))
                posible = parecidas_a.get(k)
        elif fa is None:
            est = "solo_b"
            if ca_n:
                obs.insert(0, ("Duplicado sin par", "Llave", None, k, None,
                               f"La llave está {veces(cb_n)} en {B} y {veces(ca_n)} en {A}; esta ocurrencia (#{n + 1}) quedó sin par"))
            else:
                obs.insert(0, ("No encontrado", "Llave", None, k, None, f"No se encontró en {A}"))
                posible = parecidas_b.get(k)
        elif n_dif:
            est = "dif"
        elif ca_n > 1 or cb_n > 1:
            est = "dup"
        elif n_vac:
            est = "inc"
        else:
            est = "ok"

        if posible:
            otro = B if est == "solo_a" else A
            obs.append(("Posible error de captura", "Llave", k if est == "solo_a" else posible[0],
                        posible[0] if est == "solo_a" else k, None,
                        f"Parecida a {posible[0]} en {otro} ({'1 carácter distinto' if posible[3] == 1 else f'{posible[3]} caracteres distintos'}): {posible[2]}"))
        if k and (ca_n > 1 or cb_n > 1) and not (obs and obs[0][0] == "Duplicado sin par"):
            obs.append(("Duplicado", "Llave", None, None, None,
                        f"La llave aparece {veces(ca_n)} en {A} y {veces(cb_n)} en {B}"
                        + (f"; se emparejó la ocurrencia #{n + 1}" if presente else "")))

        if not k:
            tipo_cruce = "Sin llave"
        elif not presente:
            tipo_cruce = "Sin cruce"
        elif ca_n > 1 or cb_n > 1:
            tipo_cruce = f"Exacto · ocurrencia {n + 1} ({A}: {ca_n} / {B}: {cb_n})"
        else:
            tipo_cruce = "Exacto"

        filas.append({
            "k": k, "fa": fa, "fb": fb, "veces_a": ca_n, "veces_b": cb_n, "est": est, "evals": evals,
            "obs": obs, "posible": posible, "tipo_cruce": tipo_cruce, "n_dif": n_dif, "n_vac": n_vac,
            "llave_orig": llave_original(df_a, cfg["llave_a"], fa) if fa is not None
            else llave_original(df_b, cfg["llave_b"], fb),
            "extras_a": [df_a.at[fa, c] if fa is not None else None for c in cfg["extras_a"]],
            "extras_b": [df_b.at[fb, c] if fb is not None else None for c in cfg["extras_b"]],
        })

    # ---- DataFrame para la vista
    enc = ["Estatus", "Llave", "Llave normalizada", f"Fila en {A}", f"Fila en {B}", f"Veces en {A}",
           f"Veces en {B}", "Tipo de cruce"]
    for c in comps:
        enc += [f"{c['Columna A']} ({A})", f"{c['Columna B']} ({B})", f"Resultado {c['etiqueta']}"]
        if c["Tipo"] in ("Número", "Fecha"):
            enc.append(f"Diferencia {c['etiqueta']}")
        if c["Modo"] == "Similar" and c["Tipo"] not in ("Número", "Fecha"):
            enc.append(f"Similitud % {c['etiqueta']}")
    enc += ["Observaciones", "Posible coincidencia"]
    enc += [f"{e} ({A})" for e in cfg["extras_a"]] + [f"{e} ({B})" for e in cfg["extras_b"]]
    enc = nombres_unicos(enc)

    registros = []
    for f in filas:
        fila = [L[f["est"]], f["llave_orig"], f["k"], f["fa"], f["fb"],
                f["veces_a"] if f["k"] else None, f["veces_b"] if f["k"] else None, f["tipo_cruce"]]
        for c, e in zip(comps, f["evals"]):
            fila += [e["va"], e["vb"], e["res"]]
            if c["Tipo"] in ("Número", "Fecha"):
                fila.append(e["dif"])
            if c["Modo"] == "Similar" and c["Tipo"] not in ("Número", "Fecha"):
                fila.append(e["sim"])
        fila += [" | ".join(o[5] for o in f["obs"]) or "OK",
                 f"{f['posible'][0]} ({f['posible'][3]} car. distinto)" if f["posible"] else ""]
        fila += f["extras_a"] + f["extras_b"]
        registros.append(fila)
    vista = pd.DataFrame(registros, columns=enc)

    # ---- Observaciones en formato largo
    obs_rows = []
    for f in filas:
        for tipo, col, va, vb, dif, det in f["obs"]:
            obs_rows.append({"Llave": f["k"] or f["llave_orig"], "Estatus": L[f["est"]], "Tipo de observación": tipo,
                             "Columna": col, f"Valor en {A}": va, f"Valor en {B}": vb, "Diferencia": dif,
                             "Detalle": det, f"Fila en {A}": f["fa"], f"Fila en {B}": f["fb"]})
    obs_df = pd.DataFrame(obs_rows, columns=["Llave", "Estatus", "Tipo de observación", "Columna",
                                             f"Valor en {A}", f"Valor en {B}", "Diferencia", "Detalle",
                                             f"Fila en {A}", f"Fila en {B}"])

    # ---- Posibles coincidencias
    filas_a = va_.groupby(va_).apply(lambda s: list(s.index)).to_dict() if len(va_) else {}
    filas_b = vb_.groupby(vb_).apply(lambda s: list(s.index)).to_dict() if len(vb_) else {}
    pos_rows = []
    for k, (match, sim, det, dist) in parecidas_a.items():
        pos_rows.append({"Lado": L["solo_a"], "Llave sin cruce": k, "Llave parecida en el otro archivo": match,
                         "Caracteres distintos": dist, "Similitud %": sim, "Qué cambia": det,
                         "Fila(s) origen": ", ".join(map(str, filas_a.get(k, []))),
                         "Fila(s) de la llave parecida": ", ".join(map(str, filas_b.get(match, [])))})
    for k, (match, sim, det, dist) in parecidas_b.items():
        pos_rows.append({"Lado": L["solo_b"], "Llave sin cruce": k, "Llave parecida en el otro archivo": match,
                         "Caracteres distintos": dist, "Similitud %": sim, "Qué cambia": det,
                         "Fila(s) origen": ", ".join(map(str, filas_b.get(k, []))),
                         "Fila(s) de la llave parecida": ", ".join(map(str, filas_a.get(match, [])))})
    pos_df = pd.DataFrame(pos_rows, columns=["Lado", "Llave sin cruce", "Llave parecida en el otro archivo",
                                             "Caracteres distintos", "Similitud %", "Qué cambia", "Fila(s) origen",
                                             "Fila(s) de la llave parecida"])
    if not pos_df.empty:
        pos_df = pos_df.sort_values(["Caracteres distintos", "Similitud %"], ascending=[True, False]).reset_index(drop=True)

    # ---- Duplicados
    dup_keys = sorted({k for k, v in cnt_a.items() if v > 1} | {k for k, v in cnt_b.items() if v > 1})
    dup_df = pd.DataFrame([{"Llave normalizada": k, f"Veces en {A}": cnt_a.get(k, 0), f"Veces en {B}": cnt_b.get(k, 0),
                            f"Filas en {A}": ", ".join(map(str, filas_a.get(k, []))),
                            f"Filas en {B}": ", ".join(map(str, filas_b.get(k, [])))} for k in dup_keys],
                          columns=["Llave normalizada", f"Veces en {A}", f"Veces en {B}", f"Filas en {A}", f"Filas en {B}"])

    # ---- Estadística por columna
    stats = []
    for i, c in enumerate(comps):
        res = [f["evals"][i]["res"] for f in filas]
        difs = [f["evals"][i]["dif"] for f in filas if f["evals"][i]["dif"] is not None]
        ok, no = res.count(R_OK), res.count(R_NO)
        stats.append({"Columna": c["etiqueta"], f"Columna en {B}": c["Columna B"], "Tipo": c["Tipo"],
                      "Modo": c["Modo"], "✔ Coinciden": ok, "✘ Diferentes": no, "Vacío": res.count(R_VACIO),
                      "% coincidencia": round(ok / (ok + no) * 100, 1) if ok + no else None,
                      "Suma diferencias": round(sum(difs), 2) if difs else None,
                      "Mayor diferencia abs.": round(max(abs(d) for d in difs), 2) if difs else None})
    stats_df = pd.DataFrame(stats)

    conteo = {k: sum(f["est"] == k for f in filas) for k in ORDEN}
    cruzadas = sum(1 for f in filas if f["fa"] is not None and f["fb"] is not None)
    return {"cfg": cfg, "labels": L, "comps": comps, "filas": filas, "vista": vista, "obs": obs_df,
            "posibles": pos_df, "duplicados": dup_df, "stats": stats_df, "conteo": conteo, "cruzadas": cruzadas,
            "ka": ka, "kb": kb, "fuzzy_omitida": omitida}


# ---------------------------------------------------------------------------
# Excel de salida
# ---------------------------------------------------------------------------

FUENTE = "Arial"
F_NORMAL = Font(name=FUENTE, size=10)
F_BOLD = Font(name=FUENTE, size=10, bold=True)
F_BLANCA = Font(name=FUENTE, size=10, bold=True, color="FFFFFF")
F_TITULO = Font(name=FUENTE, size=16, bold=True, color="1F4E78")
F_SUB = Font(name=FUENTE, size=10, italic=True, color="595959")
F_SECCION = Font(name=FUENTE, size=12, bold=True, color="1F4E78")
BORDE = Border(*(Side(style="thin", color="BFBFBF"),) * 4)
GRUPOS_COLOR = ["2E75B6", "548235", "BF8F00", "7030A0", "C55A11", "2F5597"]


def relleno(hex_color):
    return PatternFill(start_color=hex_color, end_color=hex_color, fill_type="solid")


def q(s):
    return '"' + str(s).replace('"', '""') + '"'


def ref_hoja(nombre):
    return "'" + nombre.replace("'", "''") + "'"


def nombre_hoja(base, usados):
    s = re.sub(r"[\[\]:*?/\\]", "", base).strip()[:31] or "Hoja"
    s0, i = s, 2
    while s.lower() in usados:
        s = f"{s0[:27]} {i}"
        i += 1
    usados.add(s.lower())
    return s


def xl(v):
    if es_vacio(v):
        return None
    if isinstance(v, pd.Timestamp):
        return v.to_pydatetime()
    if hasattr(v, "item") and not isinstance(v, (str, bytes, datetime.date)):
        v = v.item()
    if isinstance(v, str):
        v = ILLEGAL_CHARACTERS_RE.sub("", v)
    return v


def escribir(ws, fila, col, valor, fuente=F_NORMAL, formato=None, borde=True):
    cel = ws.cell(row=fila, column=col, value=xl(valor))
    cel.font = fuente
    if formato:
        cel.number_format = formato
    elif isinstance(cel.value, (datetime.datetime, datetime.date)):
        cel.number_format = "dd/mm/yyyy"
    if borde:
        cel.border = BORDE
    return cel


def escribir_dato(ws, fila, col, valor, formato=None):
    """Escribe un dato (nunca fórmula): textos que empiezan con '=' se guardan como texto."""
    cel = escribir(ws, fila, col, valor, formato=formato)
    if isinstance(cel.value, str) and cel.value.startswith("="):
        cel.data_type = "s"
    return cel


def formula(ws, fila, col, texto, fuente=F_NORMAL, formato=None):
    cel = ws.cell(row=fila, column=col, value=texto)
    cel.font, cel.border = fuente, BORDE
    if formato:
        cel.number_format = formato
    return cel


def encabezado_tabla(ws, fila, col_ini, titulos, color="1F4E78", alto=30):
    for i, t in enumerate(titulos):
        c = ws.cell(row=fila, column=col_ini + i, value=t)
        c.fill, c.font, c.border = relleno(color), F_BLANCA, BORDE
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    ws.row_dimensions[fila].height = alto


def titulo_hoja(ws, titulo, subtitulo):
    ws["A1"] = titulo
    ws["A1"].font = F_TITULO
    ws["A2"] = subtitulo
    ws["A2"].font = F_SUB
    ws.sheet_view.showGridLines = False


def contar(rango, celda, sumproduct):
    return f"SUMPRODUCT(--({rango}={celda}))" if sumproduct else f"COUNTIF({rango},{celda})"


def generar_excel(R, df_a, df_b, origenes):
    cfg, L, comps, filas = R["cfg"], R["labels"], R["comps"], R["filas"]
    A, B = cfg["nombre_a"], cfg["nombre_b"]
    n = len(filas)
    ini, fin = 5, 4 + n

    # COUNTIF trata llaves numéricas o con comodines de forma inexacta: en ese caso se usa SUMPRODUCT
    def riesgosa(k):
        if not k:
            return False
        try:
            float(k)
            return True
        except ValueError:
            return bool(re.search(r"[*?~<>=]", k)) or len(k) > 250
    todas = set(R["ka"]) | set(R["kb"])
    usar_sp = any(riesgosa(k) for k in todas)
    formulas_veces = not (usar_sp and n * (len(df_a) + len(df_b)) > 40_000_000)

    wb = Workbook()
    wb.calculation = CalcProperties(fullCalcOnLoad=True)
    usados = {"resumen", "cuadre", "observaciones", "posibles coincidencias", "duplicados", "configuración", "leyenda"}
    ws_res = wb.active
    ws_res.title = "Resumen"
    ws_c = wb.create_sheet("Cuadre")
    ws_o = wb.create_sheet("Observaciones")
    ws_p = wb.create_sheet("Posibles coincidencias")
    ws_d = wb.create_sheet("Duplicados")
    h_a, h_b = nombre_hoja(f"Datos {A}", usados), nombre_hoja(f"Datos {B}", usados)
    ws_da, ws_db = wb.create_sheet(h_a), wb.create_sheet(h_b)
    ws_cfg = wb.create_sheet("Configuración")
    ws_ley = wb.create_sheet("Leyenda")
    fin_a, fin_b = 4 + max(len(df_a), 1), 4 + max(len(df_b), 1)
    rng_ka = f"{ref_hoja(h_a)}!$B$5:$B${fin_a}"
    rng_kb = f"{ref_hoja(h_b)}!$B$5:$B${fin_b}"

    # ---------------- Leyenda (rango fijo A5:B11 para VLOOKUP)
    titulo_hoja(ws_ley, "Leyenda", "Qué significa cada estatus, resultado y opción del cuadre.")
    encabezado_tabla(ws_ley, 4, 1, ["Estatus", "Prioridad", "Significado", "Acción sugerida"])
    for i, k in enumerate(ORDEN):
        r = 5 + i
        escribir(ws_ley, r, 1, L[k], F_BOLD).fill = relleno(INFO[k][0])
        escribir(ws_ley, r, 2, INFO[k][1])
        escribir(ws_ley, r, 3, INFO[k][2])
        escribir(ws_ley, r, 4, INFO[k][3])
    r = 14
    ws_ley.cell(row=r, column=1, value="Resultados por columna").font = F_SECCION
    encabezado_tabla(ws_ley, r + 1, 1, ["Símbolo", "Significado"])
    for i, (s, d) in enumerate([(R_OK, "Los valores coinciden según el tipo, modo y tolerancia."),
                                (R_NO, "Los valores son diferentes."),
                                (R_VACIO, "El dato está vacío en uno de los dos archivos."),
                                (R_AMBOS, "El dato está vacío en ambos archivos."),
                                (R_NA, "No aplica: la llave no cruzó, no hay contra qué comparar.")]):
        escribir(ws_ley, r + 2 + i, 1, s, F_BOLD)
        escribir(ws_ley, r + 2 + i, 2, d)
    r = 23
    ws_ley.cell(row=r, column=1, value="Tipos y modos de comparación").font = F_SECCION
    encabezado_tabla(ws_ley, r + 1, 1, ["Opción", "Cómo compara"])
    explicaciones = [
        ("Texto flexible", "Ignora mayúsculas, acentos, espacios, guiones, símbolos y ceros a la izquierda (MG-05 = MG5)."),
        ("Texto exacto", "Compara el texto tal cual (solo quita espacios al inicio y al final); distingue mayúsculas."),
        ("Número", "Convierte a número (quita $ y comas). Coincide si la diferencia ≤ tolerancia."),
        ("Fecha", "Convierte a fecha (día/mes/año). Coincide si la diferencia en días ≤ tolerancia."),
        ("Modo Igual", "Los valores normalizados deben ser idénticos."),
        ("Modo Contiene", "Un valor debe estar dentro del otro (MG5 ≈ MG5 - 2027)."),
        ("Modo Empieza con", "Un valor debe ser el inicio del otro."),
        ("Modo Similar", "Parecido de texto ≥ tolerancia en % (útil para nombres con errores de dedo)."),
        ("Llave compuesta", "Varias columnas unidas con '|'. Si falta una parte, el registro queda 'Sin llave'."),
        ("Duplicados", "Si una llave se repite, se empareja por orden: 1.ª de A con 1.ª de B, 2.ª con 2.ª, etc."),
        ("Posible error de captura", "Llave sin cruce que difiere en pocos caracteres de otra del otro archivo "
                                     "(cambio, falta, sobra o dos vecinos intercambiados); indica la posición."),
    ]
    for i, (o, d) in enumerate(explicaciones):
        escribir(ws_ley, r + 2 + i, 1, o, F_BOLD)
        escribir(ws_ley, r + 2 + i, 2, d)
    for col, w in zip("ABCD", (26, 95, 70, 60)):
        ws_ley.column_dimensions[col].width = w

    # ---------------- Configuración (tolerancias editables)
    titulo_hoja(ws_cfg, "Configuración del cuadre",
                "Las celdas amarillas se pueden editar: los resultados de la hoja Cuadre se recalculan solos.")
    encabezado_tabla(ws_cfg, 4, 1, ["Parámetro", "Valor"])
    params = [
        (f"Origen {A}", origenes[0]), (f"Origen {B}", origenes[1]),
        (f"Llave en {A}", " + ".join(cfg["llave_a"])), (f"Llave en {B}", " + ".join(cfg["llave_b"])),
        ("Normalización de llave", cfg["norm_llave"]),
        ("Ignorar ceros a la izquierda", "Sí" if cfg["sin_ceros"] else "No"),
        ("Usar solo los últimos N caracteres", cfg["ultimos"] or "Todos"),
        (f"Incluir registros que solo están en {B}", "Sí" if cfg["incluir_solo_b"] else "No"),
        ("Buscar posibles errores de captura", f"Sí (≤ {cfg['max_cambios']} carácter(es) distinto(s))"
         if cfg["fuzzy"] else "No"),
        ("Fecha del cuadre", datetime.datetime.now().strftime("%d/%m/%Y %H:%M")),
    ]
    for i, (p, v) in enumerate(params):
        escribir(ws_cfg, 5 + i, 1, p, F_BOLD)
        escribir_dato(ws_cfg, 5 + i, 2, v)
    r0 = 5 + len(params) + 2
    ws_cfg.cell(row=r0 - 1, column=1, value="Columnas comparadas").font = F_SECCION
    encabezado_tabla(ws_cfg, r0, 1, ["#", f"Columna en {A}", f"Columna en {B}", "Tipo", "Modo", "Tolerancia",
                                     "Qué significa la tolerancia"])
    for i, c in enumerate(comps):
        r = r0 + 1 + i
        escribir(ws_cfg, r, 1, i + 1)
        escribir_dato(ws_cfg, r, 2, c["Columna A"])
        escribir_dato(ws_cfg, r, 3, c["Columna B"])
        escribir(ws_cfg, r, 4, c["Tipo"])
        escribir(ws_cfg, r, 5, c["Modo"])
        t = escribir(ws_cfg, r, 6, c["Tolerancia"], F_BOLD, "0.00")
        t.fill = relleno("FFFF00")
        t.comment = Comment("Puedes cambiar este valor; Resultado, Estatus y Resumen se recalculan.", "Cuadre")
        sig = {"Número": "Diferencia absoluta máxima permitida", "Fecha": "Días de diferencia permitidos"}.get(
            c["Tipo"], "Similitud mínima en %" if c["Modo"] == "Similar" else "No aplica para este modo")
        escribir(ws_cfg, r, 7, sig)
        c["tol_ref"] = f"'Configuración'!$F${r}"
    if not comps:
        escribir(ws_cfg, r0 + 1, 1, "Sin columnas comparadas: solo se cruzaron llaves.")
    for col, w in zip("ABCDEFG", (38, 45, 28, 16, 14, 12, 40)):
        ws_cfg.column_dimensions[col].width = w

    # ---------------- Datos de origen
    def hoja_datos(ws, df, k_series, nombre, otro, rng_otro, hoja_nombre):
        titulo_hoja(ws, f"Datos de {nombre}",
                    f"{origenes[0] if nombre == A else origenes[1]} · datos usados en el cuadre (después de filtros). "
                    "Columnas B y C: llave normalizada y si existe en el otro archivo (fórmula).")
        cols = list(df.columns)
        encabezado_tabla(ws, 4, 1, ["Fila origen", "Llave normalizada", f"¿Está en {otro}?"] + cols, "595959")
        for j in range(1, 4):
            ws.cell(row=4, column=j).fill = relleno("1F4E78")
        for i, idx in enumerate(df.index):
            r = 5 + i
            escribir(ws, r, 1, int(idx))
            escribir_dato(ws, r, 2, k_series[idx] or None)
            if formulas_veces:
                formula(ws, r, 3, f'=IF($B{r}="","Sin llave",IF({contar(rng_otro, f"$B{r}", usar_sp)}>0,"Sí","No"))')
            else:
                escribir(ws, r, 3, "Sí" if k_series[idx] and k_series[idx] in set_otro[nombre] else
                         ("Sin llave" if not k_series[idx] else "No"))
            for j, col in enumerate(cols):
                escribir_dato(ws, r, 4 + j, df.at[idx, col])
        ult = 4 + max(len(df), 1)
        rng = f"C5:C{ult}"
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"No"'], fill=relleno("FFC7CE"),
                                                      font=Font(name=FUENTE, color="9C0006", bold=True)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Sí"'],
                                                      font=Font(name=FUENTE, color="006100", bold=True)))
        ws.conditional_formatting.add(rng, CellIsRule(operator="equal", formula=['"Sin llave"'], fill=relleno("D9D9D9")))
        ws.column_dimensions["A"].width = 10
        ws.column_dimensions["B"].width = 24
        ws.column_dimensions["C"].width = 14
        for j in range(len(cols)):
            ws.column_dimensions[get_column_letter(4 + j)].width = 16
        ws.freeze_panes = "D5"
        ws.auto_filter.ref = f"A4:{get_column_letter(3 + len(cols))}{ult}"

    set_otro = {A: set(R["kb"]) - {""}, B: set(R["ka"]) - {""}}
    hoja_datos(ws_da, df_a, R["ka"], A, B, rng_kb, h_a)
    hoja_datos(ws_db, df_b, R["kb"], B, A, rng_ka, h_b)

    # ---------------- Cuadre (hoja principal)
    cols = []

    def add(header, grupo, ancho=14, oculto=False):
        cols.append({"h": header, "g": grupo, "w": ancho, "oculto": oculto})
        return len(cols)

    G_ID = "IDENTIFICACIÓN Y LLAVE"
    c_num, c_est = add("#", G_ID, 6), add("Estatus", G_ID, 18)
    c_kn, c_ko = add("Llave normalizada", G_ID, 22), add("Llave (original)", G_ID, 24)
    c_fa, c_fb = add(f"Fila en {A}", G_ID, 10), add(f"Fila en {B}", G_ID, 10)
    c_va, c_vb = add(f"Veces en {A}", G_ID, 10), add(f"Veces en {B}", G_ID, 10)
    c_tc, c_pr = add("Tipo de cruce", G_ID, 24), add("Prioridad", G_ID, 11)
    for c in comps:
        g = c["etiqueta"].upper()
        c["c_a"] = add(f"{c['Columna A']} ({A})", g, 18)
        c["c_b"] = add(f"{c['Columna B']} ({B})", g, 18)
        if c["Modo"] == "Similar" and c["Tipo"] not in ("Número", "Fecha"):
            c["c_sim"] = add("Similitud %", g, 11)
        if c["Tipo"] in ("Número", "Fecha"):
            c["c_dif"] = add("Diferencia (días)" if c["Tipo"] == "Fecha" else "Diferencia", g, 13)
        c["c_res"] = add("Resultado", g, 11)
    c_nd, c_nv = add("N° diferencias", "RESULTADO", 12), add("N° vacíos", "RESULTADO", 10)
    c_obs = add("Observaciones (al momento del cuadre)", "RESULTADO", 80)
    c_pos, c_psim = add("Posible coincidencia (llave)", "RESULTADO", 24), add("Caracteres distintos (llave)", "RESULTADO", 12)
    c_rev, c_com = add("Revisión", "SEGUIMIENTO", 14), add("Comentarios", "SEGUIMIENTO", 32)
    c_ext_a = [add(f"{e} ({A})", "DATOS ADICIONALES", 18) for e in cfg["extras_a"]]
    c_ext_b = [add(f"{e} ({B})", "DATOS ADICIONALES", 18) for e in cfg["extras_b"]]
    ult_visible = len(cols)
    for c in comps:
        if c["Tipo"] not in ("Número", "Fecha"):
            c["c_na"] = add(f"aux {c['etiqueta']} {A}", "AUXILIARES", 14, True)
            c["c_nb"] = add(f"aux {c['etiqueta']} {B}", "AUXILIARES", 14, True)
    col = {k: get_column_letter(v) for k, v in
           dict(num=c_num, est=c_est, kn=c_kn, fa=c_fa, fb=c_fb, va=c_va, vb=c_vb, nd=c_nd, nv=c_nv,
                rev=c_rev, pr=c_pr).items()}
    LV = get_column_letter(ult_visible)

    titulo_hoja(ws_c, f"Cuadre: {A} vs {B}",
                f"Llave: {' + '.join(cfg['llave_a'])} ↔ {' + '.join(cfg['llave_b'])} · {n:,} registros · "
                "Estatus, Veces, Prioridad, Diferencia, Resultado y N° son fórmulas: si corriges un valor se recalculan.")

    # Fila 3: grupos; fila 4: encabezados
    j = 1
    color_idx = 0
    while j <= len(cols):
        g = cols[j - 1]["g"]
        k = j
        while k + 1 <= len(cols) and cols[k]["g"] == g:
            k += 1
        if g == G_ID:
            color = "1F4E78"
        elif g == "RESULTADO":
            color = "C00000"
        elif g == "SEGUIMIENTO":
            color = "7030A0"
        elif g == "DATOS ADICIONALES":
            color = "595959"
        elif g == "AUXILIARES":
            color = "A6A6A6"
        else:
            color = GRUPOS_COLOR[color_idx % len(GRUPOS_COLOR)]
            color_idx += 1
        ws_c.cell(row=3, column=j, value=g)
        if k > j:
            ws_c.merge_cells(start_row=3, start_column=j, end_row=3, end_column=k)
        for x in range(j, k + 1):
            c3 = ws_c.cell(row=3, column=x)
            c3.fill, c3.font, c3.border = relleno(color), F_BLANCA, BORDE
            c3.alignment = Alignment(horizontal="center", vertical="center")
            c4 = ws_c.cell(row=4, column=x, value=cols[x - 1]["h"])
            c4.fill = relleno("DDEBF7" if color not in ("A6A6A6",) else "EDEDED")
            c4.font = F_BOLD
            c4.border = BORDE
            c4.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        j = k + 1
    ws_c.row_dimensions[4].height = 42

    for i, f in enumerate(filas):
        r = ini + i
        escribir(ws_c, r, c_num, i + 1)
        kn = f"${col['kn']}{r}"
        formula(ws_c, r, c_est,
                f'=IF({kn}="",{q(L["sin"])},IF(${col["fb"]}{r}="",{q(L["solo_a"])},IF(${col["fa"]}{r}="",'
                f'{q(L["solo_b"])},IF(${col["nd"]}{r}>0,{q(L["dif"])},IF(OR(N(${col["va"]}{r})>1,'
                f'N(${col["vb"]}{r})>1),{q(L["dup"])},IF(${col["nv"]}{r}>0,{q(L["inc"])},{q(L["ok"])}))))))',
                F_BOLD)
        escribir_dato(ws_c, r, c_kn, f["k"] or None)
        escribir_dato(ws_c, r, c_ko, f["llave_orig"] or None)
        escribir(ws_c, r, c_fa, f["fa"])
        escribir(ws_c, r, c_fb, f["fb"])
        if formulas_veces:
            formula(ws_c, r, c_va, f'=IF({kn}="","",{contar(rng_ka, kn, usar_sp)})')
            formula(ws_c, r, c_vb, f'=IF({kn}="","",{contar(rng_kb, kn, usar_sp)})')
        else:
            escribir(ws_c, r, c_va, f["veces_a"] if f["k"] else None)
            escribir(ws_c, r, c_vb, f["veces_b"] if f["k"] else None)
        escribir(ws_c, r, c_tc, f["tipo_cruce"])
        formula(ws_c, r, c_pr, f'=IFERROR(VLOOKUP(${col["est"]}{r},Leyenda!$A$5:$B$11,2,FALSE),"")')

        ausente = f'OR(${col["fa"]}{r}="",${col["fb"]}{r}="")'
        for c, e in zip(comps, f["evals"]):
            la, lb = get_column_letter(c["c_a"]), get_column_letter(c["c_b"])
            if c["Tipo"] in ("Número", "Fecha"):
                fmt_v = "dd/mm/yyyy" if c["Tipo"] == "Fecha" else "#,##0.00"
                escribir_dato(ws_c, r, c["c_a"], e["na"] if e["na"] is not None else e["va"], fmt_v)
                escribir_dato(ws_c, r, c["c_b"], e["nb"] if e["nb"] is not None else e["vb"], fmt_v)
                vA, vB = f"{la}{r}", f"{lb}{r}"
                formula(ws_c, r, c["c_dif"], f'=IF(AND(ISNUMBER({vA}),ISNUMBER({vB})),{vA}-{vB},"")',
                        formato="#,##0;-#,##0;0" if c["Tipo"] == "Fecha" else "#,##0.00;-#,##0.00;0.00")
                prueba = f"ABS({vA}-{vB})<={c['tol_ref']}+0.000000001"
                ambos = f"AND(NOT(ISNUMBER({vA})),NOT(ISNUMBER({vB})))"
                uno = f"OR(NOT(ISNUMBER({vA})),NOT(ISNUMBER({vB})))"
            else:
                escribir_dato(ws_c, r, c["c_a"], e["va"])
                escribir_dato(ws_c, r, c["c_b"], e["vb"])
                escribir_dato(ws_c, r, c["c_na"], e["na"])
                escribir_dato(ws_c, r, c["c_nb"], e["nb"])
                nA = f"${get_column_letter(c['c_na'])}{r}"
                nB = f"${get_column_letter(c['c_nb'])}{r}"
                exacto = c["Tipo"] == "Texto exacto"
                if c["Modo"] == "Similar":
                    escribir(ws_c, r, c["c_sim"], e["sim"], formato="0.0")
                    prueba = f"{get_column_letter(c['c_sim'])}{r}>={c['tol_ref']}"
                elif c["Modo"] == "Contiene":
                    fn = "FIND" if exacto else "SEARCH"
                    prueba = f"OR(ISNUMBER({fn}({nA},{nB})),ISNUMBER({fn}({nB},{nA})))"
                elif c["Modo"] == "Empieza con":
                    prueba = (f"OR(EXACT(LEFT({nB},LEN({nA})),{nA}),EXACT(LEFT({nA},LEN({nB})),{nB}))" if exacto
                              else f"OR(LEFT({nB},LEN({nA}))={nA},LEFT({nA},LEN({nB}))={nB})")
                else:
                    prueba = f"EXACT({nA},{nB})" if exacto else f"{nA}={nB}"
                ambos = f'AND({nA}="",{nB}="")'
                uno = f'OR({nA}="",{nB}="")'
            formula(ws_c, r, c["c_res"],
                    f'=IF({ausente},"{R_NA}",IF({ambos},"{R_AMBOS}",IF({uno},"{R_VACIO}",IF({prueba},"{R_OK}","{R_NO}"))))',
                    F_BOLD).alignment = Alignment(horizontal="center")

        if comps:
            celdas_res = [f"{get_column_letter(c['c_res'])}{r}" for c in comps]
            formula(ws_c, r, c_nd, "=" + "+".join(f'({x}="{R_NO}")' for x in celdas_res))
            formula(ws_c, r, c_nv, "=" + "+".join(f'({x}="{R_VACIO}")' for x in celdas_res))
        else:
            escribir(ws_c, r, c_nd, 0)
            escribir(ws_c, r, c_nv, 0)
        escribir_dato(ws_c, r, c_obs, " | ".join(o[5] for o in f["obs"]) or "OK")
        escribir_dato(ws_c, r, c_pos, f["posible"][0] if f["posible"] else None)
        escribir(ws_c, r, c_psim, f["posible"][3] if f["posible"] else None)
        escribir(ws_c, r, c_rev, "No requiere" if f["est"] == "ok" else "Pendiente")
        escribir(ws_c, r, c_com, None)
        for cc, v in zip(c_ext_a, f["extras_a"]):
            escribir_dato(ws_c, r, cc, v)
        for cc, v in zip(c_ext_b, f["extras_b"]):
            escribir_dato(ws_c, r, cc, v)

    for x, cdef in enumerate(cols, 1):
        cd = ws_c.column_dimensions[get_column_letter(x)]
        cd.width = cdef["w"]
        if cdef["oculto"]:
            cd.hidden = True
            cd.outlineLevel = 1
    ws_c.freeze_panes = ws_c.cell(row=ini, column=c_ko + 1)
    ws_c.auto_filter.ref = f"A4:{LV}{fin}"

    dv = DataValidation(type="list", formula1=q(",".join(REVISION)), allow_blank=True,
                        promptTitle="Revisión", prompt="Marca el avance de la revisión de este registro.")
    ws_c.add_data_validation(dv)
    dv.add(f"{col['rev']}{ini}:{col['rev']}{fin}")

    # Formato condicional: primero el de celdas (gana sobre el de fila)
    cf = ws_c.conditional_formatting
    for c in comps:
        rr = f"{get_column_letter(c['c_res'])}{ini}:{get_column_letter(c['c_res'])}{fin}"
        cf.add(rr, CellIsRule(operator="equal", formula=[f'"{R_NO}"'], fill=relleno("FF9B9B"),
                              font=Font(name=FUENTE, bold=True, color="9C0006")))
        cf.add(rr, CellIsRule(operator="equal", formula=[f'"{R_OK}"'], fill=relleno("A9D08E"),
                              font=Font(name=FUENTE, bold=True, color="006100")))
        cf.add(rr, CellIsRule(operator="equal", formula=[f'"{R_VACIO}"'], fill=relleno("FFE699"),
                              font=Font(name=FUENTE, bold=True, color="9C5700")))
        cf.add(rr, CellIsRule(operator="equal", formula=[f'"{R_NA}"'], font=Font(name=FUENTE, color="808080")))
        if "c_dif" in c:
            rd = f"{get_column_letter(c['c_dif'])}{ini}:{get_column_letter(c['c_dif'])}{fin}"
            cf.add(rd, FormulaRule(formula=[f'${get_column_letter(c["c_res"])}{ini}="{R_NO}"'],
                                   font=Font(name=FUENTE, bold=True, color="C00000")))
    for valor, color in (("Corregido", "006100"), ("Justificado", "2F5597"), ("En revisión", "C55A11"),
                         ("Pendiente", "9C0006")):
        cf.add(f"{col['rev']}{ini}:{col['rev']}{fin}",
               CellIsRule(operator="equal", formula=[q(valor)], font=Font(name=FUENTE, bold=True, color=color)))
    cf.add(f"{col['pr']}{ini}:{col['pr']}{fin}",
           CellIsRule(operator="equal", formula=['"Alta"'], font=Font(name=FUENTE, bold=True, color="C00000")))
    for k in ORDEN:
        cf.add(f"A{ini}:{LV}{fin}", FormulaRule(formula=[f"${col['est']}{ini}={q(L[k])}"], fill=relleno(INFO[k][0])))

    # ---------------- Observaciones (formato largo)
    titulo_hoja(ws_o, "Observaciones", "Una fila por observación (fotografía al momento del cuadre; no se recalcula). "
                                       "Usa los filtros para revisar por tipo o por columna.")
    ob = R["obs"]
    encabezado_tabla(ws_o, 4, 1, list(ob.columns))
    for i, row in enumerate(ob.itertuples(index=False)):
        for j2, v in enumerate(row, 1):
            escribir_dato(ws_o, 5 + i, j2, v, "#,##0.00" if j2 == 7 else None)
    fo = 4 + max(len(ob), 1)
    for t, colr in COLORES_OBS.items():
        ws_o.conditional_formatting.add(f"A5:J{fo}", FormulaRule(formula=[f"$C5={q(t)}"], fill=relleno(colr)))
    for col_l, w in zip("ABCDEFGHIJ", (24, 18, 24, 20, 20, 20, 12, 90, 10, 10)):
        ws_o.column_dimensions[col_l].width = w
    ws_o.freeze_panes = "B5"
    ws_o.auto_filter.ref = f"A4:J{fo}"
    if ob.empty:
        escribir(ws_o, 5, 1, "Sin observaciones: todo cuadra.")

    # ---------------- Posibles coincidencias
    titulo_hoja(ws_p, "Posibles errores de captura",
                f"Llaves sin cruce a ≤ {cfg['max_cambios']} carácter(es) de una llave del otro archivo "
                "(un intercambio de vecinos cuenta como 1). Verifícalas antes de corregir."
                if cfg["fuzzy"] else "La búsqueda de posibles coincidencias estaba desactivada.")
    pdf = R["posibles"]
    encabezado_tabla(ws_p, 4, 1, list(pdf.columns), "2F5597")
    for i, row in enumerate(pdf.itertuples(index=False)):
        for j2, v in enumerate(row, 1):
            escribir_dato(ws_p, 5 + i, j2, v, "0.0" if j2 == 5 else None)
    if pdf.empty:
        escribir(ws_p, 5, 1, "No se encontraron llaves parecidas." if cfg["fuzzy"] and not R["fuzzy_omitida"]
                 else "Búsqueda omitida (desactivada o demasiados registros sin cruce).")
    else:
        ws_p.conditional_formatting.add(f"A5:H{4 + len(pdf)}", FormulaRule(formula=["$D5=1"], fill=relleno("DDEBF7")))
    for col_l, w in zip("ABCDEFGH", (20, 24, 30, 12, 12, 50, 16, 22)):
        ws_p.column_dimensions[col_l].width = w
    ws_p.freeze_panes = "A5"

    # ---------------- Duplicados
    titulo_hoja(ws_d, "Llaves duplicadas",
                "Llaves que aparecen más de una vez en algún archivo. Las columnas 'Veces' son fórmulas.")
    ddf = R["duplicados"]
    encabezado_tabla(ws_d, 4, 1, list(ddf.columns), "7030A0")
    for i, row in enumerate(ddf.itertuples(index=False)):
        r = 5 + i
        escribir_dato(ws_d, r, 1, row[0])
        if formulas_veces:
            formula(ws_d, r, 2, f"={contar(rng_ka, f'$A{r}', usar_sp)}")
            formula(ws_d, r, 3, f"={contar(rng_kb, f'$A{r}', usar_sp)}")
        else:
            escribir(ws_d, r, 2, row[1])
            escribir(ws_d, r, 3, row[2])
        escribir(ws_d, r, 4, row[3])
        escribir(ws_d, r, 5, row[4])
    if ddf.empty:
        escribir(ws_d, 5, 1, "No hay llaves duplicadas.")
    else:
        for c_l in "BC":
            ws_d.conditional_formatting.add(f"{c_l}5:{c_l}{4 + len(ddf)}",
                                            CellIsRule(operator="greaterThan", formula=["1"], fill=relleno("E4DFEC"),
                                                       font=Font(name=FUENTE, bold=True, color="7030A0")))
    for col_l, w in zip("ABCDE", (28, 12, 12, 30, 30)):
        ws_d.column_dimensions[col_l].width = w

    # ---------------- Resumen
    titulo_hoja(ws_res, f"Resumen del cuadre: {A} vs {B}",
                f"Generado el {datetime.datetime.now():%d/%m/%Y %H:%M}. Todas las cifras son fórmulas sobre la hoja Cuadre.")
    rng = lambda c_l: f"Cuadre!${c_l}${ini}:${c_l}${fin}"  # noqa: E731
    encabezado_tabla(ws_res, 4, 1, ["Estatus", "Registros", "% del total", "Prioridad"])
    for i, k in enumerate(ORDEN):
        r = 5 + i
        escribir(ws_res, r, 1, L[k], F_BOLD).fill = relleno(INFO[k][0])
        formula(ws_res, r, 2, f"=COUNTIF({rng(col['est'])},{q(L[k])})", formato="#,##0")
        formula(ws_res, r, 3, f"=IF($B$12=0,0,B{r}/$B$12)", formato="0.0%")
        escribir(ws_res, r, 4, INFO[k][1])
    escribir(ws_res, 12, 1, "Total", F_BOLD)
    formula(ws_res, 12, 2, "=SUM(B5:B11)", F_BOLD, "#,##0")
    formula(ws_res, 12, 3, "=SUM(C5:C11)", F_BOLD, "0.0%")
    escribir(ws_res, 12, 4, None)

    ws_res.cell(row=14, column=1, value="Indicadores").font = F_SECCION
    encabezado_tabla(ws_res, 15, 1, ["Indicador", "Valor", "Cómo se calcula"])
    indicadores = [
        (f"Registros en {A}", f"=COUNT({ref_hoja(h_a)}!$A$5:$A${fin_a})", "#,##0", f"Filas de la hoja {h_a}"),
        (f"Registros en {B}", f"=COUNT({ref_hoja(h_b)}!$A$5:$A${fin_b})", "#,##0", f"Filas de la hoja {h_b}"),
        ("Llaves cruzadas (en ambos)", f'=COUNTIFS({rng(col["fa"])},"<>",{rng(col["fb"])},"<>")', "#,##0",
         "Registros con fila en ambos archivos"),
        ("% de cuadre", "=IF(B18=0,0,B5/B18)", "0.0%", "Coinciden ÷ llaves cruzadas"),
        (f"Cobertura de {A} en {B}", "=IF(B16=0,0,B18/B16)", "0.0%", f"Llaves cruzadas ÷ registros en {A}"),
        (f"Cobertura de {B} en {A}", "=IF(B17=0,0,B18/B17)", "0.0%", f"Llaves cruzadas ÷ registros en {B}"),
        ("Registros con al menos 1 diferencia", f'=COUNTIF({rng(col["nd"])},">0")', "#,##0", "N° diferencias > 0"),
        ("Total de diferencias encontradas", f"=SUM({rng(col['nd'])})", "#,##0", "Suma de N° diferencias"),
        ("Registros pendientes de revisar", f'=COUNTIF({rng(col["rev"])},"Pendiente")+COUNTIF({rng(col["rev"])},"En revisión")',
         "#,##0", "Columna Revisión de la hoja Cuadre"),
    ]
    for i, (nombre, f_, formato, como) in enumerate(indicadores):
        r = 16 + i
        escribir(ws_res, r, 1, nombre, F_BOLD)
        formula(ws_res, r, 2, f_, F_BOLD, formato)
        escribir(ws_res, r, 3, como, F_SUB)

    r_comp = 16 + len(indicadores) + 2
    ws_res.cell(row=r_comp, column=1, value="Resultado por columna comparada").font = F_SECCION
    encabezado_tabla(ws_res, r_comp + 1, 1, ["Columna", f"Columna en {B}", "Tipo", "Modo", "Tolerancia",
                                              f"{R_OK} Coinciden", f"{R_NO} Diferentes", R_VACIO, "% coincidencia",
                                              "Suma de diferencias", "Mayor diferencia abs."])
    for i, c in enumerate(comps):
        r = r_comp + 2 + i
        lr = get_column_letter(c["c_res"])
        escribir_dato(ws_res, r, 1, c["etiqueta"])
        escribir_dato(ws_res, r, 2, c["Columna B"])
        escribir(ws_res, r, 3, c["Tipo"])
        escribir(ws_res, r, 4, c["Modo"])
        formula(ws_res, r, 5, f"={c['tol_ref']}", formato="0.00")
        formula(ws_res, r, 6, f'=COUNTIF({rng(lr)},"{R_OK}")', formato="#,##0")
        formula(ws_res, r, 7, f'=COUNTIF({rng(lr)},"{R_NO}")', formato="#,##0")
        formula(ws_res, r, 8, f'=COUNTIF({rng(lr)},"{R_VACIO}")', formato="#,##0")
        formula(ws_res, r, 9, f"=IF(F{r}+G{r}=0,0,F{r}/(F{r}+G{r}))", formato="0.0%")
        if "c_dif" in c:
            ld = get_column_letter(c["c_dif"])
            formula(ws_res, r, 10, f"=SUM({rng(ld)})", formato="#,##0.00")
            formula(ws_res, r, 11, f"=MAX(MAX({rng(ld)}),-MIN({rng(ld)}))", formato="#,##0.00")
        else:
            escribir(ws_res, r, 10, "—")
            escribir(ws_res, r, 11, "—")
    if comps:
        ws_res.conditional_formatting.add(
            f"I{r_comp + 2}:I{r_comp + 1 + len(comps)}",
            CellIsRule(operator="lessThan", formula=["0.9"], font=Font(name=FUENTE, bold=True, color="C00000")))
    else:
        escribir(ws_res, r_comp + 2, 1, "Sin columnas comparadas.")

    r_rev = r_comp + 3 + max(len(comps), 1) + 1
    ws_res.cell(row=r_rev, column=1, value="Seguimiento de la revisión").font = F_SECCION
    encabezado_tabla(ws_res, r_rev + 1, 1, ["Revisión", "Registros"])
    for i, v in enumerate(REVISION):
        r = r_rev + 2 + i
        escribir(ws_res, r, 1, v, F_BOLD)
        formula(ws_res, r, 2, f"=COUNTIF({rng(col['rev'])},{q(v)})", formato="#,##0")

    for col_l, w in zip("ABCDEFGHIJK", (38, 16, 40, 12, 12, 13, 13, 10, 14, 18, 20)):
        ws_res.column_dimensions[col_l].width = w

    pie = PieChart()
    pie.title = "Registros por estatus"
    pie.add_data(Reference(ws_res, min_col=2, min_row=4, max_row=11), titles_from_data=True)
    pie.set_categories(Reference(ws_res, min_col=1, min_row=5, max_row=11))
    pie.dataLabels = DataLabelList()
    pie.dataLabels.showPercent = True
    pie.dataLabels.showVal = pie.dataLabels.showCatName = pie.dataLabels.showSerName = False
    pie.dataLabels.showLegendKey = False
    for i, k in enumerate(ORDEN):
        pt = DataPoint(idx=i)
        pt.graphicalProperties.solidFill = INFO[k][0]
        pt.graphicalProperties.line.solidFill = "FFFFFF"
        pie.series[0].dPt.append(pt)
    pie.height, pie.width = 9, 13
    ws_res.add_chart(pie, "M3")
    if comps:
        bar = BarChart()
        bar.type = "bar"
        bar.grouping = "stacked"
        bar.overlap = 100
        bar.title = "Resultado por columna"
        bar.add_data(Reference(ws_res, min_col=6, max_col=8, min_row=r_comp + 1, max_row=r_comp + 1 + len(comps)),
                     titles_from_data=True)
        bar.set_categories(Reference(ws_res, min_col=1, min_row=r_comp + 2, max_row=r_comp + 1 + len(comps)))
        for serie, color in zip(bar.series, ("70AD47", "E06666", "FFD966")):
            serie.graphicalProperties.solidFill = color
            serie.graphicalProperties.line.solidFill = color
        bar.height, bar.width = 7 + 0.4 * len(comps), 13
        ws_res.add_chart(bar, "M22")

    for ws in wb.worksheets:
        ws.page_setup.orientation = "landscape"
        ws.page_setup.paperSize = ws.PAPERSIZE_LETTER
        ws.page_setup.fitToWidth, ws.page_setup.fitToHeight = 1, 0
        ws.sheet_properties.pageSetUpPr.fitToPage = True
        ws.page_margins.left = ws.page_margins.right = 0.4
        if ws.title != "Resumen":
            ws.print_title_rows = "3:4" if ws.title == "Cuadre" else "4:4"
        ws.oddFooter.center.text = "&A · página &P de &N"
        ws.sheet_properties.tabColor = {"Resumen": "1F4E78", "Cuadre": "C00000", "Observaciones": "C55A11",
                                        "Posibles coincidencias": "2F5597", "Duplicados": "7030A0",
                                        "Configuración": "BF8F00", "Leyenda": "548235"}.get(ws.title, "595959")
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# Interfaz
# ---------------------------------------------------------------------------

def panel_archivo(etiqueta, clave):
    archivo = st.file_uploader(f"Archivo {etiqueta}", type=["xlsx", "xlsm", "xls", "csv"], key=f"{P}up_{clave}")
    if not archivo:
        return None, None
    contenido = archivo.getvalue()
    try:
        hojas = hojas_de(contenido, archivo.name)
    except Exception as e:  # noqa: BLE001
        st.error(f"No pude abrir el archivo: {e}")
        return None, None
    hoja = st.selectbox("Hoja", hojas, key=f"{P}hoja_{clave}_{archivo.name}") if len(hojas) > 1 else hojas[0]
    with st.spinner("Leyendo..."):
        crudo = leer_crudo(contenido, archivo.name, hoja)
    if crudo.empty:
        st.warning("La hoja está vacía.")
        return None, None

    pos_auto = detectar_encabezado(crudo)
    fila_enc = st.number_input("Fila de encabezados (como se ve en Excel)", min_value=1,
                               value=int(crudo.index[pos_auto]) + 1, step=1,
                               key=f"{P}enc_{clave}_{archivo.name}_{hoja}",
                               help="Detectada automáticamente; cámbiala si no es correcta.")
    posiciones = [p for p, idx in enumerate(crudo.index) if idx + 1 >= fila_enc]
    if not posiciones:
        st.error("Esa fila está fuera del rango de datos.")
        return None, None
    df = armar_tabla(crudo, posiciones[0])

    with st.expander("Filtrar filas (opcional)"):
        col_f = st.selectbox("Columna", ["(sin filtro)"] + list(df.columns), key=f"{P}fcol_{clave}")
        if col_f != "(sin filtro)":
            valores = sorted(df[col_f].map(fmt).unique().tolist())[:500]
            elegidos = st.multiselect("Conservar solo estos valores", valores, key=f"{P}fval_{clave}_{col_f}")
            if elegidos:
                df = df[df[col_f].map(fmt).isin(elegidos)]
    st.caption(f"{len(df):,} filas · {df.shape[1]} columnas")
    with st.expander("Vista previa"):
        st.dataframe(para_mostrar(df.head(20)), width="stretch")
    return df, f"{archivo.name} / {hoja}"


def estilo_vista(df, L):
    colores = {L[k]: "#" + INFO[k][0] for k in ORDEN}
    res_cols = [c for c in df.columns if c.startswith("Resultado ")]

    def c_est(v):
        return f"background-color: {colores[v]}; color: #000; font-weight: 600" if v in colores else ""

    def c_res(v):
        return {R_NO: "color: #9C0006; font-weight: 700", R_OK: "color: #006100; font-weight: 700",
                R_VACIO: "color: #9C5700; font-weight: 700"}.get(v, "")
    sty = df.style
    aplicar = getattr(sty, "map", None) or sty.applymap
    sty = aplicar(c_est, subset=["Estatus"])
    if res_cols:
        sty = (getattr(sty, "map", None) or sty.applymap)(c_res, subset=res_cols)
    return sty


def mostrar_resultados(R, excel, nombre_archivo, cfg_json):
    L, conteo, A, B = R["labels"], R["conteo"], R["cfg"]["nombre_a"], R["cfg"]["nombre_b"]
    st.subheader("Resultados")
    m1 = st.columns(4)
    m1[0].metric("✅ Coinciden", f"{conteo['ok']:,}")
    m1[1].metric("🟠 Con diferencias", f"{conteo['dif']:,}")
    m1[2].metric(f"🔴 Solo en {A}", f"{conteo['solo_a']:,}")
    m1[3].metric(f"🔴 Solo en {B}", f"{conteo['solo_b']:,}")
    m2 = st.columns(4)
    m2[0].metric("🟡 Incompletos", f"{conteo['inc']:,}")
    m2[1].metric("🟣 Duplicados", f"{conteo['dup']:,}")
    m2[2].metric("⚪ Sin llave", f"{conteo['sin']:,}")
    m2[3].metric("% de cuadre", f"{conteo['ok'] / R['cruzadas']:.1%}" if R["cruzadas"] else "—",
                 help="Coinciden ÷ registros cuya llave cruzó en ambos archivos.")

    d1, d2 = st.columns([2, 1])
    d1.download_button("⬇️ Descargar Excel del cuadre", data=excel, file_name=nombre_archivo,
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                       type="primary", width="stretch")
    d2.download_button("💾 Guardar configuración", data=cfg_json, file_name="config_cuadre.json",
                       mime="application/json", width="stretch",
                       help="Súbela la próxima vez para repetir este cuadre con un clic.")
    if R["fuzzy_omitida"]:
        st.info("La búsqueda de posibles errores de captura se omitió por la cantidad de llaves sin cruce.")

    t1, t2, t3, t4, t5 = st.tabs(["📊 Resumen", "📋 Detalle", "🔎 Observaciones",
                                  f"🧩 Posibles coincidencias ({len(R['posibles'])})",
                                  f"♻️ Duplicados ({len(R['duplicados'])})"])
    with t1:
        resumen = pd.DataFrame({"Estatus": [L[k] for k in ORDEN], "Registros": [conteo[k] for k in ORDEN],
                                "Prioridad": [INFO[k][1] for k in ORDEN], "Qué hacer": [INFO[k][3] for k in ORDEN]})
        resumen = resumen[resumen["Registros"] > 0]
        st.bar_chart(resumen.set_index("Estatus")["Registros"], horizontal=True)
        st.dataframe(resumen, hide_index=True, width="stretch")
        if not R["stats"].empty:
            st.markdown("**Resultado por columna comparada**")
            st.dataframe(R["stats"], hide_index=True, width="stretch")
    with t2:
        vista = R["vista"]
        f1, f2 = st.columns([2, 1])
        presentes = [L[k] for k in ORDEN if conteo[k]]
        sel = f1.multiselect("Estatus", presentes, default=presentes, key=f"{P}f_est")
        buscar = f2.text_input("Buscar texto", key=f"{P}f_txt")
        v = vista[vista["Estatus"].isin(sel)]
        if buscar:
            v = v[v.astype(str).apply(lambda s: s.str.contains(buscar, case=False, regex=False)).any(axis=1)]
        st.caption(f"{len(v):,} de {len(vista):,} registros")
        v = para_mostrar(v)
        st.dataframe(estilo_vista(v, L) if len(v) <= 20_000 else v, hide_index=True, width="stretch", height=480)
    with t3:
        ob = R["obs"]
        if ob.empty:
            st.success("Sin observaciones: todo cuadra.")
        else:
            frec = ob.groupby(["Tipo de observación", "Columna"]).size().reset_index(name="Veces")
            st.dataframe(frec.sort_values("Veces", ascending=False), hide_index=True, width="stretch")
            tipos = sorted(ob["Tipo de observación"].unique())
            tsel = st.multiselect("Tipo de observación", tipos, default=tipos, key=f"{P}f_obs")
            st.dataframe(para_mostrar(ob[ob["Tipo de observación"].isin(tsel)]), hide_index=True, width="stretch")
    with t4:
        if R["posibles"].empty:
            st.info("No se encontraron llaves parecidas." if R["cfg"]["fuzzy"] else "Búsqueda desactivada.")
        else:
            st.caption("Llaves sin cruce que difieren en muy pocos caracteres de una del otro archivo: "
                       "probablemente un error de captura.")
            st.dataframe(R["posibles"], hide_index=True, width="stretch")
    with t5:
        if R["duplicados"].empty:
            st.success("No hay llaves duplicadas.")
        else:
            st.caption("Se emparejaron por orden de aparición (1.ª con 1.ª, 2.ª con 2.ª…).")
            st.dataframe(R["duplicados"], hide_index=True, width="stretch")


def main():
    ss = st.session_state
    st.title("🔀 Cuadre de archivos")
    st.caption("Cruza dos archivos por una llave, compara columnas, genera observaciones y un Excel de trabajo.")

    with st.sidebar:
        st.header("⚙️ Opciones del cuadre")
        arch_cfg = st.file_uploader("Cargar configuración guardada (.json)", type=["json"], key=f"{P}cfg_up")
        cfg_g = {}
        if arch_cfg:
            try:
                cfg_g = json.loads(arch_cfg.getvalue().decode("utf-8"))
                st.success("Configuración cargada.")
            except (ValueError, UnicodeDecodeError):
                st.error("El archivo de configuración no es válido.")
        sello = hashlib.md5(json.dumps(cfg_g, sort_keys=True).encode()).hexdigest()[:8] if cfg_g else "0"
        nombre_a = st.text_input("Nombre del archivo A", cfg_g.get("nombre_a", "Archivo A"), key=f"{P}na_{sello}").strip() or "Archivo A"
        nombre_b = st.text_input("Nombre del archivo B", cfg_g.get("nombre_b", "Archivo B"), key=f"{P}nb_{sello}").strip() or "Archivo B"
        if nombre_a.lower() == nombre_b.lower():
            nombre_b += " (2)"
        with st.expander("🔑 Normalización de la llave", expanded=False):
            norm = st.selectbox("Limpiar la llave", NORMS_LLAVE,
                                index=NORMS_LLAVE.index(cfg_g.get("norm_llave", "Estándar"))
                                if cfg_g.get("norm_llave") in NORMS_LLAVE else 0, key=f"{P}norm_{sello}",
                                help="Estándar: mayúsculas, sin acentos ni espacios sobrantes. "
                                     "Solo letras y números: quita guiones, puntos y espacios. "
                                     "Solo dígitos: ideal para teléfonos.")
            sin_ceros = st.checkbox("Ignorar ceros a la izquierda", cfg_g.get("sin_ceros", False), key=f"{P}ceros_{sello}")
            ultimos = st.number_input("Usar solo los últimos N caracteres (0 = todos)", 0, 50,
                                      int(cfg_g.get("ultimos", 0)), key=f"{P}ult_{sello}",
                                      help="Ej. 10 para teléfonos: +52 312 123 4567 → 3121234567.")
        with st.expander("📋 Resultado", expanded=False):
            incluir_solo_b = st.checkbox("Incluir registros que solo están en B", cfg_g.get("incluir_solo_b", True),
                                         key=f"{P}solob_{sello}")
            fuzzy = st.checkbox("Buscar posibles errores de captura en llaves", cfg_g.get("fuzzy", True),
                                key=f"{P}fz_{sello}")
            max_cambios = st.slider("Máximo de caracteres distintos", 1, 3, int(cfg_g.get("max_cambios", 1)),
                                    key=f"{P}mc_{sello}", disabled=not fuzzy,
                                    help="1 detecta un dígito mal capturado o dos dígitos intercambiados. "
                                         "Más de 1 puede confundir llaves consecutivas (ej. VINs de la misma serie).")

    c1, c2 = st.columns(2)
    with c1:
        st.subheader(f"📄 {nombre_a}")
        df_a, origen_a = panel_archivo("A", "a")
    with c2:
        st.subheader(f"📄 {nombre_b}")
        df_b, origen_b = panel_archivo("B", "b")
    if df_a is None or df_b is None:
        st.info("Sube los dos archivos para continuar.")
        return
    if df_a.empty or df_b.empty:
        st.warning("Uno de los archivos quedó sin filas (revisa el filtro o la fila de encabezados).")
        return

    st.divider()
    opciones_modo = ["Automático", "Manual"] + (["Configuración guardada"] if cfg_g else [])
    modo = st.radio("¿Cómo quieres configurar el cuadre?", opciones_modo, horizontal=True,
                    index=len(opciones_modo) - 1 if cfg_g else 0, key=f"{P}modo_{sello}",
                    help="En todos los modos puedes editar la llave y las columnas antes de ejecutar.")
    cols_a, cols_b = list(df_a.columns), list(df_b.columns)
    firma_arch = f"{origen_a}_{origen_b}_{len(df_a)}_{len(df_b)}"
    opts_llave = (norm, sin_ceros, int(ultimos))

    # ---- 1. Llave
    st.subheader("1. Llave")
    def_a, def_b = [], []
    if modo == "Automático":
        with st.spinner("Buscando la columna llave..."):
            sa, sb = sugerir_llave(df_a, df_b, *opts_llave)
        if sa:
            def_a, def_b = [sa], [sb]
            st.success(f"Llave sugerida: **{sa}** ({nombre_a}) ↔ **{sb}** ({nombre_b})")
        else:
            st.warning("No encontré una llave clara; elígela manualmente.")
    elif modo == "Configuración guardada":
        def_a = [c for c in cfg_g.get("llave_a", []) if c in cols_a]
        def_b = [c for c in cfg_g.get("llave_b", []) if c in cols_b]
        if len(def_a) != len(cfg_g.get("llave_a", [])) or len(def_b) != len(cfg_g.get("llave_b", [])):
            st.warning("Algunas columnas llave de la configuración no existen en estos archivos.")
    k1, k2 = st.columns(2)
    key_a = k1.multiselect(f"Llave en {nombre_a}", cols_a, default=def_a, key=f"{P}ka_{modo}_{sello}_{firma_arch}",
                           help="Elige varias columnas para una llave compuesta (ej. Nombre + Fecha).")
    key_b = k2.multiselect(f"Llave en {nombre_b}", cols_b, default=def_b, key=f"{P}kb_{modo}_{sello}_{firma_arch}",
                           help="Mismo número de columnas y en el mismo orden que en A.")
    if not key_a or not key_b:
        st.info("Elige la columna llave en ambos archivos.")
        return
    if len(key_a) != len(key_b):
        st.error("La llave debe tener el mismo número de columnas en ambos archivos.")
        return

    ka = construir_llave(df_a, key_a, *opts_llave)
    kb = construir_llave(df_b, key_b, *opts_llave)
    with st.expander("🩺 Diagnóstico de la llave", expanded=True):
        st.dataframe(diagnostico_llave(ka, kb, nombre_a, nombre_b), width="stretch")
        ej = [(llv, k) for llv, k in zip(df_a[key_a[0]].head(50), ka.head(50)) if k][:3]
        if ej:
            st.caption("Ejemplo de normalización: " + " · ".join(f"`{fmt(o)}` → `{k}`" for o, k in ej))

    # ---- 2. Columnas a comparar
    st.subheader("2. Columnas a comparar")
    fila_vacia = {"Columna A": None, "Columna B": None, "Tipo": "Texto flexible", "Modo": "Igual", "Tolerancia": 0.0}
    if modo == "Automático":
        with st.spinner("Buscando columnas equivalentes..."):
            sugerencias = sugerir_comparaciones(df_a, df_b, ka, kb, list(key_a), list(key_b))
        if not sugerencias:
            st.warning("No encontré columnas equivalentes; agrégalas en la tabla.")
        base = pd.DataFrame(sugerencias or [fila_vacia])
    elif modo == "Configuración guardada":
        guardadas = [c for c in cfg_g.get("comparaciones", [])
                     if c.get("Columna A") in cols_a and c.get("Columna B") in cols_b]
        base = pd.DataFrame(guardadas or [fila_vacia])
    else:
        base = pd.DataFrame([fila_vacia])
    st.caption("**Texto flexible** ignora mayúsculas, acentos, espacios, guiones y ceros a la izquierda. "
               "**Contiene**: un valor dentro del otro (MG5 ≈ MG5 - 2027). **Similar**: % mínimo de parecido. "
               "**Tolerancia**: diferencia permitida en Número, días en Fecha o % en Similar.")
    editor = st.data_editor(
        base, num_rows="dynamic", width="stretch",
        key=f"{P}cmp_{modo}_{sello}_{firma_arch}_{'|'.join(key_a)}_{'|'.join(key_b)}",
        column_config={
            "Columna A": st.column_config.SelectboxColumn(f"Columna en {nombre_a}", options=cols_a, required=True),
            "Columna B": st.column_config.SelectboxColumn(f"Columna en {nombre_b}", options=cols_b, required=True),
            "Tipo": st.column_config.SelectboxColumn("Tipo", options=TIPOS, default="Texto flexible"),
            "Modo": st.column_config.SelectboxColumn("Modo", options=MODOS, default="Igual"),
            "Tolerancia": st.column_config.NumberColumn("Tolerancia", min_value=0.0, default=0.0, format="%.2f"),
        },
    )
    comparaciones = [{"Columna A": r["Columna A"], "Columna B": r["Columna B"], "Tipo": r.get("Tipo") or "Texto flexible",
                      "Modo": r.get("Modo") or "Igual",
                      "Tolerancia": 0.0 if es_vacio(r.get("Tolerancia")) else float(r["Tolerancia"])}
                     for r in editor.to_dict("records")
                     if r.get("Columna A") in cols_a and r.get("Columna B") in cols_b]

    # ---- 3. Columnas adicionales
    with st.expander("3. Columnas adicionales para mostrar en el resultado (opcional)"):
        e1, e2 = st.columns(2)
        extras_a = e1.multiselect(f"De {nombre_a}", [c for c in cols_a if c not in key_a],
                                  default=[c for c in cfg_g.get("extras_a", []) if c in cols_a and c not in key_a]
                                  if modo == "Configuración guardada" else [], key=f"{P}exa_{modo}_{sello}_{firma_arch}")
        extras_b = e2.multiselect(f"De {nombre_b}", [c for c in cols_b if c not in key_b],
                                  default=[c for c in cfg_g.get("extras_b", []) if c in cols_b and c not in key_b]
                                  if modo == "Configuración guardada" else [], key=f"{P}exb_{modo}_{sello}_{firma_arch}")

    cfg = {"nombre_a": nombre_a, "nombre_b": nombre_b, "llave_a": list(key_a), "llave_b": list(key_b),
           "comparaciones": comparaciones, "extras_a": list(extras_a), "extras_b": list(extras_b),
           "norm_llave": norm, "sin_ceros": bool(sin_ceros), "ultimos": int(ultimos),
           "incluir_solo_b": bool(incluir_solo_b), "fuzzy": bool(fuzzy), "max_cambios": int(max_cambios)}
    firma = hashlib.md5((json.dumps(cfg, sort_keys=True, default=str) + firma_arch).encode()).hexdigest()

    if st.button("▶️ Ejecutar cuadre", type="primary", width="stretch"):
        with st.spinner("Cuadrando y armando el Excel..."):
            R = ejecutar_cuadre(df_a, df_b, cfg)
            excel = generar_excel(R, df_a, df_b, (origen_a, origen_b))
        nombre = re.sub(r"[^\w\-]+", "_", f"Cuadre_{nombre_a}_vs_{nombre_b}_{datetime.date.today():%Y%m%d}") + ".xlsx"
        ss[f"{P}res"] = {"R": R, "excel": excel, "firma": firma, "nombre": nombre,
                         "cfg_json": json.dumps(cfg, ensure_ascii=False, indent=2, default=str)}

    guardado = ss.get(f"{P}res")
    if not guardado:
        return
    st.divider()
    if guardado["firma"] != firma:
        st.warning("Cambiaste archivos u opciones después del último cuadre: vuelve a ejecutarlo para actualizar.")
    mostrar_resultados(guardado["R"], guardado["excel"], guardado["nombre"], guardado["cfg_json"])


if __name__ == "__main__":
    try:
        st.set_page_config(page_title="Cuadre de archivos", page_icon="🔀", layout="wide")
    except Exception:  # noqa: BLE001  (ya configurada por la app principal)
        pass
    main()
