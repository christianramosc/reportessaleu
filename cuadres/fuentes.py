# -*- coding: utf-8 -*-
"""
Fuentes que nutren la base de seguimiento:

  · Bitácoras mensuales de Cetelem (JUNIO_2026_FINAL.xlsx, …): fuente OFICIAL de Cetelem
    (folio, estatus, mensualidad, plazo, teléfono). Se reconocen por la columna "Folio CCK".
  · "Bitácora de F&I" (Google Sheets o .xlsx): todas las financieras y
    estatus; aporta rechazos de otros bancos, Sale-U, comentarios de F&I y contratos.

Todo se consolida en "solicitudes" (una por folio o, si no hay folio, por cliente).
"""
import io
import re
import difflib
import datetime as dt

import pandas as pd

from cuadres.base import norm, de_serial

# Columnas que se guardan de las bitácoras de Cetelem (no se guarda la fecha de nacimiento).
CAP_COLS = ["Folio CCK", "STATUS", "Tipo de Persona", "Nombre Completo", "Teléfono Móvil",
            "Correo Electrónico", "Nombre del Vendedor", "Marca", "Vehículo", "Año Modelo", "Accesorios",
            "Tipo de Seguro", "Aseguradora", "Precio Venta (c/IVA)", "Enganche", "¿Tiene GAP?", "Monto GAP",
            "Garantía Extendida", "Plazo (meses)", "Tasa Interés Anual", "Monto Total a Financiar",
            "Mensualidad Mínima", "Mensualidad Máxima"]
CAP_EXTRA = ["Hoja", "Archivo", "Fecha de carga"]

MESES = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6, "JULIO": 7,
         "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10, "NOVIEMBRE": 11, "DICIEMBRE": 12}

ESTATUS = {
    "RECHAZADO": "Rechazado", "RECHAZADOS": "Rechazado", "RECHAZADAS": "Rechazado", "RECHAZADA": "Rechazado",
    "APROBADO": "Aprobado", "APROBADOS": "Aprobado", "APROBADAS": "Aprobado", "APROBADA": "Aprobado",
    "CONTRAPROPUESTA": "Contrapropuesta", "CONTRAPROPUESTAS": "Contrapropuesta",
    "FINANCIADOS": "Contrato", "FINANCIADO": "Contrato", "FINANCIADAS": "Contrato", "CONTRATO": "Contrato",
    "ANALISIS": "Análisis", "EN ANALISIS": "Análisis",
}
PRIORIDAD = {"Rechazado": 0, "Análisis": 1, "Contrapropuesta": 2, "Aprobado": 3, "Contrato": 4}


# =============================================================================
# Normalización
# =============================================================================
def vacio(v):
    return v is None or v == "" or (isinstance(v, float) and v != v)


def txt(v):
    if vacio(v):
        return None
    s = re.sub(r"\s+", " ", str(v)).strip()
    return s or None


def name_key(n):
    """Nombre comparable sin importar orden de palabras, acentos ni puntos (S.A DE C.V = SA DE CV)."""
    n = norm(n).replace(".", "")
    return " ".join(sorted(re.sub(r"[^A-Z ]", " ", n).split()))


def similar(a, b):
    return bool(a) and bool(b) and (a == b or difflib.SequenceMatcher(None, a, b).ratio() >= 0.9)


def to_num(v):
    if vacio(v):
        return None
    if isinstance(v, (int, float)):
        return float(v)
    s = re.sub(r"[^\d.\-]", "", str(v))
    try:
        return float(s) if s else None
    except ValueError:
        return None


def to_folio(v):
    if vacio(v):
        return None
    if isinstance(v, float):
        v = int(v)
    s = re.sub(r"\D", "", re.sub(r"\.0+$", "", str(v)))
    return int(s) if 7 <= len(s) <= 9 else None


def to_phone(v):
    if vacio(v):
        return None
    if isinstance(v, float):
        v = int(v)
    s = re.sub(r"\D", "", str(v or ""))
    return int(s) if len(s) == 10 else None


def to_date(v):
    v = de_serial(v)
    if isinstance(v, dt.datetime):
        return v.date()
    if isinstance(v, dt.date):
        return v
    if isinstance(v, str):
        # dd/mm/aaaa; también el error de tecleo "14/1072026" (un 7 en lugar de "/": Shift+7)
        m = re.fullmatch(r"\s*(\d{1,2})/(\d{1,2})[/7](\d{4})\s*", v)
        if m:
            try:
                return dt.date(int(m.group(3)), int(m.group(2)), int(m.group(1)))
            except ValueError:
                return None
    return None


def folio_date(folio, anio=None):
    """En los folios CCK los primeros dígitos son el día del año (27866028 → día 278)."""
    try:
        s = str(int(folio))
        dia = int(s[:-5])
        if 1 <= dia <= 366:
            return dt.date(anio or dt.date.today().year, 1, 1) + dt.timedelta(dia - 1)
    except (TypeError, ValueError):
        pass
    return None


def mail(v):
    v = txt(v)
    return v.lower() if v and "@" in v else None


def estatus(v):
    if not v:
        return None
    return ESTATUS.get(norm(v), str(v).strip().capitalize())


def anio_de(texto):
    m = re.search(r"20\d{2}", str(texto))
    return int(m.group()) if m else None


# =============================================================================
# Bitácoras mensuales de Cetelem
# =============================================================================
def leer_capturas(archivos):
    """archivos: lista de (nombre, bytes). Devuelve un DataFrame con CAP_COLS + Hoja + Archivo."""
    partes = []
    for nombre, contenido in archivos:
        xls = pd.ExcelFile(io.BytesIO(contenido))
        for hoja in xls.sheet_names:
            d = pd.read_excel(xls, sheet_name=hoja, dtype=str).fillna("")
            mapa = {norm(c): c for c in d.columns}
            if "FOLIO CCK" not in mapa:
                continue
            sub = pd.DataFrame({c: d[mapa[norm(c)]] if norm(c) in mapa else "" for c in CAP_COLS})
            sub = sub[sub["Folio CCK"].map(to_folio).notna()]
            sub["Hoja"] = hoja.strip()
            sub["Archivo"] = nombre
            partes.append(sub)
    if not partes:
        return pd.DataFrame(columns=CAP_COLS + CAP_EXTRA)
    df = pd.concat(partes, ignore_index=True)
    df["Fecha de carga"] = dt.datetime.now().replace(microsecond=0)
    return df


def combinar_capturas(guardadas, nuevas):
    """Las hojas (meses) que vienen en `nuevas` reemplazan completas a las guardadas del mismo mes."""
    if guardadas is None or guardadas.empty:
        return nuevas.copy()
    meses_nuevos = set(nuevas["Hoja"].map(norm))
    quedan = guardadas[~guardadas["Hoja"].map(norm).isin(meses_nuevos)]
    todo = pd.concat([quedan[[c for c in CAP_COLS + CAP_EXTRA if c in quedan.columns]], nuevas], ignore_index=True)
    return todo.fillna("")


def _orden_mes(hoja):
    t = norm(hoja)
    mes = next((MESES[m] for m in MESES if m in t), 0)
    return (anio_de(t) or 0) * 100 + mes


def capturas_a_archivos(df, carpeta):
    """Escribe un .xlsx por mes (formato original) para los módulos de cuadre que leen archivos."""
    import os
    rutas = []
    for hoja, g in df.groupby("Hoja", sort=False):
        nombre = re.sub(r"[^\w\- ]", "", str(hoja)).strip().replace(" ", "_") or "CETELEM"
        ruta = os.path.join(carpeta, f"{nombre}_CETELEM.xlsx")
        g[[c for c in CAP_COLS if c in g.columns]].to_excel(ruta, sheet_name=str(hoja)[:31], index=False)
        rutas.append(ruta)
    return rutas


def registros_capturas(df):
    out = []
    for _, r in df.iterrows():
        folio = to_folio(r.get("Folio CCK"))
        if not folio:
            continue
        anio = anio_de(r.get("Hoja")) or anio_de(r.get("Archivo"))
        tp = norm(r.get("Tipo de Persona")) or None
        out.append(dict(
            origen=f"CETELEM {norm(r.get('Hoja'))}", fecha=folio_date(folio, anio), folio=folio,
            cliente=txt(r.get("Nombre Completo")), telefono=to_phone(r.get("Teléfono Móvil")),
            correo=mail(r.get("Correo Electrónico")), vendedor=txt(r.get("Nombre del Vendedor")),
            unidad=txt(r.get("Vehículo")), anio=int(to_num(r.get("Año Modelo"))) if to_num(r.get("Año Modelo")) else None,
            banco="CETELEM", estatus=estatus(r.get("STATUS")), canal2=None, estatus_canal2=None,
            plazo=to_num(r.get("Plazo (meses)")),
            mensualidad=to_num(r.get("Mensualidad Máxima")) or to_num(r.get("Mensualidad Mínima")),
            monto=to_num(r.get("Monto Total a Financiar")), sale_u=None, comentario_fi=None, comentario_asesor=None,
            factura=None, vin=None, gap=norm(r.get("¿Tiene GAP?")) or None, accesorios=norm(r.get("Accesorios")) or None,
            tipo_persona=tp, key=name_key(r.get("Nombre Completo"))))
    return out


# =============================================================================
# "Bitácora de F&I"
# =============================================================================
def tablas_de_excel(contenido):
    import openpyxl
    wb = openpyxl.load_workbook(io.BytesIO(contenido), data_only=True)
    return {ws.title: [list(f) for f in ws.iter_rows(values_only=True)] for ws in wb.worksheets}


def tablas_de_google(cliente, url):
    from gspread.utils import ValueRenderOption, DateTimeOption
    libro = cliente.open_by_url(url)
    out = {}
    for ws in libro.worksheets():
        out[ws.title] = ws.get(value_render_option=ValueRenderOption.unformatted,
                               date_time_render_option=DateTimeOption.serial_number, pad_values=True)
    return out, libro.title


# Posiciones (1-based) del formato de captura de F&I (DETALLE, RECHAZ, ANF, CONTRAPROPUESTA…)
_POS_DETALLE = dict(vendedor=2, fecha=3, cliente=4, unidad=6, anio=7, telefono=9, banco=10, estatus=11,
                    comentario_fi=13, monto=14, folio=16, canal2=18, estatus_canal2=19, comentario_asesor=21,
                    sale_u=22, correo=27, plazo=28, factura=29, vin=30, mensualidad=31)
# Formato de cierre mensual (CIERRE AGOSTO…)
_POS_CIERRE = dict(vendedor=2, fecha=3, cliente=4, unidad=7, anio=8, telefono=10, banco=19, estatus=12,
                   comentario_fi=14, monto=15, folio=17, canal2=19, estatus_canal2=21, comentario_asesor=22,
                   sale_u=23, correo=28, plazo=29, factura=30, vin=31, mensualidad=32)


def _formato(encabezado):
    e = [norm(x) for x in (encabezado or [])] + [""] * 6
    if e[0] == "INGRESOS" and e[1] == "VENDEDOR" and e[3] == "CLIENTE":
        if e[4] == "MODELO":
            return "DETALLE"
        if e[4] == "PAGADA":
            return "CIERRE"
    return None


def registros_financiamiento(tablas):
    """tablas: {nombre_hoja: filas}. Lee solo las hojas con el formato de captura de F&I.
    Las copias de Cetelem que trae el archivo (TRIMESTRE, Hoja2) se ignoran: manda la bitácora de Cetelem."""
    out, hojas_usadas = [], []
    for hoja, filas in tablas.items():
        if not filas:
            continue
        fmt = _formato(filas[0])
        if not fmt:
            continue
        pos = _POS_DETALLE if fmt == "DETALLE" else _POS_CIERRE
        usadas = 0
        for n, fila in enumerate(filas[1:], start=2):
            g = lambda k: fila[pos[k] - 1] if len(fila) >= pos[k] else None
            cliente = txt(g("cliente"))
            if not cliente:
                continue
            fecha = to_date(g("fecha"))
            if fmt == "CIERRE" and not fecha:
                continue            # tablas de resumen al pie de la hoja
            if fmt == "DETALLE" and not txt(g("estatus")):
                continue
            banco = txt(g("banco"))
            if fmt == "CIERRE":
                banco = txt(fila[18]) if len(fila) > 18 and txt(fila[18]) else txt(fila[10]) if len(fila) > 10 else None
            anio = g("anio")
            out.append(dict(
                origen=f"{hoja} fila {n}", fecha=fecha, folio=to_folio(g("folio")), cliente=cliente,
                telefono=to_phone(g("telefono")), correo=mail(g("correo")), vendedor=txt(g("vendedor")),
                unidad=txt(g("unidad")), anio=int(anio) if isinstance(anio, (int, float)) and 2000 < anio < 2100 else None,
                banco=banco, estatus=estatus(g("estatus")), canal2=txt(g("canal2")), estatus_canal2=txt(g("estatus_canal2")),
                plazo=to_num(g("plazo")), mensualidad=to_num(g("mensualidad")), monto=to_num(g("monto")),
                sale_u=norm(g("sale_u")) if norm(g("sale_u")) in ("SI", "NO") else None,
                comentario_fi=txt(g("comentario_fi")), comentario_asesor=txt(g("comentario_asesor")),
                factura=txt(g("factura")), vin=txt(g("vin")), gap=None, accesorios=None, tipo_persona=None,
                key=name_key(cliente)))
            usadas += 1
        if usadas:
            hojas_usadas.append(f"{hoja} ({usadas})")
    return out, hojas_usadas


# =============================================================================
# Consolidación: una solicitud por folio
# =============================================================================
def consolidar(registros):
    grupos, por_folio = [], {}
    for r in registros:
        if r["folio"]:
            if r["folio"] not in por_folio:
                por_folio[r["folio"]] = []
                grupos.append(por_folio[r["folio"]])
            por_folio[r["folio"]].append(r)
    for r in registros:
        if r["folio"]:
            continue

        def cerca(g):
            fd = folio_date(g[0]["folio"], g[0]["fecha"].year if g[0]["fecha"] else None) if g[0]["folio"] else None
            return not (r["fecha"] and fd) or abs((r["fecha"] - fd).days) <= 20
        cand = [g for g in grupos if any(similar(r["key"], x["key"]) for x in g) and cerca(g)]
        if len(cand) == 1:
            cand[0].append(r)
        else:
            grupos.append([r])

    out = []
    for g in grupos:
        es_cet = lambda x: x["origen"].startswith("CETELEM")

        def primero(campo, preferir_cetelem=False):
            seq = sorted(g, key=lambda x: 0 if es_cet(x) else 1) if preferir_cetelem else g
            return next((x[campo] for x in seq if x.get(campo) not in (None, "")), None)

        internas = [x for x in g if not es_cet(x)]
        cetelem = [x for x in g if es_cet(x) and x["estatus"]]
        vigentes = cetelem + internas
        fecha_int = next((x["fecha"] for x in internas if x["fecha"]), None)
        folio = primero("folio")
        anio = (fecha_int or dt.date.today()).year
        if cetelem:
            est_sol = cetelem[-1]["estatus"]          # el estatus oficial de Cetelem manda
        else:
            est_sol = "Rechazado" if any(x["estatus"] == "Rechazado" for x in g) else primero("estatus")
        estados = {x["estatus"] for x in vigentes if x["estatus"]}
        conflicto = ("; ".join(f"{x['origen'].split(' fila')[0]}: {x['estatus']}" for x in vigentes if x["estatus"])
                     if len(estados) > 1 else None)
        finales = [x for x in vigentes if x["estatus"] in PRIORIDAD]
        fin = max(finales, key=lambda x: PRIORIDAD[x["estatus"]]) if finales else None
        fin_banco = None
        if fin:
            c2 = fin.get("canal2")
            fin_banco = c2 if c2 and norm(c2) != "OTROS" else fin.get("banco")
            if not fin_banco or norm(fin_banco) == "OTROS":
                fin_banco = "otra financiera"
        banco = primero("banco", True)
        if banco and norm(banco) == "OTROS":
            banco = next((x["canal2"] for x in g if x.get("canal2")), "OTROS")
        variantes = [x["cliente"].upper() for x in g if x.get("cliente")]
        con_acento = [v for v in variantes if re.search(r"[ÑÁÉÍÓÚ]", v)]
        out.append(dict(
            fecha=(folio_date(folio, anio) if folio else None) or fecha_int,
            folio=folio, cliente=(con_acento or [primero("cliente", True) or ""])[0],
            key=g[0]["key"], keys={x["key"] for x in g},
            telefono=primero("telefono", True), correo=primero("correo", True), vendedor=primero("vendedor", True),
            unidad=primero("unidad", True), anio=primero("anio"), banco=banco, est_sol=est_sol,
            canal2=primero("canal2"), est_fin=fin["estatus"] if fin else None, fin_banco=fin_banco,
            plazo=primero("plazo", True), mensualidad=primero("mensualidad", True), monto=primero("monto", True),
            sale_u=primero("sale_u"), gap=primero("gap", True), accesorios=primero("accesorios", True),
            tipo_persona=primero("tipo_persona", True),
            com_fi="; ".join(dict.fromkeys(x["comentario_fi"] for x in g if x.get("comentario_fi"))) or None,
            com_as="; ".join(dict.fromkeys(x["comentario_asesor"] for x in g if x.get("comentario_asesor"))) or None,
            factura=primero("factura"), vin=primero("vin"), conflicto=conflicto,
            fuentes=", ".join(x["origen"] for x in g), origenes=g,
            mes_cetelem=cetelem[-1]["origen"].replace("CETELEM ", "").title() if cetelem else None))
    out.sort(key=lambda x: (x["fecha"] or dt.date(2100, 1, 1), x["cliente"] or ""))
    return out


BITACORA_COLS = ["Fecha", "Folio", "Cliente", "Teléfono", "Correo", "Vendedor", "Unidad", "Año", "Banco solicitud",
                 "Estatus solicitud", "Canal 2", "Estatus final", "Financiera final", "Plazo", "Mensualidad",
                 "Monto a financiar", "Sale-U", "Comentarios F&I", "Comentarios asesor", "# Factura", "VIN",
                 "Fila en 'MG Colima Rechazos'", "Fuente en la bitácora"]


def filas_bitacora(solicitudes, filas_en_rechazos):
    """filas_en_rechazos: {id(solicitud): [filas]}"""
    out = []
    for s in solicitudes:
        out.append([s["fecha"], s["folio"], s["cliente"], s["telefono"], s["correo"], s["vendedor"], s["unidad"],
                    s["anio"], s["banco"], s["est_sol"], s["canal2"], s["est_fin"], s["fin_banco"], s["plazo"],
                    s["mensualidad"], s["monto"], s["sale_u"], s["com_fi"], s["com_as"], s["factura"], s["vin"],
                    ", ".join(map(str, filas_en_rechazos.get(id(s), []))), s["fuentes"]])
    return out
