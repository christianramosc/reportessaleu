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
    # FINANCIADO (Cetelem) y CONTRATO (F&I) significan lo mismo: el crédito se formalizó
    "FINANCIADOS": "Formalizado", "FINANCIADO": "Formalizado", "FINANCIADAS": "Formalizado",
    "CONTRATO": "Formalizado", "CONTRATOS": "Formalizado", "FORMALIZADO": "Formalizado", "FORMALIZADOS": "Formalizado",
    "ANALISIS": "Análisis", "EN ANALISIS": "Análisis",
}
PRIORIDAD = {"Rechazado": 0, "Análisis": 1, "Contrapropuesta": 2, "Aprobado": 3, "Formalizado": 4}


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
    """Teléfono a número de 10 dígitos (o None). Misma limpieza que base.tel_texto."""
    from cuadres.base import tel_texto
    s = tel_texto(v)
    return int(s) if s else None


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


def etiqueta_hoja(v):
    """Nombre del mes de una captura ("JUNIO 2026"). Si Google lo guardó como fecha (01/06/2026), se regresa a texto."""
    if isinstance(v, (int, float)) and not isinstance(v, bool) and 30000 < v < 80000:
        v = de_serial(v)                     # desde la API de Google llega como número de serie
    if isinstance(v, (dt.date, dt.datetime)):
        return f"{ETIQUETA_MES[v.month]} {v.year}"
    return "" if vacio(v) else str(v).strip()


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


COL_SALEU_BITACORA = "Registro en Sale-U"          # antes «Sale-U» (SI/NO de F&I)
COL_ID_RECHAZOS = "ID en 'MG Colima Rechazos'"      # antes «Fila en 'MG Colima Rechazos'» (el número de fila)
COL_FILA_VIEJA = "Fila en 'MG Colima Rechazos'"
BITACORA_COLS = ["Fecha", "Cliente", "Teléfono", "Folio", "Correo", "Vendedor", "Unidad", "Año", "Banco solicitud",
                 "Estatus solicitud", "Canal 2", "Estatus final", "Financiera final", "Plazo", "Mensualidad",
                 "Monto a financiar", COL_SALEU_BITACORA, "Comentarios F&I", "Comentarios asesor", "# Factura", "VIN",
                 COL_ID_RECHAZOS, "Fuente en la bitácora"]


def filas_bitacora(solicitudes, filas_en_rechazos):
    """filas_en_rechazos: {id(solicitud): [IDs de MG Colima Rechazos]} (o números de fila si aún no hay IDs)"""
    out = []
    for s in solicitudes:
        out.append([s["fecha"], s["cliente"], s["telefono"], s["folio"], s["correo"], s["vendedor"], s["unidad"],
                    s["anio"], s["banco"], s["est_sol"], s["canal2"], s["est_fin"], s["fin_banco"], s["plazo"],
                    s["mensualidad"], s["monto"], s["sale_u"], s["com_fi"], s["com_as"], s["factura"], s["vin"],
                    ", ".join(map(str, filas_en_rechazos.get(id(s), []))), s["fuentes"]])
    return out


# =============================================================================
# Pestaña "Bitacora" → solicitudes para los cuadres (todas las financieras)
# =============================================================================
ETIQUETA_MES = {v: k for k, v in MESES.items() if k != "SETIEMBRE"}


def _estatus_bitacora(v):
    e = estatus(v) if not vacio(v) else None
    return e or (txt(v) or None)


def _financiera(v):
    t = txt(v)
    return None if not t or norm(t) == "OTROS" else t.upper()


def solicitudes_de_bitacora(df):
    """df: pestaña Bitacora leída de la base (BITACORA_COLS). Devuelve una fila por solicitud con:
    Folio, Nombre, Telefono, Vendedor, Vehiculo, Monto_Financiar, Fecha, Mes, _anio, _mes,
    Est_sol / Fin_sol (estatus y financiera de la solicitud) y Est_fin / Fin_fin (resultado final),
    con los estatus en mayúsculas sin acentos (RECHAZADO, APROBADO, FORMALIZADO, CONTRAPROPUESTA, ANALISIS)."""
    cols = ["Folio", "Nombre", "Telefono", "Vendedor", "Vehiculo", "Monto_Financiar", "Fecha", "Mes", "_anio", "_mes",
            "Est_sol", "Fin_sol", "Est_fin", "Fin_fin", "Fuente", "Plazo", "Mensualidad", "Correo", "Fila_rechazos"]
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    g = lambda r, c: r[c] if c in r.index else None
    filas = []
    for _, r in df.iterrows():
        nombre = txt(g(r, "Cliente"))
        if not nombre:
            continue
        fecha = to_date(g(r, "Fecha"))
        folio = to_folio(g(r, "Folio"))
        est_sol, est_fin = _estatus_bitacora(g(r, "Estatus solicitud")), _estatus_bitacora(g(r, "Estatus final"))
        fin_sol = _financiera(g(r, "Banco solicitud")) or _financiera(g(r, "Canal 2")) or "OTRA FINANCIERA"
        fin_fin = _financiera(g(r, "Financiera final")) or fin_sol
        tel = to_phone(g(r, "Teléfono"))
        monto = to_num(g(r, "Monto a financiar"))
        filas.append(dict(
            Folio=str(folio) if folio else "", Nombre=" ".join(nombre.split()), Telefono=str(tel) if tel else "",
            Vendedor=txt(g(r, "Vendedor")) or "", Vehiculo=txt(g(r, "Unidad")) or "",
            Monto_Financiar=f"{monto:,.2f}" if monto else "", Fecha=fecha,
            Mes=f"{ETIQUETA_MES[fecha.month][:3]} {fecha.year}" if fecha else "SIN FECHA",
            _anio=fecha.year if fecha else 0, _mes=fecha.month if fecha else 0,
            Est_sol=norm(est_sol), Fin_sol=fin_sol, Est_fin=norm(est_fin or est_sol), Fin_fin=fin_fin,
            Fuente=txt(g(r, "Fuente en la bitácora")) or "",
            Plazo=int(to_num(g(r, "Plazo"))) if to_num(g(r, "Plazo")) else None,
            Mensualidad=round(to_num(g(r, "Mensualidad")), 2) if to_num(g(r, "Mensualidad")) else None,
            Correo=mail(g(r, "Correo")) or "",
            Fila_rechazos=str(g(r, COL_ID_RECHAZOS) or g(r, COL_FILA_VIEJA) or "")))
    return pd.DataFrame(filas, columns=cols)


def financiera_igual(a, b):
    """CETELEM ~ Cetelem, SCOTIA ~ SCOTIABANK, MSTAR ~ M STAR…"""
    a, b = re.sub(r"[^A-Z]", "", norm(a)), re.sub(r"[^A-Z]", "", norm(b))
    return bool(a and b) and (a in b or b in a)


# =============================================================================
# Copia de la Bitácora de F&I en la base (pestaña "Bitacora F&I")
# =============================================================================
FI_COLS = ["Origen", "Fecha", "Folio", "Cliente", "Teléfono", "Correo", "Vendedor", "Unidad", "Año", "Banco",
           "Estatus", "Canal 2", "Estatus canal 2", "Plazo", "Mensualidad", "Monto", "Sale-U", "Comentario F&I",
           "Comentario asesor", "# Factura", "VIN", "Fecha de carga"]
_FI_LLAVES = ["origen", "fecha", "folio", "cliente", "telefono", "correo", "vendedor", "unidad", "anio", "banco",
              "estatus", "canal2", "estatus_canal2", "plazo", "mensualidad", "monto", "sale_u", "comentario_fi",
              "comentario_asesor", "factura", "vin"]


def filas_fi(registros):
    ahora = dt.datetime.now().replace(microsecond=0)
    return [[r.get(k) for k in _FI_LLAVES] + [ahora] for r in registros]


def registros_fi_de_tabla(df):
    """Pestaña "Bitacora F&I" → registros como los de registros_financiamiento()."""
    out = []
    if df is None or df.empty:
        return out
    for _, f in df.iterrows():
        g = lambda c: None if c not in f.index or vacio(f[c]) else f[c]
        cliente = txt(g("Cliente"))
        if not cliente:
            continue
        anio = to_num(g("Año"))
        out.append(dict(
            origen=txt(g("Origen")) or "Bitacora F&I", fecha=to_date(g("Fecha")), folio=to_folio(g("Folio")),
            cliente=cliente, telefono=to_phone(g("Teléfono")), correo=mail(g("Correo")), vendedor=txt(g("Vendedor")),
            unidad=txt(g("Unidad")), anio=int(anio) if anio else None, banco=txt(g("Banco")),
            estatus=estatus(g("Estatus")), canal2=txt(g("Canal 2")), estatus_canal2=txt(g("Estatus canal 2")),
            plazo=to_num(g("Plazo")), mensualidad=to_num(g("Mensualidad")), monto=to_num(g("Monto")),
            sale_u=norm(g("Sale-U")) if norm(g("Sale-U")) in ("SI", "NO") else None,
            comentario_fi=txt(g("Comentario F&I")), comentario_asesor=txt(g("Comentario asesor")),
            factura=txt(g("# Factura")), vin=txt(g("VIN")), gap=None, accesorios=None, tipo_persona=None,
            key=name_key(cliente)))
    return out


def _llave_fi(r):
    """Misma solicitud de F&I: por folio; sin folio, mismo cliente, banco y fecha."""
    if r.get("folio"):
        return ("folio", r["folio"])
    return ("cliente", r.get("key") or name_key(r.get("cliente")), norm(r.get("banco")), r.get("fecha"))


def combinar_fi(guardados, nuevos):
    """Acumula la Bitácora de F&I: lo que viene en la carga nueva reemplaza al mismo registro guardado
    (mismo folio o mismo cliente/banco/fecha) y lo demás guardado se conserva (meses que ya no trae el archivo)."""
    llaves = {_llave_fi(r) for r in nuevos}
    return [r for r in guardados if _llave_fi(r) not in llaves] + list(nuevos)
