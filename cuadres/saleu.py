# -*- coding: utf-8 -*-
"""
cuadre Sale-U vs el Sheet
-------------------------
Cuadre de créditos: Sale-U (reporteLeadsconCredito*.csv/xlsx) vs. los clientes de la base de seguimiento
(«MG Colima Rechazos», todas las financieras, y «MG Colima Créditos Finalizados»; ver
saleu_sheet.solicitudes_del_sheet). Acepta cualquier tabla con la forma de fuentes.solicitudes_de_bitacora().
Sin base, también acepta los Excel mensuales de Cetelem (JULIO_2026_FINAL.xlsx, …).

Lógica de cruce (por cliente):
  A) Teléfono igual + nombre parecido            -> Coincide
  B) Nombre muy parecido, teléfono distinto      -> Coincide por nombre (revisar teléfono)
  C) Teléfono igual pero nombre distinto         -> Coincide por teléfono (revisar nombre)
  D) Nombre medianamente parecido                -> Posible coincidencia (revisar)
  Se asigna primero lo más confiable (A), luego B, C y D, para que un teléfono
  mal capturado en Sale-U no "robe" el registro de otro cliente.

Con el reporte Fuente (opcional, ver saleu_fuente.py) los que faltan se separan en «Dar de alta» (no están en
Sale-U) y «Registrar crédito» (ya son lead, falta su crédito).

Salidas: Cuadre_SaleU_<periodo>_<fecha>.xlsx  y  .pdf (formato común en cuadres/salidas.py)

Uso en Colab: ejecuta la celda y sube todos los archivos (el reporte de Sale-U
y los Excel mensuales de Cetelem) cuando te lo pida.
"""

import io, os, re, glob, unicodedata, subprocess, sys
from datetime import date, datetime
from difflib import SequenceMatcher
import pandas as pd

from cuadres.fuentes import financiera_igual, folio_date, to_folio
from cuadres import salidas as sal

# =============================================================================
# CONFIGURACIÓN
# =============================================================================
PDV = "MG COLIMA"
UMBRAL_NOMBRE_CON_TEL = 0.50   # parecido mínimo de nombre si el teléfono coincide
UMBRAL_NOMBRE_FUERTE  = 0.80   # parecido para aceptar coincidencia SOLO por nombre
UMBRAL_NOMBRE_POSIBLE = 0.60   # parecido para marcar "posible coincidencia"
SIMILITUD_PALABRA     = 0.80   # dos palabras se consideran iguales (Monserrat/Montserrat)

# Cómo debería verse en Sale-U cada estatus de la bitácora
ESTATUS_ESPERADO = {
    "APROBADO": "ACEPTADO",
    "FINANCIADO": "ACEPTADO",
    "FINANCIADOS": "ACEPTADO",
    "FORMALIZADO": "ACEPTADO",
    "RECHAZADO": "RECHAZADO",
    "FINALIZADO": "ACEPTADO",  # MG Colima Créditos Finalizados
    "CONTRAPROPUESTA": None,   # None = no se juzga, se marca para revisión
    "ANALISIS": None,
}

MESES = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
         "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10,
         "NOVIEMBRE": 11, "DICIEMBRE": 12}

PALABRAS_VACIAS = {"DE", "DEL", "LA", "LAS", "LOS", "Y", "SA", "CV", "SAPI", "RL", "MA"}


# =============================================================================
# 1. CARGA DE ARCHIVOS
# =============================================================================
def obtener_archivos():
    """En Colab pide subir archivos; fuera de Colab usa la carpeta actual."""
    try:
        from google.colab import files  # noqa
        print("Sube el reporte de Sale-U y los Excel mensuales de Cetelem:")
        subidos = files.upload()
        rutas = list(subidos.keys())
    except ImportError:
        rutas = glob.glob("*.csv") + glob.glob("*.xlsx")
    saleu = [r for r in rutas if "lead" in r.lower() or "saleu" in r.lower()]
    cetelem = [r for r in rutas if r.lower().endswith((".xlsx", ".xls"))
               and r not in saleu and not r.startswith("Cuadre_SaleU")]
    if not saleu:
        raise FileNotFoundError("No encontré el reporte de Sale-U (nombre con 'lead').")
    if not cetelem:
        raise FileNotFoundError("No encontré capturas mensuales de Cetelem.")
    return saleu[0], cetelem


def detectar_mes(nombre_archivo, nombre_hoja):
    texto = f"{nombre_hoja} {nombre_archivo}".upper()
    mes = next((MESES[m] for m in MESES if m in texto), 0)
    anio = re.search(r"20\d{2}", texto)
    anio = int(anio.group()) if anio else datetime.now().year
    etiqueta = next((m for m in MESES if MESES[m] == mes and m != "SETIEMBRE"), "SIN MES")
    return anio * 100 + mes, f"{etiqueta[:3]} {anio}"


def _fila_encabezado(filas):
    """Busca la fila que tiene los títulos reales (la que contiene 'Nombre del Cliente')."""
    for i, fila in enumerate(filas):
        texto = quitar_acentos(" ".join(str(x) for x in fila)).upper()
        if "NOMBRE DEL CLIENTE" in texto and "TELEFONO" in texto:
            return i
    raise ValueError("No encontré la fila de encabezados ('Nombre del Cliente', 'Teléfono') en Sale-U.")


def leer_saleu(ruta):
    """Lee el reporte de Sale-U tal cual se descarga: ignora las líneas de título
    ('Leads Crédito', 'Agencia:', 'Rango:', renglones vacíos) y arranca en los encabezados."""
    if ruta.lower().endswith(".csv"):
        import csv
        texto = None
        for enc in ("utf-8-sig", "latin-1"):
            try:
                with open(ruta, encoding=enc, newline="") as f:
                    texto = f.read()
                break
            except UnicodeDecodeError:
                continue
        filas = list(csv.reader(io.StringIO(texto)))
        h = _fila_encabezado(filas)
        info = " ".join(" ".join(f) for f in filas[:h])
        df = pd.read_csv(io.StringIO(texto), dtype=str, skiprows=h, skip_blank_lines=False,
                         keep_default_na=False)
    else:
        crudo = pd.read_excel(ruta, header=None, dtype=str).fillna("")
        h = _fila_encabezado(crudo.values.tolist())
        info = " ".join(" ".join(map(str, f)) for f in crudo.values.tolist()[:h])
        df = crudo.iloc[h + 1:].copy()
        df.columns = crudo.iloc[h].tolist()
    info = re.sub(r"\s+", " ", info).strip()
    if info:
        print(f"Encabezado del reporte Sale-U: {info}")
    df = df[[c for c in df.columns if str(c).strip()]]          # columnas sin título
    df = df[df.apply(lambda r: "".join(map(str, r)).strip() != "", axis=1)]   # renglones vacíos
    df.columns = df.columns.str.strip()
    ren = {}
    for c in df.columns:
        cu = quitar_acentos(c).upper()
        if "NOMBRE DEL CLIENTE" in cu: ren[c] = "Nombre"
        elif "TELEFONO" in cu: ren[c] = "Telefono"
        elif cu == "ASESOR": ren[c] = "Asesor"
        elif "ESTATUS DEL CREDITO" in cu: ren[c] = "Estatus_SaleU"
        elif "FINANCIERA" in cu: ren[c] = "Financiera"
        elif "FECHA ALTA" in cu: ren[c] = "Fecha_Alta_Credito"
        elif "ESTATUS DE LEAD" in cu: ren[c] = "Estatus_Lead"
        elif "AUTO DE INTERES" in cu: ren[c] = "Auto_Interes"
        elif cu == "FUENTE": ren[c] = "Fuente"
        elif "PLAZO" in cu: ren[c] = "Plazos_SaleU"
        elif "MONTO MENSUAL" in cu: ren[c] = "Mensualidad_SaleU"
    df = df.rename(columns=ren)
    for c in ["Nombre", "Telefono", "Asesor", "Estatus_SaleU", "Financiera",
              "Fecha_Alta_Credito", "Estatus_Lead", "Auto_Interes", "Fuente", "Plazos_SaleU", "Mensualidad_SaleU"]:
        if c not in df.columns:
            df[c] = ""
    df = df.fillna("")
    df["Tel_norm"] = df["Telefono"].map(normalizar_tel)
    df["Nombre_norm"] = df["Nombre"].map(normalizar_nombre)
    df["Fin_norm"] = df["Financiera"].map(lambda x: quitar_acentos(x).upper().strip())
    df["_fecha"] = df["Fecha_Alta_Credito"].map(convertir_fecha)
    df["Fecha_Alta_Credito"] = (df["_fecha"].dt.strftime("%d/%m/%Y %H:%M")
                                .where(df["_fecha"].notna(), df["Fecha_Alta_Credito"]))

    # Duplicados exactos (mismo teléfono y nombre): se queda el más reciente
    df = df.sort_values("_fecha", ascending=False)
    df["Veces_registrado"] = df.groupby(["Tel_norm", "Nombre_norm"])["Nombre"].transform("size")
    df = df.drop_duplicates(["Tel_norm", "Nombre_norm"]).reset_index(drop=True)
    df["id_s"] = range(len(df))
    df.attrs["rango"] = rango_reporte(info) or (
        (df["_fecha"].min().to_pydatetime(), df["_fecha"].max().to_pydatetime()) if df["_fecha"].notna().any() else None)
    return df


def rango_reporte(info):
    """'Rango: 01/08/2026 - 15/10/2026' del encabezado del reporte → (inicio, fin del día)."""
    m = re.search(r"(\d{1,2}/\d{1,2}/\d{4})\s*(?:-|–|a|al)\s*(\d{1,2}/\d{1,2}/\d{4})", info or "")
    if not m:
        return None
    try:
        ini = datetime.strptime(m.group(1), "%d/%m/%Y")
        fin = datetime.strptime(m.group(2), "%d/%m/%Y").replace(hour=23, minute=59)
    except ValueError:
        return None
    return ini, fin


def convertir_fecha(x):
    """Acepta '2026-09-23 10:23:21' (formato actual de Sale-U) y '23/09/2026 10:23'."""
    x = str(x).strip()
    if not x:
        return pd.NaT
    if re.match(r"\d{4}-\d{2}-\d{2}", x):
        return pd.to_datetime(x, errors="coerce")
    return pd.to_datetime(x, dayfirst=True, errors="coerce")


def leer_cetelem(rutas):
    partes = []
    for ruta in rutas:
        xls = pd.ExcelFile(ruta)
        for hoja in xls.sheet_names:
            d = pd.read_excel(ruta, sheet_name=hoja, dtype=str).fillna("")
            d.columns = d.columns.str.strip()
            col = {quitar_acentos(c).upper(): c for c in d.columns}
            req = {"Folio": "FOLIO CCK", "Estatus": "STATUS",
                   "Nombre": "NOMBRE COMPLETO", "Telefono": "TELEFONO MOVIL",
                   "Vendedor": "NOMBRE DEL VENDEDOR"}
            if not all(v in col for v in req.values()):
                print(f"  ⚠ Hoja '{hoja}' de {ruta} no tiene el formato esperado, se omite.")
                continue
            orden, etiqueta = detectar_mes(ruta, hoja)
            sub = pd.DataFrame({k: d[col[v]] for k, v in req.items()})
            for extra, nombre in [("VEHICULO", "Vehiculo"),
                                  ("MONTO TOTAL A FINANCIAR", "Monto_Financiar"), ("PLAZO (MESES)", "Plazo"),
                                  ("MENSUALIDAD MAXIMA", "Mensualidad"), ("CORREO ELECTRONICO", "Correo")]:
                sub[nombre] = d[col[extra]] if extra in col else ""
            sub["Mes"] = etiqueta
            sub["Fecha"] = sub["Folio"].map(lambda f: folio_date(to_folio(f), orden // 100))  # día del año en el folio
            sub["Financiera"] = "CETELEM"
            sub["_orden_mes"] = orden
            sub = sub[sub["Nombre"].str.strip() != ""]
            partes.append(sub)
    df = pd.concat(partes, ignore_index=True)
    df["Estatus"] = df["Estatus"].str.strip().str.upper()
    df["Nombre"] = df["Nombre"].str.split().str.join(" ")
    df["Tel_norm"] = df["Telefono"].map(normalizar_tel)
    df["Nombre_norm"] = df["Nombre"].map(normalizar_nombre)
    return df.sort_values(["_orden_mes", "Folio"]).reset_index(drop=True)


def desde_bitacora(sol):
    """sol: fuentes.solicitudes_de_bitacora(). Para Sale-U cuenta el resultado final de cada solicitud."""
    df = sol.rename(columns={"Est_fin": "Estatus", "Fin_fin": "Financiera"})[
        ["Folio", "Estatus", "Nombre", "Telefono", "Vendedor", "Vehiculo", "Monto_Financiar", "Mes",
         "Financiera", "Fecha", "_anio", "_mes", "Plazo", "Mensualidad", "Correo"]].copy()
    df["_siempre"] = sol["_siempre"].astype(bool).values if "_siempre" in sol else False
    df = df[df["Financiera"] != "CONTADO"].copy()          # ventas de contado: no son créditos en Sale-U
    df["Estatus"] = df["Estatus"].replace("", "SIN ESTATUS")
    df["_orden_mes"] = df["_anio"] * 100 + df["_mes"]
    df["Tel_norm"] = df["Telefono"].map(normalizar_tel)
    df["Nombre_norm"] = df["Nombre"].map(normalizar_nombre)
    return df.sort_values(["_orden_mes", "Folio"]).reset_index(drop=True).drop(columns=["_anio", "_mes"])


# =============================================================================
# 2. NORMALIZACIÓN Y SIMILITUD
# =============================================================================
def quitar_acentos(t):
    t = unicodedata.normalize("NFKD", str(t))
    return "".join(ch for ch in t if not unicodedata.combining(ch))


def normalizar_tel(t):
    """Teléfono a 10 dígitos (texto); misma limpieza en todo el sistema (base.tel_texto)."""
    from cuadres.base import tel_texto
    return tel_texto(t)


def normalizar_nombre(n):
    n = quitar_acentos(n).upper()
    if "@" in n or re.search(r"\.(COM|MX|NET)\b", n):
        return ""                        # se capturó un correo en lugar de nombre
    n = re.sub(r"[^A-ZÑ ]", " ", n)
    return " ".join(p for p in n.split() if len(p) > 1 and p not in PALABRAS_VACIAS)


def similitud_nombre(a, b):
    """Proporción de palabras del nombre más corto que aparecen en el más largo."""
    ta, tb = a.split(), b.split()
    if not ta or not tb:
        return 0.0, 0
    corto, largo = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    usados, coinc = set(), 0
    for p in corto:
        mejor, idx = 0, None
        for i, q in enumerate(largo):
            if i in usados:
                continue
            r = SequenceMatcher(None, p, q).ratio()
            if r > mejor:
                mejor, idx = r, i
        if mejor >= SIMILITUD_PALABRA:
            coinc += 1
            usados.add(idx)
    return coinc / len(corto), coinc


# =============================================================================
# 3. CONSOLIDAR CLIENTES DE LA BITÁCORA (un cliente puede aparecer varios meses)
# =============================================================================
def consolidar_clientes(cet):
    clientes = []   # cada uno: dict con tels, nombre, filas
    cet["id_c"] = -1
    for i, r in cet.iterrows():
        destino = None
        for c in clientes:
            sim, n = similitud_nombre(r["Nombre_norm"], c["nombre_norm"])
            mismo_tel = r["Tel_norm"] and r["Tel_norm"] in c["tels"]
            if (mismo_tel and sim >= UMBRAL_NOMBRE_CON_TEL) or (sim >= 0.9 and n >= 3):
                destino = c
                break
        if destino is None:
            destino = {"id": len(clientes), "tels": set(), "filas": []}
            clientes.append(destino)
        destino["nombre_norm"] = r["Nombre_norm"]
        destino["tels"].add(r["Tel_norm"])
        destino["filas"].append(i)
        cet.at[i, "id_c"] = destino["id"]

    def ultimo(sub, col):
        """Último dato no vacío de la columna (la solicitud más reciente que lo trae)."""
        if col not in sub:
            return ""
        vals = [v for v in sub[col] if v is not None and v == v and str(v).strip() not in ("", "nan", "None")]
        return vals[-1] if vals else ""

    filas = []
    for c in clientes:
        sub = cet.loc[c["filas"]]
        ult = sub.iloc[-1]
        fecha = ultimo(sub, "Fecha")
        filas.append({
            "id_c": c["id"],
            "Fecha": fecha if isinstance(fecha, (date, datetime)) else None,   # solicitud más reciente
            "Cliente_bitacora": ult["Nombre"],
            "Nombre_norm": ult["Nombre_norm"],
            "Tel_bitacora": ult["Tel_norm"] or ult["Telefono"],
            "tels": {t for t in c["tels"] if t},
            "Vendedor": ult["Vendedor"],
            "Folios": ", ".join(dict.fromkeys(sub["Folio"])),
            "Financiera": ", ".join(dict.fromkeys(sub["Financiera"])),
            "financieras": set(sub["Financiera"]),
            "Historial": " → ".join(f"{m}: {e} ({f})" for m, e, f in zip(sub["Mes"], sub["Estatus"], sub["Financiera"])),
            "Ultimo_mes": ult["Mes"],
            "Estatus_final": ult["Estatus"],
            "Vehiculo": ult["Vehiculo"],
            # datos de contacto y del crédito para la hoja «Faltan en Sale-U»
            "Financiera_ultima": ult["Financiera"],
            "Plazo": _numero(ultimo(sub, "Plazo"), entero=True),
            "Unidad": ultimo(sub, "Vehiculo"),
            "Mensualidad": _numero(ultimo(sub, "Mensualidad")),
            "Correo": str(ultimo(sub, "Correo")).strip().lower(),
        })
    return pd.DataFrame(filas)


def _numero(v, entero=False):
    s = re.sub(r"[^\d.]", "", str(v or ""))
    try:
        n = float(s) if s else None
    except ValueError:
        return None
    return (int(n) if entero else round(n, 2)) if n else None


# =============================================================================
# 4. CRUCE
# =============================================================================
SOLO_SALEU = "En Sale-U sin solicitud en la bitácora"
SIN_FINANCIERA = "SIN FINANCIERA"


def de(fuente):
    """«la bitácora» → «de la bitácora»; «el Sheet» → «del Sheet»."""
    return "del " + fuente[3:] if fuente.startswith("el ") else "de " + fuente


def cruzar(cli, su, fuente="la bitácora"):
    candidatos = []
    for _, c in cli.iterrows():
        for _, s in su.iterrows():
            sim, n = similitud_nombre(c["Nombre_norm"], s["Nombre_norm"])
            tel = bool(s["Tel_norm"]) and s["Tel_norm"] in c["tels"]
            if tel and sim >= UMBRAL_NOMBRE_CON_TEL:
                nivel = 1
            elif not tel and sim >= UMBRAL_NOMBRE_FUERTE and n >= 2:
                nivel = 2
            elif tel:
                nivel = 3
            elif sim >= UMBRAL_NOMBRE_POSIBLE and n >= 3:
                nivel = 4
            else:
                continue
            candidatos.append((nivel, -sim, c["id_c"], s["id_s"], sim, tel))

    # Mismo criterio que la columna «Registro en Sale-U» del Sheet (saleu_sheet.py): está en Sale-U si coincide
    # el teléfono (nivel 1 con nombre parecido, nivel 3 con otro nombre: se avisa para verificar) o el nombre es
    # muy parecido (nivel 2). Nombre medio parecido (4) NO cuenta: se avisa como posible para revisar.
    candidatos.sort()
    usados_s, pares, posibles = set(), {}, {}
    for nivel, _, ic, is_, sim, tel in candidatos:
        if nivel <= 3:
            if ic not in pares:
                pares[ic] = (is_, nivel, sim)
                usados_s.add(is_)
        elif ic not in posibles:
            posibles[ic] = (is_, nivel)

    tipo = {1: "Coincide (teléfono y nombre)",
            2: "Coincide por nombre – teléfono distinto",
            3: "Coincide por teléfono – nombre distinto",
            4: "Posible coincidencia por nombre"}
    tel_dueno = {t: r["Cliente_bitacora"] for _, r in cli.iterrows() for t in r["tels"]}

    filas = []
    for _, c in cli.iterrows():
        base = c.drop(labels=["tels", "Nombre_norm", "financieras"]).to_dict()
        if c["id_c"] in pares:
            is_, nivel, sim = pares[c["id_c"]]
            s = su.loc[su["id_s"] == is_].iloc[0]
            obs = []
            if nivel in (2, 4) and s["Tel_norm"] in tel_dueno and tel_dueno[s["Tel_norm"]] != c["Cliente_bitacora"]:
                obs.append(f"El teléfono de Sale-U pertenece a {tel_dueno[s['Tel_norm']]} en {fuente}")
            elif nivel in (2, 4):
                obs.append(f"Teléfono Sale-U {s['Telefono'] or '(vacío)'} ≠ {fuente} {c['Tel_bitacora']}")
            if nivel == 3:
                obs.append("Mismo teléfono pero el nombre en Sale-U es distinto: verificar y corregir nombre")
            if s["Fin_norm"] == "":
                obs.append("Financiera vacía en Sale-U")
            elif (c["financieras"] - {SIN_FINANCIERA}
                  and not any(financiera_igual(s["Fin_norm"], f) for f in c["financieras"])):
                obs.append(f"En Sale-U dice financiera: {s['Financiera']} ({fuente}: {c['Financiera']})")
            if s["Veces_registrado"] > 1:
                obs.append(f"Registrado {s['Veces_registrado']} veces en Sale-U")
            est_cuadre, obs_est = comparar_estatus(c["Estatus_final"], s["Estatus_SaleU"], fuente)
            if obs_est:
                obs.append(obs_est)
            from cuadres.saleu_fuente import describir_credito
            base.update({
                "Cómo está en Sale-U": describir_credito(
                    s["Nombre"], s["Telefono"], s["Asesor"], s["Estatus_SaleU"], s["Financiera"],
                    s["Fecha_Alta_Credito"], {1: "teléfono y nombre", 2: "nombre, otro teléfono",
                                              3: "teléfono, otro nombre"}.get(nivel, "")),
                "Resultado": tipo[nivel],
                "Similitud_nombre": round(sim, 2),
                "Cliente_SaleU": s["Nombre"], "Tel_SaleU": s["Telefono"],
                "Asesor_SaleU": s["Asesor"], "Estatus_SaleU": s["Estatus_SaleU"],
                "Financiera_SaleU": s["Financiera"], "Fuente_SaleU": s["Fuente"],
                "Fecha_alta_SaleU": s["Fecha_Alta_Credito"],
                "Cuadre_estatus": est_cuadre, "Observaciones": "; ".join(obs),
            })
        else:
            obs_no = "Dar de alta en Sale-U"
            como_su = "No se encontró por nombre ni por teléfono"
            if c["id_c"] in posibles:
                is_, nivel = posibles[c["id_c"]]
                s = su.loc[su["id_s"] == is_].iloc[0]
                obs_no += (f"; revisar posible registro: {s['Nombre']} "
                           + ("(mismo teléfono, otro nombre)" if nivel == 3 else f"(nombre parecido, tel. {s['Telefono']})"))
                como_su = (f"Posible coincidencia por nombre · Se encuentra como: «{' '.join(str(s['Nombre']).split())}» "
                           f"con teléfono: {s['Tel_norm'] or s['Telefono'] or 'sin teléfono'} "
                           "(verifica que sea la misma persona)")
            base.update({"Resultado": "NO registrado en Sale-U", "Cómo está en Sale-U": como_su, "Similitud_nombre": None,
                         "Cliente_SaleU": "", "Tel_SaleU": "", "Asesor_SaleU": "",
                         "Estatus_SaleU": "", "Financiera_SaleU": "", "Fuente_SaleU": "", "Fecha_alta_SaleU": "",
                         "Cuadre_estatus": "—", "Observaciones": obs_no})
        filas.append(base)
    cuadre = pd.DataFrame(filas)

    cubiertas = set().union(*cli["financieras"]) if len(cli) else set()
    solo = su[~su["id_s"].isin(usados_s)].copy()
    solo["Clasificacion"] = solo["Fin_norm"].map(
        lambda f: f"Financiera que no está en {fuente} (no aplica)"
        if f and not any(financiera_igual(f, x) for x in cubiertas) else SOLO_SALEU)
    solo["Observaciones"] = ""
    solo.loc[solo["Tel_norm"] == "", "Observaciones"] = "Teléfono inválido o dañado en Sale-U"
    solo.loc[solo["Nombre_norm"] == "", "Observaciones"] += " Nombre capturado como correo/ilegible"
    return cuadre, solo


def comparar_estatus(est_cet, est_su, fuente="la bitácora"):
    esperado = ESTATUS_ESPERADO.get(est_cet, None)
    su = quitar_acentos(est_su).upper().strip()
    if esperado is None:
        return "Revisar", f"En {fuente}: {est_cet.lower()}; Sale-U: {est_su or 'vacío'}"
    if su == esperado:
        return "OK", ""
    return "No cuadra", f"Estatus: {fuente} {est_cet} vs Sale-U {est_su or 'vacío'}"


# =============================================================================
# 5. EXCEL
# =============================================================================
COLORES = {"Coincide (teléfono y nombre)": "C6EFCE",
           "Coincide por nombre – teléfono distinto": "FFEB9C",
           "Coincide por teléfono – nombre distinto": "FFEB9C",
           "Posible coincidencia por nombre": "FCE4D6",
           "NO registrado en Sale-U": "FFC7CE",
           "No cuadra": "FFC7CE", "Revisar": "FFEB9C", "OK": "C6EFCE"}


def fecha_txt(v):
    return v.strftime("%d/%m/%Y") if isinstance(v, (date, datetime)) else ""


COLS_FALTANTES = ["Fecha", "Nombre", "Celular", "Financiera", "Estatus", "Plazo", "Unidad", "Mensualidad", "Correo"]
FALTA_ALTA, FALTA_CREDITO = "Dar de alta", "Registrar crédito"


COL_SALEU = "Cómo está en Sale-U"


def marcar_fuente(faltantes, fuente_df):
    """Agrega Falta («Dar de alta» / «Registrar crédito») y, en «Cómo está en Sale-U», cómo está el lead en el
    reporte Fuente (nombre, teléfono, asesor, alta, estatus y por qué coincidió)."""
    from cuadres.saleu_fuente import Buscador, describir_lead
    b = Buscador(fuente_df)
    falta, como_su, fuente_su = [], [], []
    for _, r in faltantes.iterrows():
        s, como = None, ""
        tels = r.get("tels")
        for tel in (sorted(tels) if isinstance(tels, (set, list)) and tels else [r["Tel_bitacora"]]):
            s, como = b.buscar(r["Cliente_bitacora"], tel)
            if s is not None:
                break
        falta.append(FALTA_CREDITO if s is not None else FALTA_ALTA)
        fuente_su.append(" ".join(str(s.get("Fuente") or "").split()) if s is not None else "")
        como_su.append(describir_lead(s, como) if s is not None else
                       (r[COL_SALEU] if str(r.get(COL_SALEU, "")).startswith("Posible") else
                        "No se encontró por nombre ni por teléfono"))
    return faltantes.assign(Falta=falta, Fuente_SaleU=fuente_su, **{COL_SALEU: como_su})


def tabla_faltantes(faltantes):
    """Hoja «Faltan en Sale-U»: los datos del cliente para darlo de alta."""
    extra = {c: faltantes[c].values for c in ("Falta", COL_SALEU) if c in faltantes}
    if "Fuente_SaleU" in faltantes:
        extra["Fuente"] = faltantes["Fuente_SaleU"].values          # fuente del lead en Sale-U
    if extra:
        quitar = [c for c in ("Falta", COL_SALEU, "Fuente_SaleU") if c in faltantes]
        return tabla_faltantes(faltantes.drop(columns=quitar)).assign(**extra)
    return pd.DataFrame({
        "Fecha": faltantes["Fecha"] if "Fecha" in faltantes else None,
        "Nombre": faltantes["Cliente_bitacora"], "Celular": faltantes["Tel_bitacora"],
        "Financiera": faltantes["Financiera_ultima"].replace(SIN_FINANCIERA, ""), "Estatus": faltantes["Estatus_final"], "Plazo": faltantes["Plazo"], "Unidad": faltantes["Unidad"],
        "Mensualidad": faltantes["Mensualidad"], "Correo": faltantes["Correo"]}, columns=COLS_FALTANTES)


COLS_CUADRE = ["Resultado", "Cuadre_estatus", "Cliente_bitacora", "Tel_bitacora", "Cómo está en Sale-U", "Vendedor",
               "Financiera",
               "Folios", "Historial", "Estatus_final", "Cliente_SaleU", "Tel_SaleU", "Asesor_SaleU", "Estatus_SaleU",
               "Financiera_SaleU", "Fuente", "Fecha_alta_SaleU", "Similitud_nombre", "Observaciones"]
COLS_SOLO = ["Clasificacion", "Nombre", "Telefono", "Asesor", "Estatus_SaleU", "Financiera", "Fecha_Alta_Credito",
             "Fuente", "Observaciones"]


def _faltantes_con_vendedor(faltantes):
    vista = tabla_faltantes(faltantes)
    vista.insert(2, "Vendedor", faltantes["Vendedor"].replace("(sin vendedor)", "").values)
    return vista


def _orden_faltantes(faltantes):
    return faltantes.sort_values(["Falta", "Vendedor", "Cliente_bitacora"] if "Falta" in faltantes
                                 else ["Vendedor", "Cliente_bitacora"])


def exportar_excel(ruta, resumen, cuadre, detalle, faltantes, revisar, solo, sub="", titulo=""):
    con_fuente = "Falta" in faltantes
    fu = lambda df: df.assign(Fuente=df["Fuente_SaleU"]) if "Fuente_SaleU" in df else df   # fuente del lead en Sale-U
    cuadre, revisar = fu(cuadre), fu(revisar)
    sal.excel(ruta, titulo or f"Cuadre Sale-U – {sal.PDV}", sub, [
        ("Resumen", resumen, ["Concepto", "Valor"]),
        ("Faltan en Sale-U", _faltantes_con_vendedor(_orden_faltantes(faltantes)),
         (["Falta"] if con_fuente else []) + COLS_FALTANTES[:2] + ["Vendedor"] + COLS_FALTANTES[2:]
         + [COL_SALEU, "Fuente"]),
        ("A revisar", revisar, COLS_CUADRE),
        ("Solo en Sale-U", solo, COLS_SOLO),
        ("Detalle por mes", detalle, list(detalle.columns)),
        ("Cuadre por cliente", cuadre, COLS_CUADRE),
    ], COLORES)


# =============================================================================
# 6. PDF
# =============================================================================
def exportar_pdf(ruta, resumen, faltantes, no_cuadra, revisar_cruce, solo_cet, fuente="la bitácora", sub="",
                 cifras=None, titulo=""):
    pdf = sal.PDF(ruta, titulo or f"Cuadre Sale-U – {sal.PDV}", sub)
    fu = lambda df: df.assign(Fuente=df["Fuente_SaleU"]) if "Fuente_SaleU" in df else df   # fuente del lead en Sale-U
    no_cuadra, revisar_cruce = fu(no_cuadra), fu(revisar_cruce)
    pdf.cifras(cifras or [])
    cols = ["Fecha", "Nombre", "Vendedor", "Celular", "Financiera", "Estatus", "Mensualidad", "Correo", COL_SALEU]
    anchos = [1.9, 3.4, 2.8, 2.0, 1.9, 2.2, 1.9, 3.6, 5.8]
    por_vendedor = lambda df: ([("Por vendedor", df.groupby("Vendedor").size().sort_values(ascending=False))]
                               if not df.empty else None)
    if "Falta" in faltantes:
        alta, cred = faltantes[faltantes["Falta"] == FALTA_ALTA], faltantes[faltantes["Falta"] == FALTA_CREDITO]
        pdf.seccion("Dar de alta en Sale-U", _faltantes_con_vendedor(_orden_faltantes(alta)), cols, anchos, "#C00000",
                    "No están ni en el reporte Fuente ni en el de Crédito (ni por teléfono ni por nombre): darlos de "
                    "alta en Sale-U con el asesor correspondiente y registrar su crédito.", por_vendedor(alta))
        pdf.seccion("Registrar el crédito en Sale-U (ya son lead)", _faltantes_con_vendedor(_orden_faltantes(cred)),
                    ["Fecha", "Nombre", "Vendedor", "Celular", "Financiera", "Estatus", "Mensualidad", COL_SALEU,
                     "Fuente"],
                    [1.8, 3.4, 2.7, 2.0, 1.8, 2.1, 1.8, 6.8, 3.1], "#BF8F00",
                    "Salen en el reporte Fuente pero no en el de Crédito: busca el lead como aparece en «Cómo está "
                    "en Sale-U» (nombre y teléfono con que se capturó) y registra ahí el crédito, sin darlo de alta "
                    "otra vez.", por_vendedor(cred))
    else:
        pdf.seccion("Clientes que faltan en Sale-U", _faltantes_con_vendedor(_orden_faltantes(faltantes)), cols,
                    anchos, "#C00000", "Darlos de alta en Sale-U con el asesor correspondiente.", por_vendedor(faltantes))
    pdf.seccion("Estatus distinto en Sale-U", no_cuadra,
                ["Cliente_bitacora", "Vendedor", "Historial", COL_SALEU, "Fuente", "Observaciones"],
                [3.9, 3.2, 4.4, 6.0, 2.9, 5.1], "#C55A11",
                f"El cliente sí está en Sale-U (búscalo como dice «Cómo está en Sale-U»), pero su estatus no "
                f"corresponde con {fuente}.")
    pdf.seccion("Coincidencias a verificar", revisar_cruce,
                ["Resultado", "Cliente_bitacora", "Tel_bitacora", COL_SALEU, "Fuente", "Observaciones"],
                [3.7, 4.0, 2.2, 6.9, 2.9, 5.8], "#BF8F00",
                "Se emparejaron por nombre con teléfono distinto, o por teléfono con nombre distinto. "
                "Confirma que sea la misma persona y corrige el dato en Sale-U.")
    pdf.seccion(f"En Sale-U sin solicitud en {fuente}", solo_cet,
                ["Nombre", "Telefono", "Asesor", "Estatus_SaleU", "Financiera", "Fecha_Alta_Credito", "Observaciones"],
                [5, 2.4, 4, 2.4, 2.6, 3, 6], "#7F7F7F",
                f"Están en Sale-U con una financiera que sí trae {fuente}, pero no se encontró su solicitud.")
    pdf.guardar(resumen, f"Cada cliente {de(fuente)} se busca en Sale-U por teléfono o por nombre: basta con que "
                         "coincida uno de los dos (si el otro dato es distinto se avisa para verificar). «Cómo está en "
                         "Sale-U» dice con qué nombre, teléfono, asesor y estatus está capturado. Un cliente que "
                         "aparece en varios meses se cuenta una sola vez con su último estatus.")


# =============================================================================
# 7. PROGRAMA PRINCIPAL
# =============================================================================
def procesar(ruta_su, rutas_cet=None, carpeta=".", bitacora=None, fuente="la bitácora", ruta_fuente=None):
    """Hace todo el cuadre y devuelve un diccionario con tablas y rutas de salida.
    bitacora: DataFrame de fuentes.solicitudes_de_bitacora() (todas las financieras). Si no se da,
    se leen los Excel mensuales de Cetelem de `rutas_cet`.
    fuente: cómo se nombra la base en el Resumen y el PDF («la bitácora», «el Sheet»).
    ruta_fuente: reporte Fuente de Sale-U (opcional) o su DataFrame de saleu_fuente.leer_fuente(); con él, los
    faltantes se separan en «Dar de alta» y «Registrar crédito»."""
    su = leer_saleu(ruta_su)
    rango = su.attrs.get("rango")
    if bitacora is not None:
        cet = desde_bitacora(bitacora)
        if rango:   # solo las solicitudes del periodo que cubre el reporte de Sale-U
            ini, fin = rango
            en_periodo = cet["Fecha"].map(lambda d: bool(d) and ini.date() <= d <= fin.date())
            cet = cet[en_periodo | cet["_siempre"]].reset_index(drop=True)
        cet = cet.drop(columns=["_siempre"])
    else:
        cet = leer_cetelem(rutas_cet)
    if cet.empty:
        raise ValueError(f"{fuente[0].upper()}{fuente[1:]} no tiene registros en el periodo del reporte de Sale-U.")
    meses = [m for m in dict.fromkeys(cet.sort_values("_orden_mes")["Mes"]) if m != "SIN FECHA"]
    financieras = sorted(set(cet["Financiera"]) - {"", SIN_FINANCIERA}, key=lambda f: (f != "CETELEM", f))
    cli = consolidar_clientes(cet)
    cuadre, solo = cruzar(cli, su, fuente)

    detalle = cet.merge(cuadre[["id_c", "Resultado", "Cliente_SaleU", "Estatus_SaleU", "Cuadre_estatus"]],
                        on="id_c", how="left")
    detalle = detalle[["Mes", "Folio", "Estatus", "Financiera", "Nombre", "Telefono", "Vendedor",
                       "Vehiculo", "Monto_Financiar", "Resultado", "Cliente_SaleU",
                       "Estatus_SaleU", "Cuadre_estatus"]]

    encontrados = cuadre["Resultado"] != "NO registrado en Sale-U"
    faltantes = cuadre[~encontrados]
    fu = None
    if ruta_fuente is not None:
        from cuadres.saleu_fuente import leer_fuente
        fu = ruta_fuente if isinstance(ruta_fuente, pd.DataFrame) else leer_fuente(ruta_fuente)
        tels = cli.set_index("id_c")["tels"]
        faltantes = marcar_fuente(faltantes.assign(tels=faltantes["id_c"].map(tels)), fu).drop(columns=["tels"])
        for c in ("Fuente_SaleU", COL_SALEU):          # también en «Cuadre por cliente»
            cuadre.loc[faltantes.index, c] = faltantes[c]
    n_alta = int((faltantes["Falta"] == FALTA_ALTA).sum()) if fu is not None else None
    n_cred = len(faltantes) - n_alta if fu is not None else None
    revisar_cruce = cuadre[cuadre["Resultado"].isin([
        "Coincide por nombre – teléfono distinto", "Coincide por teléfono – nombre distinto",
        "Posible coincidencia por nombre"])]
    no_cuadra = cuadre[encontrados & (cuadre["Cuadre_estatus"] == "No cuadra")]
    revisar = cuadre[encontrados & ((cuadre["Cuadre_estatus"] != "OK") | cuadre.index.isin(revisar_cruce.index)
                                    | (cuadre["Observaciones"] != ""))]
    solo_cet = solo[solo["Clasificacion"] == SOLO_SALEU]

    resumen = pd.DataFrame([
        ("Periodo del reporte de Sale-U", f"{rango[0]:%d/%m/%Y} – {rango[1]:%d/%m/%Y}" if rango else "—"),
        (f"Registros en {fuente} (todas las financieras)", len(cet)),
        (f"Clientes únicos en {fuente}", len(cli)),
        ("   · por financiera", " · ".join(f"{k} {v}" for k, v in
                                          cli["Financiera"].str.split(", ").explode().value_counts().items())),
        ("Registros en Sale-U (sin duplicados)", len(su)),
        (f"Clientes {de(fuente)} encontrados en Sale-U", int(encontrados.sum())),
        ("   · por teléfono y nombre", int((cuadre["Resultado"] == "Coincide (teléfono y nombre)").sum())),
        ("   · por nombre, teléfono distinto", int((cuadre["Resultado"] == "Coincide por nombre – teléfono distinto").sum())),
        ("   · por teléfono, nombre distinto (verificar)",
         int((cuadre["Resultado"] == "Coincide por teléfono – nombre distinto").sum())),
        (f"Clientes {de(fuente)} NO registrados en Sale-U", len(faltantes)),
        *([("   · no están en Sale-U (dar de alta)", n_alta),
           ("   · ya son lead en el reporte Fuente (registrar crédito)", n_cred)] if fu is not None else []),
        ("   · con posible registro a revisar (mismo teléfono u otro nombre parecido)",
         int(faltantes["Observaciones"].str.contains("revisar posible").sum())),
        ("Porcentaje de cuadre", f"{encontrados.mean():.1%}" if len(cuadre) else "—"),
        ("Estatus que no cuadra", len(no_cuadra)),
        ("En contrapropuesta o análisis (revisar estatus)", int((cuadre["Cuadre_estatus"] == "Revisar").sum())),
        (f"En Sale-U sin registro en {fuente}", len(solo_cet)),
        (f"Sale-U de financieras que no están en {fuente} (no aplica)", len(solo) - len(solo_cet)),
    ], columns=["Concepto", "Valor"])

    periodo = sal.periodo_de_fechas(*rango) if rango else sal.periodo_de_meses(meses)
    base_nombre = sal.nombre_archivo("SaleU", periodo)
    ruta_xlsx = os.path.join(carpeta, base_nombre + ".xlsx")
    ruta_pdf = os.path.join(carpeta, base_nombre + ".pdf")
    titulo = f"Cuadre Sale-U vs {fuente.split(' ', 1)[1]} – {sal.PDV}"
    sub = ((f"Periodo del reporte: {rango[0]:%d/%m/%Y} – {rango[1]:%d/%m/%Y} · " if rango else "")
           + f"Meses: {', '.join(meses)} · Financieras: {', '.join(financieras)} · "
           f"Generado: {datetime.now():%d/%m/%Y %H:%M}")
    cifras = [
        (f"Clientes en {fuente.split(' ', 1)[1]}", len(cli), "clientes únicos", "neutro"),
        ("En Sale-U", int(encontrados.sum()), f"{encontrados.mean():.0%} de cuadre" if len(cuadre) else "", "bien"),
        *([("Dar de alta", n_alta, "no están en Sale-U", "mal" if n_alta else "bien"),
           ("Registrar crédito", n_cred, "ya son lead (Fuente)", "aviso" if n_cred else "bien")] if fu is not None else
          [("Faltan en Sale-U", len(faltantes), "por dar de alta", "mal" if len(faltantes) else "bien")]),
        ("Estatus distinto", len(no_cuadra), "corregir en Sale-U", "aviso" if len(no_cuadra) else "bien"),
        ("A verificar", len(revisar_cruce), "nombre o teléfono distinto", "aviso" if len(revisar_cruce) else "bien"),
    ]
    exportar_excel(ruta_xlsx, resumen, cuadre, detalle, faltantes, revisar, solo, sub, titulo)
    exportar_pdf(ruta_pdf, resumen, faltantes.copy(), no_cuadra.copy(), revisar_cruce.copy(), solo_cet, fuente,
                 sub, cifras, titulo)
    return dict(resumen=resumen, cuadre=cuadre, detalle=detalle, faltantes=faltantes,
                no_cuadra=no_cuadra, revisar_cruce=revisar_cruce, revisar=revisar, solo=solo,
                solo_cet=solo_cet, meses=meses, financieras=financieras, rango=rango, fuente=fuente,
                xlsx=ruta_xlsx, pdf=ruta_pdf)


def main():
    ruta_su, rutas_cet = obtener_archivos()
    print(f"Sale-U: {ruta_su}\nCetelem: {', '.join(rutas_cet)}")
    r = procesar(ruta_su, rutas_cet)
    resumen, ruta_xlsx, ruta_pdf = r["resumen"], r["xlsx"], r["pdf"]

    print("\n" + resumen.to_string(index=False))
    print(f"\n✔ Archivos generados: {ruta_xlsx} y {ruta_pdf}")
    try:
        from google.colab import files
        files.download(ruta_xlsx)
        files.download(ruta_pdf)
    except ImportError:
        pass


if __name__ == "__main__":
    main()
