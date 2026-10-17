# -*- coding: utf-8 -*-
"""
cuadre de seguimiento de rechazos
---------------------------------
Cuadre: rechazos de la bitácora (todas las financieras: capturas de Cetelem + Bitácora de F&I,
consolidadas en la pestaña "Bitacora" de la base) vs. la base de seguimiento,
tomando SOLO las pestañas cuyo nombre contiene "COLIMA".
Sin base, también acepta los Excel mensuales de Cetelem (JULIO_2026_FINAL.xlsx, …).

Orden de cruce (por cliente, de lo más confiable a lo menos):
  0) Mismo folio (y coincide teléfono o nombre)
  1) Mismo teléfono + nombre parecido
  2) Nombre muy parecido, teléfono distinto
  3) Mismo teléfono, nombre distinto
  4) Nombre medianamente parecido (posible)

Salidas: Cuadre_Seguimiento_<periodo>_<fecha>.xlsx y .pdf (formato común en cuadres/salidas.py)
Uso en Colab: ejecuta y sube el Excel de seguimiento + los Excel mensuales de Cetelem.
"""

import os, re, glob, unicodedata, subprocess, sys, warnings, calendar
from datetime import datetime, timedelta
from difflib import SequenceMatcher
import pandas as pd

warnings.filterwarnings("ignore")
from cuadres.fuentes import financiera_igual
from cuadres import salidas as sal

# =============================================================================
# CONFIGURACIÓN
# =============================================================================
PDV = "MG COLIMA"
FILTRO_HOJAS = "COLIMA"            # solo pestañas cuyo nombre contenga esto
PALABRA_ARCHIVO_SEGUIMIENTO = "SEGUIMIENTO"

UMBRAL_NOMBRE_CON_TEL = 0.50
UMBRAL_NOMBRE_FUERTE = 0.80
UMBRAL_NOMBRE_POSIBLE = 0.60
SIMILITUD_PALABRA = 0.80

# La base de seguimiento solo registra rechazos: de la bitácora solo se cuadran estos estatus.
ESTATUS_A_CUADRAR = {"RECHAZADO"}

# En qué pestaña debería estar cada estatus (texto que contiene el nombre de la hoja).
# None = debe estar capturado, pero no se juzga en qué pestaña.
HOJA_ESPERADA = {
    "RECHAZADO": "RECHAZO",
    "FINANCIADO": "FINALIZADO",
    "FINANCIADOS": "FINALIZADO",
    "FORMALIZADO": "FINALIZADO",
    "APROBADO": None,
    "CONTRAPROPUESTA": None,
}

MESES = {"ENERO": 1, "FEBRERO": 2, "MARZO": 3, "ABRIL": 4, "MAYO": 5, "JUNIO": 6,
         "JULIO": 7, "AGOSTO": 8, "SEPTIEMBRE": 9, "SETIEMBRE": 9, "OCTUBRE": 10,
         "NOVIEMBRE": 11, "DICIEMBRE": 12}
PALABRAS_VACIAS = {"DE", "DEL", "LA", "LAS", "LOS", "Y", "SA", "CV", "SAPI", "RL", "MA"}


# =============================================================================
# 1. UTILIDADES
# =============================================================================
def quitar_acentos(t):
    t = unicodedata.normalize("NFKD", str(t))
    return "".join(ch for ch in t if not unicodedata.combining(ch))


def normalizar_tel(t):
    t = str(t).strip()
    if t in ("", "nan", "None") or re.search(r"[eE]\+", t):
        return ""
    t = re.sub(r"\.0$", "", t)
    dig = re.sub(r"\D", "", t)
    if len(dig) > 10 and dig.startswith("52"):
        dig = dig[2:]
    return dig[-10:] if len(dig) >= 10 else ""


def normalizar_folio(f):
    f = re.sub(r"\.0$", "", str(f).strip())
    return f if f.isdigit() else ""


def normalizar_nombre(n):
    n = quitar_acentos(n).upper()
    if "@" in n:
        return ""
    n = re.sub(r"[^A-Z ]", " ", n)
    return " ".join(p for p in n.split() if len(p) > 1 and p not in PALABRAS_VACIAS)


def similitud_nombre(a, b):
    """Proporción de palabras del nombre más corto que aparecen en el más largo
    (no importa el orden: 'CARPIO GUZMAN VICTOR' = 'VICTOR CARPIO GUZMAN')."""
    ta, tb = a.split(), b.split()
    if not ta or not tb:
        return 0.0, 0
    corto, largo = (ta, tb) if len(ta) <= len(tb) else (tb, ta)
    usados, coinc = set(), 0
    for p in corto:
        mejor, idx = 0, None
        for i, q in enumerate(largo):
            if i not in usados:
                r = SequenceMatcher(None, p, q).ratio()
                if r > mejor:
                    mejor, idx = r, i
        if mejor >= SIMILITUD_PALABRA:
            coinc += 1
            usados.add(idx)
    return coinc / len(corto), coinc


def fecha_desde_folio(folio, anio):
    """En los folios CCK los primeros dígitos son el día del año (18265507 -> día 182)."""
    if len(folio) < 6:
        return None
    try:
        dia = int(folio[:-5])
        if 1 <= dia <= 366:
            return datetime(anio, 1, 1) + timedelta(days=dia - 1)
    except ValueError:
        pass
    return None


def buscar_col(columnas, *opciones):
    norm = {quitar_acentos(c).upper().strip(): c for c in columnas}
    for op in opciones:
        for k, c in norm.items():
            if k == op or k.startswith(op):
                return c
    return None


# =============================================================================
# 2. CARGA
# =============================================================================
def obtener_archivos():
    try:
        from google.colab import files
        print("Sube el Excel de seguimiento y los Excel mensuales de Cetelem:")
        rutas = list(files.upload().keys())
    except ImportError:
        rutas = glob.glob("*.xlsx")
    rutas = [r for r in rutas if r.lower().endswith((".xlsx", ".xls")) and not r.startswith("Cuadre_")]
    seg = [r for r in rutas if PALABRA_ARCHIVO_SEGUIMIENTO in quitar_acentos(r).upper()]
    cet = [r for r in rutas if r not in seg]
    if not seg:
        raise FileNotFoundError("No encontré el Excel de seguimiento (nombre con 'Seguimiento').")
    if not cet:
        raise FileNotFoundError("No encontré capturas mensuales de Cetelem.")
    return seg[0], cet


def detectar_mes(archivo, hoja):
    texto = quitar_acentos(f"{hoja} {archivo}").upper()
    mes = next((MESES[m] for m in MESES if m in texto), 0)
    anio = re.search(r"20\d{2}", texto)
    anio = int(anio.group()) if anio else datetime.now().year
    etiqueta = next((m for m in MESES if MESES[m] == mes and m != "SETIEMBRE"), "SIN MES")
    return anio, mes, f"{etiqueta[:3]} {anio}"


def leer_cetelem(rutas):
    partes = []
    for ruta in rutas:
        xls = pd.ExcelFile(ruta)
        for hoja in xls.sheet_names:
            d = pd.read_excel(ruta, sheet_name=hoja, dtype=str).fillna("")
            d.columns = d.columns.str.strip()
            c = {k: buscar_col(d.columns, v) for k, v in [
                ("Folio", "FOLIO"), ("Estatus", "STATUS"), ("Nombre", "NOMBRE COMPLETO"),
                ("Telefono", "TELEFONO MOVIL"), ("Vendedor", "NOMBRE DEL VENDEDOR")]}
            if None in c.values():
                print(f"  ⚠ Hoja '{hoja}' de {ruta} sin formato de captura Cetelem, se omite.")
                continue
            anio, mes, etiqueta = detectar_mes(ruta, hoja)
            sub = pd.DataFrame({k: d[v] for k, v in c.items()})
            sub["Mes"], sub["_anio"], sub["_mes"] = etiqueta, anio, mes
            sub["Financiera"] = "CETELEM"
            partes.append(sub[sub["Nombre"].str.strip() != ""])
    df = pd.concat(partes, ignore_index=True)
    df["Estatus"] = df["Estatus"].str.strip().str.upper()
    df["Nombre"] = df["Nombre"].str.split().str.join(" ")
    df["Folio_norm"] = df["Folio"].map(normalizar_folio)
    df["Tel_norm"] = df["Telefono"].map(normalizar_tel)
    df["Nombre_norm"] = df["Nombre"].map(normalizar_nombre)
    return df.sort_values(["_anio", "_mes", "Folio"]).reset_index(drop=True)


def desde_bitacora(sol):
    """sol: fuentes.solicitudes_de_bitacora(). Estatus y financiera de cada solicitud."""
    df = sol.rename(columns={"Est_sol": "Estatus", "Fin_sol": "Financiera"})[
        ["Folio", "Estatus", "Nombre", "Telefono", "Vendedor", "Mes", "_anio", "_mes", "Financiera"]].copy()
    df = df[df["Financiera"] != "CONTADO"].copy()          # ventas de contado: no hay crédito que rechazar
    df["Folio_norm"] = df["Folio"].map(normalizar_folio)
    df["Tel_norm"] = df["Telefono"].map(normalizar_tel)
    df["Nombre_norm"] = df["Nombre"].map(normalizar_nombre)
    return df.sort_values(["_anio", "_mes", "Folio"]).reset_index(drop=True)


def leer_seguimiento(ruta, anio_ref):
    xls = pd.ExcelFile(ruta)
    hojas = [h for h in xls.sheet_names if FILTRO_HOJAS in quitar_acentos(h).upper()]
    print(f"Pestañas de seguimiento usadas: {', '.join(hojas)}")
    partes = []
    for h in hojas:
        d = pd.read_excel(ruta, sheet_name=h, dtype=object)
        if d.empty:
            continue
        d.columns = [str(c).strip() for c in d.columns]
        c_nom = buscar_col(d.columns, "NOMBRE DE CLIENTE", "NOMBRE")
        c_tel = buscar_col(d.columns, "CELULAR", "TELEFONO")
        if not c_nom:
            continue
        c_fol = buscar_col(d.columns, "FOLIO")
        c_fin = buscar_col(d.columns, "FINANCIERA")
        c_ase = buscar_col(d.columns, "ASESOR DE VENTAS", "ASESOR ASIGNADO", "ASESOR")
        c_est = buscar_col(d.columns, "STATUS", "ESTATUS")
        c_fec = buscar_col(d.columns, "FECHA INGRESO")
        c_su = buscar_col(d.columns, "CLIENTE CON REGISTRO EN SALES")
        val = lambda col: d[col] if col else pd.Series([""] * len(d))
        sub = pd.DataFrame({
            "Hoja": h, "Fila_excel": d.index + 2,
            "Nombre": val(c_nom).astype(str).str.strip(), "Telefono": val(c_tel).astype(str),
            "Folio": val(c_fol).astype(str), "Financiera": val(c_fin).astype(str),
            "Asesor": val(c_ase).astype(str), "Estatus_seg": val(c_est).astype(str),
            "Fecha_ingreso": pd.to_datetime(val(c_fec), errors="coerce"),
            "Registro_SalesU": val(c_su).astype(str),
        })
        partes.append(sub[~sub["Nombre"].isin(["", "nan", "None"])])
    df = pd.concat(partes, ignore_index=True)
    for col in ["Telefono", "Folio", "Financiera", "Asesor", "Estatus_seg", "Registro_SalesU"]:
        df[col] = df[col].replace({"nan": "", "None": "", "NaT": ""}).str.strip()
    df["Folio_norm"] = df["Folio"].map(normalizar_folio)
    df["Tel_norm"] = df["Telefono"].map(normalizar_tel)
    df["Nombre_norm"] = df["Nombre"].map(normalizar_nombre)
    df["Fin_norm"] = df["Financiera"].map(lambda x: quitar_acentos(x).upper())

    # Fechas imposibles (p. ej. capturaron la fecha de nacimiento): se estiman con el folio
    df["Obs_fecha"] = ""
    malas = df["Fecha_ingreso"].isna() | (df["Fecha_ingreso"].dt.year < anio_ref - 1)
    for i in df.index[malas]:
        est = fecha_desde_folio(df.at[i, "Folio_norm"], anio_ref)
        if pd.notna(df.at[i, "Fecha_ingreso"]):
            df.at[i, "Obs_fecha"] = (f"Fecha de ingreso inválida ({df.at[i, 'Fecha_ingreso']:%d/%m/%Y})"
                                     + (", se estimó con el folio" if est else ""))
        if est:
            df.at[i, "Fecha_ingreso"] = est

    # Duplicados: misma hoja, folio, teléfono y nombre (sin importar orden/acentos)
    df["_llave"] = (df["Hoja"] + "|" + df["Folio_norm"] + "|" + df["Tel_norm"] + "|"
                    + df["Nombre_norm"].map(lambda n: " ".join(sorted(n.split()))) + "|" + df["Fin_norm"])
    df["Veces_capturado"] = df.groupby("_llave")["Nombre"].transform("size")
    df = df.drop_duplicates("_llave").reset_index(drop=True)
    df["id_s"] = range(len(df))
    return df


# =============================================================================
# 3. CLIENTES ÚNICOS DE LA BITÁCORA
# =============================================================================
def consolidar_clientes(cet):
    grupos = []
    cet["id_c"] = -1
    for i, r in cet.iterrows():
        destino = None
        for g in grupos:
            sim, n = similitud_nombre(r["Nombre_norm"], g["nombre"])
            if ((r["Tel_norm"] and r["Tel_norm"] in g["tels"] and sim >= UMBRAL_NOMBRE_CON_TEL)
                    or (sim >= 0.9 and n >= 3)):
                destino = g
                break
        if destino is None:
            destino = {"id": len(grupos), "tels": set(), "filas": []}
            grupos.append(destino)
        destino["nombre"] = r["Nombre_norm"]
        destino["tels"].add(r["Tel_norm"])
        destino["filas"].append(i)
        cet.at[i, "id_c"] = destino["id"]
    filas = []
    for g in grupos:
        sub = cet.loc[g["filas"]]
        ult = sub.iloc[-1]
        filas.append({
            "id_c": g["id"], "Cliente_bitacora": ult["Nombre"], "Nombre_norm": ult["Nombre_norm"],
            "Tel_bitacora": ult["Tel_norm"] or ult["Telefono"], "tels": {t for t in g["tels"] if t},
            "folios": set(sub["Folio_norm"]) - {""}, "Vendedor": ult["Vendedor"],
            "Folios": ", ".join(dict.fromkeys(sub["Folio"])),
            "Financiera": ", ".join(dict.fromkeys(sub["Financiera"])),
            "financieras": set(sub["Financiera"]),
            "Historial": " → ".join(f"{m}: {e} ({f})" for m, e, f in zip(sub["Mes"], sub["Estatus"], sub["Financiera"])),
            "Estatus_final": ult["Estatus"],
        })
    return pd.DataFrame(filas)


# =============================================================================
# 4. CRUCE
# =============================================================================
TIPOS = {0: "Coincide por folio", 1: "Coincide (teléfono y nombre)",
         2: "Coincide por nombre – teléfono distinto", 3: "Coincide por teléfono – nombre distinto",
         4: "Posible coincidencia por nombre"}


def cruzar(cli, seg):
    # Solo filas de la misma financiera (o sin financiera, como la pestaña de finalizados) son
    # candidatas, salvo que el folio coincida.
    candidatos = []
    for _, c in cli.iterrows():
        for _, s in seg.iterrows():
            es_cet = s["Fin_norm"] == "" or any(financiera_igual(s["Fin_norm"], f) for f in c["financieras"])
            folio = bool(s["Folio_norm"]) and s["Folio_norm"] in c["folios"]
            sim, n = similitud_nombre(c["Nombre_norm"], s["Nombre_norm"])
            tel = bool(s["Tel_norm"]) and s["Tel_norm"] in c["tels"]
            if folio and (tel or sim >= UMBRAL_NOMBRE_CON_TEL):
                nivel = 0
            elif not es_cet:
                continue
            elif tel and sim >= UMBRAL_NOMBRE_CON_TEL:
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
    candidatos.sort()

    usados_c, usados_s, principal = set(), set(), {}
    for nivel, _, ic, is_, sim, tel in candidatos:
        if ic in usados_c or is_ in usados_s:
            continue
        usados_c.add(ic); usados_s.add(is_)
        principal[ic] = (is_, nivel, sim)

    # Otras capturas de la misma persona (otra solicitud / otra pestaña)
    extras = {}
    for nivel, _, ic, is_, sim, tel in candidatos:
        if is_ not in usados_s and ic in principal and nivel <= 1:
            extras.setdefault(ic, []).append(is_)
            usados_s.add(is_)

    tel_dueno = {t: r["Cliente_bitacora"] for _, r in cli.iterrows() for t in r["tels"]}
    S = seg.set_index("id_s")
    filas = []
    for _, c in cli.iterrows():
        base = {k: c[k] for k in ["Cliente_bitacora", "Tel_bitacora", "Vendedor", "Financiera", "Folios",
                                  "Historial", "Estatus_final"]}
        esperada = HOJA_ESPERADA.get(c["Estatus_final"])
        if c["id_c"] not in principal:
            base.update({"Resultado": "NO capturado en seguimiento", "Cuadre_pestana": "—",
                         "Pestanas_encontrado": "", "Cliente_seguimiento": "", "Tel_seguimiento": "",
                         "Folio_seguimiento": "", "Asesor_seguimiento": "", "Fecha_ingreso_seg": "",
                         "Similitud_nombre": None,
                         "Observaciones": "Capturar en la pestaña de "
                                          + ("rechazos" if esperada == "RECHAZO" else
                                             "créditos finalizados" if esperada == "FINALIZADO" else
                                             "seguimiento (confirmar pestaña)")})
            filas.append(base)
            continue
        is_, nivel, sim = principal[c["id_c"]]
        s = S.loc[is_]
        todas = [is_] + extras.get(c["id_c"], [])
        hojas = list(dict.fromkeys(S.loc[todas, "Hoja"]))
        obs = []
        if sim < UMBRAL_NOMBRE_CON_TEL:
            obs.append(f"Nombre en seguimiento distinto: {s['Nombre']}")
        if s["Folio_norm"] and s["Folio_norm"] not in c["folios"]:
            obs.append(f"Folio en seguimiento {s['Folio_norm']} no coincide con la bitácora")
        elif not s["Folio_norm"] and c["folios"]:
            obs.append(f"Falta el folio en el seguimiento (bitácora: {', '.join(sorted(c['folios']))})")
        if nivel == 0 and s["Tel_norm"] and s["Tel_norm"] not in c["tels"]:
            obs.append(f"Teléfono en seguimiento {s['Telefono']} ≠ bitácora {c['Tel_bitacora']}")
        if nivel in (2, 4):
            if s["Tel_norm"] in tel_dueno and tel_dueno[s["Tel_norm"]] != c["Cliente_bitacora"]:
                obs.append(f"El teléfono capturado pertenece a {tel_dueno[s['Tel_norm']]}")
            else:
                obs.append(f"Teléfono en seguimiento {s['Telefono'] or '(vacío)'} ≠ bitácora {c['Tel_bitacora']}")
        if s["Fin_norm"] and not any(financiera_igual(s["Fin_norm"], f) for f in c["financieras"]):
            obs.append(f"Capturado con financiera {s['Financiera']} (en la bitácora: {c['Financiera']})")
        dup = int(S.loc[todas, "Veces_capturado"].sum())
        if dup > 1:
            obs.append(f"Capturado {dup} veces")
        obs += [o for o in S.loc[todas, "Obs_fecha"] if o]

        if esperada is None:
            cuadre = "Revisar"
            obs.append(f"En la bitácora: {c['Estatus_final'].lower()}, confirmar pestaña")
        elif any(esperada in quitar_acentos(h).upper() for h in hojas):
            cuadre = "OK"
        else:
            cuadre = "No cuadra"
            obs.append(f"En la bitácora {c['Estatus_final']} pero está en '{', '.join(hojas)}'")

        base.update({
            "Resultado": TIPOS[nivel], "Cuadre_pestana": cuadre, "Pestanas_encontrado": ", ".join(hojas),
            "Cliente_seguimiento": s["Nombre"], "Tel_seguimiento": s["Telefono"],
            "Folio_seguimiento": s["Folio_norm"], "Asesor_seguimiento": s["Asesor"],
            "Fecha_ingreso_seg": s["Fecha_ingreso"].strftime("%d/%m/%Y") if pd.notna(s["Fecha_ingreso"]) else "",
            "Similitud_nombre": round(sim, 2), "Observaciones": "; ".join(obs)})
        filas.append(base)
    return pd.DataFrame(filas), seg[~seg["id_s"].isin(usados_s)].copy()


SOLO_PERIODO = "En seguimiento pero NO en la bitácora del periodo"


def clasificar_sobrantes(solo, ini, fin, cubiertas):
    """cubiertas: financieras que trae la bitácora (las demás no se pueden cuadrar)."""
    def clas(r):
        if r["Fin_norm"] and not any(financiera_igual(r["Fin_norm"], f) for f in cubiertas):
            return "Financiera que no está en la bitácora (no aplica)"
        if pd.isna(r["Fecha_ingreso"]):
            return "Sin fecha (no se puede ubicar en el periodo)"
        if ini <= r["Fecha_ingreso"] <= fin:
            return SOLO_PERIODO
        return "Fuera del periodo de la bitácora"
    solo["Clasificacion"] = solo.apply(clas, axis=1)
    solo["Fecha_ingreso"] = solo["Fecha_ingreso"].dt.strftime("%d/%m/%Y").fillna("")
    solo["Observaciones"] = solo["Obs_fecha"]
    solo.loc[solo["Veces_capturado"] > 1, "Observaciones"] += " Capturado varias veces."
    return solo


# =============================================================================
# 5. EXCEL Y PDF
# =============================================================================
COLORES = {"Coincide por folio": "C6EFCE", "Coincide (teléfono y nombre)": "C6EFCE",
           "Coincide por nombre – teléfono distinto": "FFEB9C",
           "Coincide por teléfono – nombre distinto": "FFEB9C",
           "Posible coincidencia por nombre": "FCE4D6", "NO capturado en seguimiento": "FFC7CE",
           "No cuadra": "FFC7CE", "Revisar": "FFEB9C", "OK": "C6EFCE",
           SOLO_PERIODO: "FFC7CE"}

COLS_CUADRE = ["Resultado", "Cuadre_pestana", "Cliente_bitacora", "Tel_bitacora", "Vendedor", "Financiera", "Folios",
               "Historial", "Estatus_final", "Pestanas_encontrado", "Cliente_seguimiento",
               "Tel_seguimiento", "Folio_seguimiento", "Asesor_seguimiento", "Fecha_ingreso_seg",
               "Similitud_nombre", "Observaciones"]
COLS_SOLO = ["Clasificacion", "Hoja", "Fila_excel", "Fecha_ingreso", "Folio_norm", "Nombre", "Telefono",
             "Financiera", "Asesor", "Estatus_seg", "Veces_capturado", "Observaciones"]


COLS_FALTAN = ["Cliente_bitacora", "Tel_bitacora", "Vendedor", "Financiera", "Folios", "Historial", "Estatus_final",
               "Observaciones"]
COLS_CORREGIR = ["Resultado", "Cuadre_pestana", "Cliente_bitacora", "Tel_bitacora", "Cliente_seguimiento",
                 "Tel_seguimiento", "Folio_seguimiento", "Pestanas_encontrado", "Vendedor", "Financiera", "Folios",
                 "Observaciones"]
NOTA_METODO = ("Cada cliente rechazado en la bitácora (capturas de Cetelem y Bitácora de F&I) se busca en el "
               "seguimiento primero por folio, luego por teléfono (validando el nombre) y al final por nombre aunque el "
               "teléfono sea distinto. Un cliente que aparece en varios meses se cuenta una vez con su último estatus. "
               "Solo se cuadran las solicitudes RECHAZADAS, porque la base de seguimiento registra únicamente rechazos.")


def subtitulo(meses, hojas, financieras):
    return (f"Bitácora: {', '.join(meses)} · Financieras: {', '.join(financieras)} · "
            f"Pestañas: {', '.join(hojas)} · Generado: {datetime.now():%d/%m/%Y %H:%M}")


def exportar_excel(ruta, resumen, cuadre, faltantes, revisar, solo, sub=""):
    sal.excel(ruta, f"Seguimiento de rechazos vs bitácora – {sal.PDV}", sub, [
        ("Resumen", resumen, ["Concepto", "Valor"]),
        ("Faltan en seguimiento", faltantes, COLS_FALTAN),
        ("A corregir", revisar, COLS_CORREGIR),
        ("Solo en seguimiento", solo, COLS_SOLO),
        ("Cuadre por cliente", cuadre, COLS_CUADRE),
    ], COLORES)


def exportar_pdf(ruta, resumen, faltantes, no_cuadra, verificar, solo_periodo, sub="", cifras=None):
    pdf = sal.PDF(ruta, f"Seguimiento de rechazos vs bitácora – {sal.PDV}", sub)
    pdf.cifras(cifras or [])
    pdf.seccion("Rechazos que faltan en el seguimiento", faltantes.sort_values(["Financiera", "Vendedor"]),
                ["Cliente_bitacora", "Tel_bitacora", "Vendedor", "Financiera", "Folios", "Historial", "Observaciones"],
                [4.6, 2.3, 4.2, 2.4, 2.8, 5.4, 3.5], "#C00000",
                "Capturarlos en la pestaña MG Colima Rechazos.",
                [("Por financiera", faltantes.groupby("Financiera").size()),
                 ("Por vendedor", faltantes.groupby("Vendedor").size().sort_values(ascending=False))]
                if not faltantes.empty else None)
    pdf.seccion("Capturados en la pestaña equivocada", no_cuadra,
                ["Cliente_bitacora", "Historial", "Pestanas_encontrado", "Observaciones"], [5.5, 6, 5, 8.7],
                "#C55A11", "Moverlos a la pestaña que corresponde a su último estatus.")
    pdf.seccion("Datos a corregir en el seguimiento", verificar,
                ["Resultado", "Cliente_bitacora", "Tel_bitacora", "Cliente_seguimiento", "Tel_seguimiento",
                 "Observaciones"], [4.2, 4.6, 2.3, 4.6, 2.3, 7.2], "#BF8F00",
                "Nombre, teléfono, folio o fecha distintos a la bitácora, o capturas duplicadas. "
                "La columna Observaciones dice qué corregir.")
    pdf.seccion("En el seguimiento pero no en la bitácora", solo_periodo,
                ["Hoja", "Fila_excel", "Fecha_ingreso", "Folio_norm", "Nombre", "Telefono", "Asesor", "Observaciones"],
                [3.6, 1.2, 2.2, 2.2, 5, 2.4, 4.5, 4.2], "#7F7F7F",
                "Capturados como rechazo, con fecha dentro de los meses de la bitácora, que ahí no salen como "
                "rechazados. Observaciones dice con qué estatus aparecen.")
    pdf.guardar(resumen, NOTA_METODO)


# =============================================================================
# 6. PRINCIPAL
# =============================================================================
def procesar(ruta_seg, rutas_cet=None, carpeta=".", bitacora=None):
    """Hace todo el cuadre y devuelve un diccionario con tablas y rutas de salida.
    bitacora: DataFrame de fuentes.solicitudes_de_bitacora() (todas las financieras). Si no se da,
    se leen los Excel mensuales de Cetelem de `rutas_cet`."""
    cet_todo = desde_bitacora(bitacora) if bitacora is not None else leer_cetelem(rutas_cet)
    cet = cet_todo[cet_todo["Estatus"].isin(ESTATUS_A_CUADRAR)].reset_index(drop=True)
    if cet.empty:
        raise ValueError("La bitácora no tiene solicitudes rechazadas.")
    con_fecha = cet_todo[cet_todo["_anio"] > 0]
    anio_ref = int(con_fecha["_anio"].max()) if not con_fecha.empty else datetime.now().year
    seg = leer_seguimiento(ruta_seg, anio_ref)
    hojas = list(dict.fromkeys(seg["Hoja"]))

    periodos = con_fecha[["_anio", "_mes"]].drop_duplicates().sort_values(["_anio", "_mes"])
    if periodos.empty:
        raise ValueError("La bitácora no trae fechas.")
    a0, m0 = periodos.iloc[0]; a1, m1 = periodos.iloc[-1]
    ini = datetime(int(a0), int(m0), 1)
    fin = datetime(int(a1), int(m1), calendar.monthrange(int(a1), int(m1))[1], 23, 59)
    meses = list(dict.fromkeys(con_fecha.sort_values(["_anio", "_mes"])["Mes"]))
    financieras = sorted(set(cet_todo["Financiera"]) - {""}, key=lambda f: (f != "CETELEM", f))

    cli = consolidar_clientes(cet)
    cuadre, solo = cruzar(cli, seg)
    solo = clasificar_sobrantes(solo, ini, fin, financieras)

    # Si un rechazo del seguimiento no está entre los rechazos de la bitácora, se dice con qué
    # estatus aparece (p. ej. se capturó como rechazo pero quedó aprobado).
    por_folio = {r["Folio_norm"]: r for _, r in cet_todo.iterrows() if r["Folio_norm"]}
    por_tel = {r["Tel_norm"]: r for _, r in cet_todo.iterrows() if r["Tel_norm"]}
    def estatus_en_capturas(r):
        c = por_folio.get(r["Folio_norm"]) if r["Folio_norm"] else None
        if c is None and r["Tel_norm"]:
            c = por_tel.get(r["Tel_norm"])
        if c is None:
            return ""
        return (f"En la bitácora: {c['Estatus']} con {c['Financiera']} "
                f"({c['Mes']}, folio {c['Folio'] or '—'}, {c['Nombre']})")
    if not solo.empty:
        extra = solo.apply(estatus_en_capturas, axis=1)
        solo["Observaciones"] = (solo["Observaciones"].str.strip() + " " + extra).str.strip()

    enc = cuadre["Resultado"] != "NO capturado en seguimiento"
    faltantes = cuadre[~enc]
    no_cuadra = cuadre[enc & (cuadre["Cuadre_pestana"] == "No cuadra")]
    # A corregir: cruces dudosos (nombre o teléfono distinto, posibles) o con algo que corregir en Observaciones.
    # Un cruce por teléfono y nombre sin observaciones ya está bien y no se lista.
    verificar = cuadre[enc & (cuadre["Resultado"].isin([TIPOS[2], TIPOS[3], TIPOS[4]]) |
                              cuadre["Observaciones"].str.contains(
                                  "distinto|≠|no coincide|veces|inválida|financiera|Falta el folio"))
                       & ~cuadre["Observaciones"].str.fullmatch(r"En la bitácora: [\w ]+, confirmar pestaña")]
    revisar = cuadre[enc & ((cuadre["Cuadre_pestana"] != "OK") | cuadre.index.isin(verificar.index))]
    solo_periodo = solo[solo["Clasificacion"] == SOLO_PERIODO]

    cuenta = lambda r: int((cuadre["Resultado"] == r).sum())
    por_fin = cli["Financiera"].str.split(", ").explode().value_counts()
    resumen = pd.DataFrame([
        ("Solicitudes en la bitácora (todos los estatus y financieras)", len(cet_todo)),
        ("Solicitudes RECHAZADAS en la bitácora", len(cet)),
        ("Clientes únicos rechazados", len(cli)),
        ("   · por financiera", " · ".join(f"{k} {v}" for k, v in por_fin.items())),
        ("Registros en pestañas de Colima del seguimiento (sin duplicados)", len(seg)),
        ("Clientes rechazados encontrados en el seguimiento", int(enc.sum())),
        ("   · por folio", cuenta(TIPOS[0])),
        ("   · por teléfono y nombre", cuenta(TIPOS[1])),
        ("   · por nombre, teléfono distinto", cuenta(TIPOS[2])),
        ("   · por teléfono, nombre distinto", cuenta(TIPOS[3])),
        ("   · posibles (revisar)", cuenta(TIPOS[4])),
        ("Clientes rechazados NO capturados en el seguimiento", len(faltantes)),
        ("Porcentaje de cuadre", f"{enc.mean():.1%}" if len(cuadre) else "—"),
        ("Capturados en pestaña equivocada", len(no_cuadra)),
        ("Datos a corregir en el seguimiento", len(verificar)),
        ("En seguimiento pero no en la bitácora del periodo", len(solo_periodo)),
        ("Seguimiento fuera del periodo o sin fecha (informativo)",
         int(solo["Clasificacion"].str.contains("Fuera|Sin fecha").sum())),
        ("Seguimiento de financieras que no están en la bitácora (no aplica)",
         int(solo["Clasificacion"].str.startswith("Financiera que no").sum())),
    ], columns=["Concepto", "Valor"])

    base_nombre = sal.nombre_archivo("Seguimiento", sal.periodo_de_meses(meses))
    rx = os.path.join(carpeta, base_nombre + ".xlsx")
    rp = os.path.join(carpeta, base_nombre + ".pdf")
    sub = subtitulo(meses, hojas, financieras)
    cifras = [
        ("Rechazados en la bitácora", len(cli), "clientes únicos", "neutro"),
        ("Encontrados", int(enc.sum()), f"{enc.mean():.0%} de cuadre" if len(cuadre) else "", "bien"),
        ("Faltan en seguimiento", len(faltantes), "por capturar", "mal" if len(faltantes) else "bien"),
        ("Pestaña equivocada", len(no_cuadra), "por mover", "aviso" if len(no_cuadra) else "bien"),
        ("Datos a corregir", len(verificar), "ver Observaciones", "aviso" if len(verificar) else "bien"),
    ]
    exportar_excel(rx, resumen, cuadre, faltantes, revisar, solo, sub)
    exportar_pdf(rp, resumen, faltantes, no_cuadra, verificar, solo_periodo, sub, cifras)
    return dict(resumen=resumen, cuadre=cuadre, faltantes=faltantes, no_cuadra=no_cuadra,
                verificar=verificar, revisar=revisar, solo=solo, solo_periodo=solo_periodo,
                meses=meses, hojas=hojas, financieras=financieras, xlsx=rx, pdf=rp)


def main():
    ruta_seg, rutas_cet = obtener_archivos()
    print(f"Seguimiento: {ruta_seg}\nCetelem: {', '.join(rutas_cet)}")
    r = procesar(ruta_seg, rutas_cet)
    resumen, rx, rp = r["resumen"], r["xlsx"], r["pdf"]
    print("\n" + resumen.to_string(index=False))
    print(f"\n✔ Archivos generados: {rx} y {rp}")
    try:
        from google.colab import files
        files.download(rx); files.download(rp)
    except ImportError:
        pass


if __name__ == "__main__":
    main()
