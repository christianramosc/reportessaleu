# -*- coding: utf-8 -*-
"""
Registro en Sale-U de la hoja «MG Colima Rechazos»: cómo está cada rechazo en Sale-U.

Parte de lo que propone saleu_sheet.proponer() (Registrado / Falta crédito / No Registrado; busca por teléfono o por
nombre) y agrega lo que ese cuadre no revisa:
  · todos los créditos que tiene el cliente en el reporte de Crédito (sin quitar duplicados: uno por financiera),
  · si el crédito con la financiera de la hoja está rechazado en Sale-U, o si solo tiene crédito con otra financiera,
  · con qué nombre y teléfono está capturado cuando coincidió solo por uno de los dos datos,
  · las filas fuera del periodo del reporte de Crédito (el cuadre no las toca).

    rep = reporte(base, cambios, (ini, fin), creditos, fuente)
    hojas_excel(rep)  → hojas para el Excel del cuadre de Sale-U (o excel(rep) si va solo)
    pdf(rep)          → bytes del PDF
"""
from datetime import datetime

import pandas as pd

from cuadres import saleu as cu
from cuadres import saleu_fuente as sf
from cuadres import salidas as sal
from cuadres.actualizar import leer_rechazos
from cuadres.base import TAB_RECHAZOS

# Grupos, del más urgente al que ya cuadra: (clave, título, qué hacer, color del PDF, tono de la cifra)
GRUPOS = [
    ("alta", "No Registrado: falta darlo de alta en Sale-U",
     "No está ni en el reporte de Crédito ni en el Fuente: darlo de alta como lead y registrar el crédito rechazado.",
     "#C00000"),
    ("credito", "Falta registrar el crédito (ya es lead)",
     "Está en el reporte Fuente pero no en el de Crédito: buscar el lead como dice «Cómo está en Sale-U» y "
     "registrar el crédito rechazado, sin darlo de alta otra vez.", "#BF8F00"),
    ("no_sale", "Dice «Registrado» pero no aparece en el reporte de Crédito",
     "El cuadre lo deja igual (nunca baja de nivel lo que ya dice el Sheet). Confirmarlo en Sale-U.", "#C55A11"),
    ("estatus", "Registrado, pero en Sale-U el crédito no está rechazado",
     "Sale-U tiene el crédito con la misma financiera como Aceptado o En proceso. Actualizar el estatus en Sale-U, "
     "o la hoja si de verdad se aprobó.", "#C55A11"),
    ("otra_fin", "Registrado, pero con otra financiera",
     "El cliente está en Sale-U, pero con crédito de otra financiera: el rechazo de la hoja no está capturado. "
     "Registrar ese crédito.", "#BF8F00"),
    ("dato", "Registrado, con nombre o teléfono distinto en Sale-U",
     "Coincidió solo por teléfono o solo por nombre. Corregir el dato en Sale-U (o en la hoja) con lo que dice "
     "«Cómo está en Sale-U».", "#2E75B6"),
    ("fuera", "Fuera del periodo del reporte de Crédito",
     "La fecha de ingreso cae fuera del periodo del reporte de Crédito: el cuadre no las toca. Descarga el reporte "
     "con un rango que las cubra para revisarlas.", "#7F7F7F"),
    ("ok", "Registrado y cuadra", "Sin pendientes.", "#1E8449"),
]
TITULO = {k: t for k, t, _, _ in GRUPOS}
QUE_HACER = {k: q for k, _, q, _ in GRUPOS}
COLOR = {k: c for k, _, _, c in GRUPOS}
_EQUIV = {"CELETEM": "CETELEM", "BANCOMER": "BBVA"}

COLS_DETALLE = ["Situacion", "Fila", "ID", "Fecha_ingreso", "Cliente", "Celular", "Financiera", "Asesor",
                "Registro_hoja", "Registro_cuadre", "Cambia", "Como_en_SaleU", "Creditos_SaleU",
                "Estatus_misma_financiera", "Lead_Fuente", "Observaciones"]
ETIQUETAS = {"Situacion": "Situación", "Fecha_ingreso": "Fecha ingreso", "Registro_hoja": "Registro en Sale-U (hoja)",
             "Registro_cuadre": "Registro según el cuadre", "Como_en_SaleU": "Cómo está en Sale-U",
             "Creditos_SaleU": "Créditos en Sale-U", "Estatus_misma_financiera": "Estatus en Sale-U con su financiera",
             "Lead_Fuente": "Lead en Fuente"}
sal.ETIQUETAS.update({k: v for k, v in ETIQUETAS.items() if k not in sal.ETIQUETAS})
sal.COLS_LARGAS.update({"Como_en_SaleU", "Creditos_SaleU", "Lead_Fuente"})


def financiera(v):
    k = cu.quitar_acentos(v or "").upper().strip()
    return _EQUIV.get(k, k)


def _texto(v):
    return "" if v is None or (isinstance(v, float) and v != v) else " ".join(str(v).split())


def _fecha(v):
    d = pd.to_datetime(v, errors="coerce", dayfirst=True)
    return d.strftime("%d/%m/%Y") if pd.notna(d) else ""


def _creditos_de(creditos, nombre, tel):
    """Créditos del cliente en el reporte de Crédito: por teléfono; si no, por nombre (mismas reglas del cuadre)."""
    if creditos is None or creditos.empty:
        return creditos
    if tel:
        m = creditos[creditos["Tel_norm"] == tel]
        if len(m):
            return m
    nom = cu.normalizar_nombre(nombre)
    sims = creditos["Nombre_norm"].map(lambda n: cu.similitud_nombre(nom, n))
    return creditos[sims.map(lambda s: s[0] >= cu.UMBRAL_NOMBRE_FUERTE and s[1] >= 2)]


def _como_saleu(p):
    """«Coincidió por … · Se encuentra como «…» con teléfono: …» (el teléfono va cuando coincidió solo por un dato)."""
    if p is None:
        return ""
    if _texto(p.get("Cliente_SaleU")):
        return "Crédito: " + sf.describir(p["Cliente_SaleU"], p.get("Como"), p.get("Tel_SaleU"))
    if _texto(p.get("Cliente_Fuente")):
        return "Fuente (sin crédito): " + sf.describir(p["Cliente_Fuente"], p.get("Como_Fuente"), p.get("Tel_Fuente"))
    return ""


def reporte(base, cambios, periodo, creditos, fuente=None):
    """base: el Sheet; cambios: DataFrame de saleu_sheet.proponer(); periodo: (ini, fin) del reporte de Crédito;
    creditos: cu.leer_saleu(..., quitar_duplicados=False); fuente: leads del reporte Fuente (saleu_fuente)."""
    ini, fin = periodo
    prop = cambios[cambios["Hoja"] == TAB_RECHAZOS].set_index("Fila") if len(cambios) else pd.DataFrame()
    buscador = sf.Buscador(fuente) if fuente is not None and len(fuente) else None
    id_de_fila, asesor_de_fila = _ids_y_asesores(base)
    filas = []
    for r in leer_rechazos(base):
        nombre, tel = _texto(r["nombre"]), str(r["cel"]) if r["cel"] else ""
        tel = cu.normalizar_tel(tel)
        fin_hoja = financiera(r["fin"])
        actual = _texto(r["saleu"])
        p = prop.loc[r["fila"]] if r["fila"] in prop.index else None
        if isinstance(p, pd.DataFrame):
            p = p.iloc[0]
        en_periodo = bool(r["fecha"] and ini <= r["fecha"] <= fin)
        registro = p["Después"] if p is not None else actual
        cs = _creditos_de(creditos, nombre, tel)
        cs = cs.sort_values("_fecha") if cs is not None and len(cs) else cs
        lista = "; ".join(f"{financiera(c['Financiera']) or '(sin financiera)'} {_texto(c['Estatus_SaleU']).lower()}"
                          f" ({_fecha(c['_fecha'])})" for _, c in cs.iterrows()) if cs is not None and len(cs) else ""
        misma = cs[cs["Financiera"].map(financiera) == fin_hoja] if cs is not None and len(cs) else None
        est_misma = (", ".join(sorted({_texto(x).lower() for x in misma["Estatus_SaleU"]}))
                     if misma is not None and len(misma) else "")
        lead = None
        if buscador is not None:
            lead, _ = buscador.buscar(nombre, tel)
        como = _como_saleu(p)
        obs = []
        # ---- situación (la más urgente manda)
        if not en_periodo and p is None:
            sit = "fuera"
            obs.append(f"Ingreso fuera del periodo del reporte de Crédito ({ini:%d/%m/%Y} – {fin:%d/%m/%Y})")
        elif registro == "No Registrado":
            sit = "alta"
            obs.append("No está en Crédito ni en Fuente")
        elif registro == "Falta crédito":
            sit = "credito"
            obs.append("Está como lead en Fuente, sin crédito")
        elif p is not None and _texto(p.get("Revisar")):
            sit = "no_sale"
            obs.append(p["Revisar"])
        elif cs is not None and len(cs) and fin_hoja and misma is not None and len(misma) and "rechazado" not in est_misma:
            sit = "estatus"
            obs.append(f"El crédito con {fin_hoja} está «{est_misma}» en Sale-U; en la hoja está rechazado")
        elif cs is not None and len(cs) and fin_hoja and (misma is None or misma.empty):
            sit = "otra_fin"
            obs.append(f"En Sale-U no hay crédito con {fin_hoja}; solo con otra financiera")
        elif p is not None and "distinto" in str(p.get("Como") or ""):
            sit = "dato"
        else:
            sit = "ok"
        if p is not None and "distinto" in str(p.get("Como") or "") and sit != "dato":
            obs.append("Coincidió solo por " + sf.por_que(p["Como"]))
        if not fin_hoja:
            obs.append("La hoja no tiene financiera")
        filas.append(dict(
            Situacion=TITULO[sit], _clave=sit, Fila=r["fila"], ID=_texto(id_de_fila.get(r["fila"], "")),
            Fecha_ingreso=datetime.combine(r["fecha"], datetime.min.time()) if r["fecha"] else None,
            Cliente=nombre, Celular=tel, Financiera=fin_hoja, Asesor=_texto(asesor_de_fila.get(r["fila"], "")).title(),
            Registro_hoja=actual, Registro_cuadre=registro,
            Cambia="Sí" if p is not None and bool(p.get("Cambia")) else "No",
            Como_en_SaleU=como, Creditos_SaleU=lista, Estatus_misma_financiera=est_misma,
            Lead_Fuente=(f"«{_texto(lead['Nombre'])}» tel. {lead['Telefono'] or 'sin teléfono'} · alta "
                         f"{sf.alta_txt(lead)} · {_texto(lead.get('Fuente'))} · {_texto(lead.get('Estatus_Lead'))}")
            if lead is not None else "",
            Observaciones=" · ".join(obs)))
    det = pd.DataFrame(filas)
    orden = {k: i for i, (k, *_) in enumerate(GRUPOS)}
    if not det.empty:
        det = det.sort_values(["Fila"]).reset_index(drop=True)
    resumen = pd.DataFrame([(t, int((det["_clave"] == k).sum()) if len(det) else 0, q) for k, t, q, _ in GRUPOS],
                           columns=["Situacion", "Filas", "Que_hacer"])
    resumen = resumen.rename(columns={"Que_hacer": "Qué hacer"})
    rep = dict(detalle=det, resumen=resumen, periodo=(ini, fin), orden=orden,
               periodo_txt=f"{ini:%d/%m/%Y} – {fin:%d/%m/%Y}",
               periodo_fuente=sf.periodo_txt(fuente) if fuente is not None else "")
    rep["por_revisar"] = det[~det["_clave"].isin(["ok", "fuera"])] if len(det) else det
    nombre = sal.nombre_archivo("SaleU_Rechazos", sal.periodo_de_fechas(ini, fin))
    rep.update(pdf_bytes=pdf(rep), pdf_nombre=nombre + ".pdf", xlsx_nombre=nombre + ".xlsx")
    return rep


def _ids_y_asesores(base):
    """{fila: ID (R-0001…)} y {fila: Asesor de Ventas} de MG Colima Rechazos."""
    from cuadres import ids
    from cuadres.base import norm
    try:
        id_de_fila = ids.mapa(base)
    except Exception:
        id_de_fila = {}
    df = base.leer(TAB_RECHAZOS)
    c = next((c for c in df.columns if norm(c) == norm("Asesor de Ventas")), None)
    asesor = {int(r["_fila"]): r[c] for _, r in df.iterrows()} if c else {}
    return id_de_fila, asesor


def hojas_excel(rep):
    """Hojas para el Excel del cuadre de Sale-U: resumen, por revisar (con su orden de urgencia) y detalle."""
    det = rep["detalle"]
    pr = rep["por_revisar"]
    if len(pr):
        pr = pr.assign(_o=pr["_clave"].map(rep["orden"])).sort_values(["_o", "Fila"])
    return [("Rechazos · Resumen", rep["resumen"], ["Situacion", "Filas", "Qué hacer"]),
            ("Rechazos · Por revisar", pr, COLS_DETALLE),
            ("Rechazos · Detalle", det, COLS_DETALLE)]


def excel(rep):
    """Excel solo con el Registro en Sale-U de Rechazos (cuando no hay cuadre contra la bitácora)."""
    h = hojas_excel(rep)
    titulo = "Registro en Sale-U · MG Colima Rechazos"
    sub = f"Reporte de Crédito {rep['periodo_txt']} · Generado: {datetime.now():%d/%m/%Y %H:%M}"
    return sal.excel(None, titulo, sub, h)


def pdf(rep):
    det = rep["detalle"]
    ini, fin = rep["periodo"]
    sub = (f"Reporte de Crédito {rep['periodo_txt']}" + (f" · Fuente {rep['periodo_fuente']}" if rep["periodo_fuente"]
                                                         else "") + f" · Generado: {datetime.now():%d/%m/%Y %H:%M}")
    doc = sal.PDF(None, "Registro en Sale-U · MG Colima Rechazos", sub)
    n = lambda k: int((det["_clave"] == k).sum()) if len(det) else 0
    revisar = len(rep["por_revisar"])
    doc.cifras([
        ("Rechazos en la hoja", len(det), "filas con cliente", "neutro"),
        ("Registrado", int((det["Registro_cuadre"] == "Registrado").sum()) if len(det) else 0, "según el cuadre", "bien"),
        ("No Registrado", n("alta"), "dar de alta en Sale-U", "mal" if n("alta") else "bien"),
        ("Falta crédito", n("credito"), "ya son lead", "aviso" if n("credito") else "bien"),
        ("Por revisar", revisar, "financiera, estatus o datos", "aviso" if revisar else "bien"),
    ])
    cols = ["Fila", "ID", "Fecha_ingreso", "Cliente", "Celular", "Financiera", "Asesor", "Creditos_SaleU",
            "Como_en_SaleU", "Observaciones"]
    anchos = [1.0, 1.8, 2.2, 3.4, 2.5, 2.2, 3.0, 3.6, 5.0, 4.2]
    for k, titulo, que, color in GRUPOS:
        if k == "ok":
            continue
        doc.seccion(titulo, det[det["_clave"] == k] if len(det) else det, cols, anchos, color, que)
    return doc.guardar(rep["resumen"].rename(columns={"Situacion": "Concepto", "Filas": "Valor"}),
                       "Lógica del cuadre de Sale-U (saleu_sheet.py): cada rechazo con fecha de ingreso dentro del "
                       "periodo del reporte de Crédito se busca por teléfono o por nombre; si está en Crédito queda "
                       "«Registrado», si solo está en Fuente «Falta crédito» y si no está en ninguno «No Registrado». "
                       "Además se revisa con qué financiera y estatus está el crédito en Sale-U contra lo que dice la "
                       "hoja. Cuando coincidió solo por nombre se indica el teléfono con que está en Sale-U, y si fue "
                       "solo por teléfono, el nombre.")
