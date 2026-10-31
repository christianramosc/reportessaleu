# -*- coding: utf-8 -*-
"""
Actualizar la base de seguimiento con las bitácoras de Cetelem y la "Bitácora de F&I".

  proponer(base, solicitudes)  → tabla de cambios propuestos (nada se escribe todavía)
  aplicar(base, propuestas, …) → respalda, escribe lo elegido, anota el Historial y
                                  reescribe las pestañas "Capturas Cetelem" y "Bitacora"
"""
import logging
import datetime as dt

import pandas as pd

from cuadres.base import (TAB_RECHAZOS, TAB_FINALIZADOS, TAB_BITACORA, TAB_CAPTURAS, TAB_FI, TAB_HISTORIAL,
                          norm, letra, canonico)
from cuadres import ids
from cuadres.fuentes import (name_key, similar, to_folio, to_phone, to_num, to_date, folio_date,
                             BITACORA_COLS, CAP_COLS, CAP_EXTRA, filas_bitacora, FI_COLS, filas_fi,
                             COL_SALEU_BITACORA, COL_ID_RECHAZOS)

AZUL_NUEVO = "#CFE2F3"     # mismo azul claro de "Agregado desde la bitácora" en la leyenda del Resumen

# Encabezados de MG Colima Rechazos que usa la app
C = dict(fecha="Fecha ingreso", folio="Folio", nombre="Nombre de Cliente", cel="Celular", perfil="Perfil",
         fin="Financiera", plazo="Plazo", unidad="Unidad", asesor="Asesor de Ventas", gap="GAP",
         acc="Accesorios", status="Status 1", mens="Mensualidad", correo="Correo",
         obs="Condición / Observaciones", saleu="Registro en Sale-U", fseg="Fecha de Seguimiento",
         com="COMENTARIOS")

log = logging.getLogger("mg_colima")


def _fecha(v):
    d = to_date(v)
    return d


def leer_rechazos(base):
    df = base.leer(TAB_RECHAZOS)
    cols = {k: next((c for c in df.columns if norm(c) == norm(v)), None) for k, v in C.items()}
    filas = []
    for _, r in df.iterrows():
        g = lambda k: r[cols[k]] if cols[k] else None
        nombre = g("nombre")
        if nombre in (None, "") or str(nombre).strip() == "":
            continue
        filas.append(dict(fila=int(r["_fila"]), fecha=_fecha(g("fecha")), folio=to_folio(g("folio")),
                          nombre=str(nombre).strip(), key=name_key(nombre), cel=to_phone(g("cel")),
                          perfil=g("perfil"), fin=norm(g("fin")) or None, plazo=to_num(g("plazo")),
                          gap=norm(g("gap")) or None, acc=norm(g("acc")) or None, mens=to_num(g("mens")),
                          correo=g("correo") or None, saleu=g("saleu") or None, com=g("com") or None))
    return filas


def _match(row, sol, por_folio):
    if row["folio"] and row["folio"] in por_folio and any(similar(row["key"], k) for k in por_folio[row["folio"]]["keys"]):
        return [por_folio[row["folio"]]]
    return [s for s in sol if any(similar(row["key"], k) for k in s["keys"])
            and (not row["fecha"] or not s["fecha"] or abs((row["fecha"] - s["fecha"]).days) <= 45)]


def _ignorados(base):
    if TAB_HISTORIAL not in base.pestanas():
        return set()
    h = base.leer(TAB_HISTORIAL)
    if h.empty or "Origen" not in h.columns:
        return set()
    h = h[h["Origen"].astype(str).str.startswith("Ignorado")]
    return {(norm(r["Cliente"]), norm(canonico(r["Campo"] or "")), norm(r["Después"])) for _, r in h.iterrows()}


def _catalogo_asesores(base):
    if "Catalogos" not in base.pestanas():
        return []
    cat = base.leer("Catalogos")
    return [str(v).strip() for v in cat.iloc[:, 0].tolist() if es_valor_catalogo(v)] if not cat.empty else []


def es_valor_catalogo(v):
    """Descarta celdas vacías y las notas que trae la pestaña Catalogos («Agrega nuevos valores debajo…»)."""
    t = str(v if v is not None else "").strip()
    return bool(t) and t not in ("None", "nan") and len(t) <= 60 and not any(ch in t for ch in "();:")


def _asesor(corto, catalogo):
    """Nombre completo del asesor según Catalogos: 'CARLA' → 'CARLA YOVET CARDENAS AHUMADA',
    'JESSICA MACHUCA' → 'JESSICA ELIZABETH MACHUCA ROQUE'. Si no hay uno solo que corresponda, se deja igual."""
    c = norm(corto)
    if not c:
        return None
    exact = [a for a in catalogo if norm(a) == c]
    if exact:
        return exact[0]
    palabras = c.split()
    contiene = [a for a in catalogo if set(palabras) <= set(norm(a).split())]     # todas sus palabras están
    if len(contiene) == 1:
        return contiene[0]
    mismo = [a for a in catalogo if norm(a).split()[0] == palabras[0]]
    return mismo[0] if len(mismo) == 1 and len(palabras) == 1 else str(corto).strip()


def proponer(base, solicitudes):
    """Devuelve (propuestas: DataFrame, contexto: dict). No escribe nada."""
    filas = leer_rechazos(base)
    por_folio = {s["folio"]: s for s in solicitudes if s["folio"]}
    ignor = _ignorados(base)
    catalogo = _catalogo_asesores(base)
    prop = []

    def add(tipo, row, campo, antes, despues, motivo, aplicar=True, payload=None, sol=None, hoja=TAB_RECHAZOS):
        cli = row["nombre"] if row else (payload or {}).get(C["nombre"], "")
        if (norm(cli), norm(campo), norm(despues)) in ignor:
            return
        prop.append(dict(Aplicar=aplicar, Tipo=tipo, Hoja=hoja, Fila=row["fila"] if row else None, Cliente=cli,
                         Campo=campo, Antes=_fmt(antes), Después=_fmt(despues), Motivo=motivo,
                         _valor=despues, _payload=payload, _sol=sol))

    usadas, fila_de_sol = set(), {}
    for row in filas:
        ms = _match(row, solicitudes, por_folio)
        for s in ms:
            usadas.add(id(s))
            fila_de_sol.setdefault(id(s), []).append(row["fila"])
            oficial = s["mes_cetelem"] and row["folio"] == s["folio"]
            mes = s["mes_cetelem"]
            if oficial:
                if s["mensualidad"] and (not row["mens"] or abs(row["mens"] - s["mensualidad"]) > 0.6):
                    add("Corrección", row, C["mens"], row["mens"], s["mensualidad"],
                        f"Mensualidad máxima según la bitácora de Cetelem de {mes}")
                if s["plazo"] and row["plazo"] and int(s["plazo"]) != int(row["plazo"]):
                    add("Corrección", row, C["plazo"], int(row["plazo"]), int(s["plazo"]), f"Plazo según Cetelem ({mes})")
                if s["telefono"] and row["cel"] and s["telefono"] != row["cel"]:
                    add("Corrección", row, C["cel"], row["cel"], s["telefono"], f"Teléfono registrado en Cetelem ({mes})")
                for k, campo, fld in (("gap", C["gap"], "gap"), ("acc", C["acc"], "accesorios")):
                    if s.get(fld) in ("SI", "NO") and row[k] != s[fld]:
                        add("Corrección", row, campo, row[k], s[fld], f"Según la bitácora de Cetelem de {mes}")
                if s.get("tipo_persona") == "MORAL" and norm(row["perfil"]) != "MORAL":
                    add("Corrección", row, C["perfil"], row["perfil"], "MORAL", f"Cetelem lo registra como persona moral ({mes})")
                fd = folio_date(s["folio"], (row["fecha"] or s["fecha"] or dt.date.today()).year)
                if row["fecha"] and fd and abs((row["fecha"] - fd).days) > 3:
                    add("Corrección", row, C["fecha"], row["fecha"], fd,
                        f"El folio {s['folio']} se generó el {fd:%d/%m/%Y} (los primeros dígitos del folio son el día del año)")
            if not row["folio"] and s["folio"] and s["mes_cetelem"] and s["est_sol"] == "Rechazado":
                add("Corrección", row, C["folio"], None, s["folio"], f"Folio del rechazo en la bitácora de Cetelem de {mes}")
            if not row["saleu"] and s["sale_u"] in ("SI", "NO"):
                add("Corrección", row, C["saleu"], None, "Registrado" if s["sale_u"] == "SI" else "No Registrado",
                    "La bitácora de F&I dice SALE_U = " + s["sale_u"])
            if not (s["est_sol"] == "Rechazado" or (row["folio"] and row["folio"] == s["folio"])):
                continue        # otra solicitud del mismo cliente (p. ej. aprobada en otro mes): no se copian datos
            if not row["correo"] and s["correo"]:
                add("Corrección", row, C["correo"], None, s["correo"], "Correo de la bitácora")
            if not row["cel"] and s["telefono"]:
                add("Corrección", row, C["cel"], None, s["telefono"], "Teléfono de la bitácora")
            if not row["mens"] and s["mensualidad"] and not oficial:
                add("Corrección", row, C["mens"], None, s["mensualidad"], "Mensualidad de la bitácora")
            if s["est_fin"] == "Formalizado" and "YA COMPRO" not in norm(row["com"]):
                t = f"Ya compró según la bitácora: formalizó con {s['fin_banco']}" + (f", factura {s['factura']}" if s["factura"] else "")
                add("Comentario", row, C["com"], row["com"], (row["com"] + " | " if row["com"] else "") + t,
                    "La bitácora registra el crédito formalizado; no requiere seguimiento de rechazo")

    # ---------------------------------------------------------------- rechazos nuevos
    claves_hoja = [r["key"] for r in filas]
    for s in solicitudes:
        if s["est_sol"] != "Rechazado" or id(s) in usadas:
            continue
        if any(similar(k, kk) for kk in s["keys"] for k in claves_hoja):
            continue
        aplicar, nota = True, []
        rechazo_cet = any(o["origen"].startswith("CETELEM") and o["estatus"] == "Rechazado" for o in s["origenes"])
        if not s["mes_cetelem"] and norm(s["banco"]) == "CETELEM":
            otros = [t for t in solicitudes if t is not s and t["mes_cetelem"] and any(similar(k, kk) for k in s["keys"] for kk in t["keys"])]
            if otros:
                aplicar = False
                nota.append("En Cetelem aparece como " + ", ".join(f"{t['est_sol']} (folio {t['folio']})" for t in otros) + ": confirmar")
        fin = "CETELEM" if rechazo_cet else (s["banco"] if norm(s["banco"]) not in ("", "OTROS") else s["canal2"])
        com = [f"Agregado desde la bitácora ({s['fuentes'].split(',')[0]})"]
        if s["est_fin"] == "Formalizado":
            com.append(f"Ya compró según la bitácora: formalizó con {s['fin_banco']}" + (f", factura {s['factura']}" if s["factura"] else ""))
        elif s["est_fin"] in ("Aprobado", "Contrapropuesta"):
            com.append(f"Después aparece como {s['est_fin']} con {s['fin_banco']}")
        payload = {
            C["fecha"]: s["fecha"], C["folio"]: s["folio"], C["nombre"]: (s["cliente"] or "").upper(),
            C["cel"]: s["telefono"], C["perfil"]: "MORAL" if s.get("tipo_persona") == "MORAL" else None,
            C["fin"]: fin, C["plazo"]: int(s["plazo"]) if s["plazo"] else None, C["unidad"]: s["unidad"],
            C["asesor"]: _asesor(s["vendedor"], catalogo), C["gap"]: s["gap"] if s["gap"] in ("SI", "NO") else None,
            C["acc"]: s["accesorios"] if s["accesorios"] in ("SI", "NO") else None, C["status"]: "RECHAZADO",
            C["mens"]: s["mensualidad"], C["correo"]: s["correo"], C["obs"]: s["com_fi"] or s["com_as"],
            C["saleu"]: {"SI": "Registrado", "NO": "No Registrado"}.get(s["sale_u"]), C["com"]: " | ".join(com),
        }
        desc = f"{s['fecha']:%d/%m/%Y} · folio {s['folio'] or '—'} · {fin}" if s["fecha"] else f"folio {s['folio'] or '—'} · {fin}"
        add("Nuevo rechazo", None, "(fila nueva)", None, desc,
            "Rechazo en " + s["fuentes"] + (". " + "; ".join(nota) if nota else ""), aplicar, payload, s)

    # ---------------------------------------------------------------- datos que faltan (por teléfono o nombre)
    # Cualquier pestaña: si el cliente cuadra con una solicitud de la bitácora por teléfono o por nombre,
    # las celdas vacías se llenan con su dato (aunque sea otra solicitud del mismo cliente).
    from cuadres import completar
    ya = {(p["Hoja"], p["Fila"], norm(p["Campo"])) for p in prop if p["Fila"] is not None}
    regs = completar.registros_de_solicitudes(solicitudes, lambda v: _asesor(v, catalogo))
    for c in completar.proponer(base, regs):
        if (c["Hoja"], c["Fila"], norm(c["Campo"])) in ya:
            continue
        ya.add((c["Hoja"], c["Fila"], norm(c["Campo"])))
        add("Dato faltante", dict(nombre=c["Cliente"], fila=c["Fila"]), c["Campo"], None, c["_valor"], c["Motivo"],
            hoja=c["Hoja"])

    df = pd.DataFrame(prop, columns=["Aplicar", "Tipo", "Hoja", "Fila", "Cliente", "Campo", "Antes", "Después",
                                     "Motivo", "_valor", "_payload", "_sol"])
    orden = {"Nuevo rechazo": 0, "Corrección": 1, "Dato faltante": 2, "Comentario": 3}
    df = df.sort_values(["Tipo", "Fila"], key=lambda s: s.map(orden) if s.name == "Tipo" else s).reset_index(drop=True)
    return df, dict(filas=filas, fila_de_sol=fila_de_sol)


def _fmt(v):
    if isinstance(v, (dt.date, dt.datetime)):
        return v.strftime("%d/%m/%Y")
    if v is None:
        return ""
    if isinstance(v, float) and v.is_integer():
        return str(int(v))
    return str(v)


# "Contrato" pasó a llamarse "Formalizado" (FINANCIADO de Cetelem y CONTRATO de F&I son lo mismo)
CAMBIOS_FORMALIZADO = [('"Contrato"', '"Formalizado"'), ("Terminaron en contrato", "Formalizados"),
                       ("'Contrato' es el estatus final con cualquier financiera (Cetelem: FINANCIADO).",
                        "'Formalizado' = el crédito se formalizó con cualquier financiera "
                        "(FINANCIADO en Cetelem, CONTRATO en la Bitácora de F&I).")]


def migrar_formalizado(base):
    """Actualiza las fórmulas y textos del Resumen y los colores de la Bitacora al nombre 'Formalizado'.
    No hace nada si ya está migrado. Devuelve cuántas celdas/reglas cambió."""
    n = base.reemplazar_textos("Resumen", CAMBIOS_FORMALIZADO) if "Resumen" in base.pestanas() else 0
    n += base.reemplazar_en_formatos(TAB_BITACORA, CAMBIOS_FORMALIZADO[:1], filas_min=2000)
    return n


FUENTE_SHEET = ("MG Colima Rechazos", "MG Colima Créditos Finalizados")    # «MG Colima Rechazos R-0005»


def filas_bitacora_del_sheet(base, fila_de_sol):
    """Filas para la Bitacora con los rechazos de «MG Colima Rechazos» que no salen de ninguna captura de Cetelem
    ni de la Bitácora de F&I (p. ej. capturados a mano o de meses anteriores). Así la Bitacora es el acumulado."""
    from cuadres.saleu_sheet import solicitudes_del_sheet
    ligadas = {f for filas in fila_de_sol.values() for f in filas}
    out = []
    for _, r in solicitudes_del_sheet(base).iterrows():
        if r["Hoja"] != "Rechazos" or r["Fila"] in ligadas:
            continue
        fila_rech = r["ID"] or r["Fila"]          # el ID no cambia aunque se ordene la hoja
        estatus, fuente = "Rechazado", f"{FUENTE_SHEET[0]} {r['ID'] or 'fila ' + str(r['Fila'])}"
        banco = "" if r["Fin_sol"] == "SIN FINANCIERA" else r["Fin_sol"]
        tel = to_phone(r["Telefono"])
        fecha = to_date(r["Fecha"]) if isinstance(r["Fecha"], dt.date) else None
        out.append([fecha, r["Nombre"], tel, to_folio(r["Folio"]), r["Correo"] or None,
                    None if r["Vendedor"] == "(sin vendedor)" else r["Vendedor"], r["Vehiculo"] or None, None,
                    banco, estatus, None, estatus, banco, r["Plazo"], r["Mensualidad"], None, None, None, None,
                    None, None, fila_rech, fuente])
    return out


COL_VENDEDOR, COL_ESTATUS_SOL = BITACORA_COLS.index("Vendedor"), BITACORA_COLS.index("Estatus solicitud")
COL_ESTATUS_FIN, COL_SALEU = BITACORA_COLS.index("Estatus final"), BITACORA_COLS.index(COL_SALEU_BITACORA)
COL_LIGA = BITACORA_COLS.index(COL_ID_RECHAZOS)
NO_APLICA = "No aplica"
# Mismo estilo que la columna de Sale-U de Rechazos y Finalizados
ESTILO_SALEU = [('="Registrado"', "#D9EAD3", "#274E13"), ('="Falta crédito"', "#FFE599", "#7F6000"),
                ('="No Registrado"', "#F4CCCC", "#990000"), ('="No aplica"', "#EFEFEF", "#434343")]
REGISTRO_TXT = {"REGISTRADO": "Registrado", "NO REGISTRADO": "No Registrado", "FALTA CREDITO": "Falta crédito"}


def _registro_de_rechazos(base):
    """{ID o fila: Registro en Sale-U} de MG Colima Rechazos (lo que dejó el cuadre de Sale-U)."""
    out = {}
    if TAB_RECHAZOS not in base.pestanas():
        return out
    df = base.leer(TAB_RECHAZOS)
    c_su = next((c for c in df.columns if norm(c) == norm(C["saleu"])), None)
    c_id = next((c for c in df.columns if norm(c) == norm(ids.COL_ID)), None)
    if not c_su:
        return out
    for _, r in df.iterrows():
        v = REGISTRO_TXT.get(norm(r[c_su]))
        if v:
            out[str(r["_fila"])] = v
            if c_id and r[c_id]:
                out[str(r[c_id]).strip().upper()] = v
    return out


def _llave_bitacora(fila):
    folio = to_folio(fila[BITACORA_COLS.index("Folio")])
    if folio:
        return str(folio)
    fecha = to_date(fila[0])
    return f"{name_key(fila[BITACORA_COLS.index('Cliente')])}|{fecha or ''}"


def _registro_previo(base):
    """{solicitud: Registro en Sale-U} de la Bitacora actual (lo que marcó el cuadre de Sale-U a los Aprobados),
    para no perderlo cuando la Bitacora se rehace."""
    if TAB_BITACORA not in base.pestanas():
        return {}
    df = base.leer(TAB_BITACORA)
    if df.empty or COL_SALEU_BITACORA not in df.columns:
        return {}
    out = {}
    for _, r in df.iterrows():
        v = REGISTRO_TXT.get(norm(r[COL_SALEU_BITACORA]))
        if v:
            out[_llave_bitacora([r.get(c) for c in BITACORA_COLS])] = v
    return out


def _registro_saleu(fila, registro, previo=None):
    """Registrado / Falta crédito / No Registrado para Aprobados y Rechazados; «No aplica» para los demás estatus.
    Rechazados ligados a Rechazos: lo que dice esa pestaña. Aprobados: lo que dejó el cuadre de Sale-U en la Bitacora
    (un «Registrado» de F&I manda). Si no hay nada de eso, lo que dice F&I (SI / NO)."""
    est = norm(fila[COL_ESTATUS_FIN])
    if not (est.startswith("APROBAD") or est.startswith("RECHAZAD")):
        return NO_APLICA
    fi = {"SI": "Registrado", "NO": "No Registrado"}.get(norm(fila[COL_SALEU]))
    antes = (previo or {}).get(_llave_bitacora(fila))
    if est.startswith("APROBAD") and antes:
        return "Registrado" if fi == "Registrado" else antes
    for liga in str(fila[COL_LIGA] or "").split(","):
        liga = liga.strip().upper()
        if liga.endswith(".0"):
            liga = liga[:-2]
        if liga in registro:
            return registro[liga]
    return fi or antes


def estilo_bitacora(base):
    """Colores y lista de la columna Registro en Sale-U de la Bitacora (no duplica reglas)."""
    c = base.col(TAB_BITACORA, COL_SALEU_BITACORA)
    if not c:
        return
    L = letra(c)
    for valor, fondo, tinta in ([] if getattr(base, "CHIPS", False) else ESTILO_SALEU):   # Google: chips de color
        base.regla_color(TAB_BITACORA, [c], f"=${L}2{valor}", fondo, 1000, tinta)
    base.lista_valores(TAB_BITACORA, [c], ["Registrado", "Falta crédito", "No Registrado", NO_APLICA], 1000)
# Formato uniforme de la Bitacora: tipo de número por columna (lo demás, texto)
FORMATO_BITACORA = {"Fecha": "fecha", "Folio": "entero", "Teléfono": "entero", "Año": "entero", "Plazo": "entero",
                    "Mensualidad": "moneda", "Monto a financiar": "moneda"}


FORMATO_FI = {"Fecha": "fecha", "Folio": "entero", "Teléfono": "entero", "Año": "entero", "Plazo": "entero",
              "Mensualidad": "moneda", "Monto": "moneda", "Fecha de carga": "fechahora"}


def filas_bitacora_completa(base, solicitudes, fila_de_sol, id_de_fila=None):
    """Filas de la pestaña Bitacora: solicitudes de las bitácoras + rechazos que solo están en el Sheet, con todos
    los estatus (formalizadas incluidas), con el vendedor según Catalogos y ordenadas por fecha (sin fecha al final).
    id_de_fila: {fila: ID}; la Bitacora guarda el ID del rechazo (si no hay, el número de fila).
    Registro en Sale-U: igual que en MG Colima Rechazos (Registrado / No Registrado) para Aprobados y Rechazados;
    «No aplica» para los demás estatus."""
    catalogo = _catalogo_asesores(base)
    id_de_fila = id_de_fila or {}
    ids_de_sol = {k: [id_de_fila.get(f, f) for f in filas] for k, filas in fila_de_sol.items()}
    filas = filas_bitacora(solicitudes, ids_de_sol) + filas_bitacora_del_sheet(base, fila_de_sol)
    registro, previo = _registro_de_rechazos(base), _registro_previo(base)
    for f in filas:
        if f[COL_VENDEDOR] and catalogo:
            f[COL_VENDEDOR] = _asesor(f[COL_VENDEDOR], catalogo)
        f[COL_SALEU] = _registro_saleu(f, registro, previo)
    filas.sort(key=lambda f: (f[0] is None or f[0] == "", f[0] or dt.date.max))
    return filas


def aplicar(base, propuestas, solicitudes, contexto, capturas=None, registros_fi=None):
    """Escribe en la base las propuestas con Aplicar=True. Devuelve un resumen."""
    marca = propuestas["Aplicar"].astype(bool) if not propuestas.empty else pd.Series([], dtype=bool)
    elegidas, ignoradas = propuestas[marca.values], propuestas[~marca.values]
    col = {k: base.col(TAB_RECHAZOS, v) for k, v in C.items()}
    faltan = [C[k] for k in ("fecha", "folio", "nombre") if not col[k]]
    if faltan:
        raise ValueError(f"No encuentro las columnas {', '.join(faltan)} en '{TAB_RECHAZOS}'.")
    hoja_de = lambda p: p["Hoja"] if "Hoja" in p and p["Hoja"] else TAB_RECHAZOS
    respaldo = base.respaldar(TAB_RECHAZOS) if len(elegidas) else None
    id_de_fila = ids.asegurar(base)        # ID fijo a los rechazos que no lo tengan (p. ej. capturados a mano)
    c_id, c_rein = base.col(TAB_RECHAZOS, ids.COL_ID), base.col(TAB_RECHAZOS, ids.COL_REINGRESO)
    l_nom, l_tel = ids.letras(base)
    siguiente_id = ids.siguiente(id_de_fila.values())
    if any(hoja_de(p) == TAB_FINALIZADOS for _, p in elegidas.iterrows()):
        base.respaldar(TAB_FINALIZADOS)

    celdas, historial, nuevas_marcas, otras = [], [], [], {}
    for _, p in elegidas[elegidas["Tipo"] != "Nuevo rechazo"].iterrows():
        hoja = hoja_de(p)
        c = base.col(hoja, p["Campo"])
        if not c:
            continue
        (celdas if hoja == TAB_RECHAZOS else otras.setdefault(hoja, [])).append((int(p["Fila"]), c, p["_valor"]))
        historial.append(dict(Origen="App · bitácoras", Hoja=hoja, Celda=f"{letra(c)}{int(p['Fila'])}",
                              Cliente=p["Cliente"], Campo=p["Campo"], Antes=p["Antes"], Después=p["Después"],
                              Motivo=p["Motivo"]))

    nuevos = elegidas[elegidas["Tipo"] == "Nuevo rechazo"]
    fila = base.ultima_fila_con_datos(TAB_RECHAZOS, (col["fecha"], col["nombre"])) + 1
    for _, p in nuevos.iterrows():
        for campo, valor in p["_payload"].items():
            c = base.col(TAB_RECHAZOS, campo)
            if c and valor not in (None, ""):
                celdas.append((fila, c, valor))
        F = letra(col["fecha"])
        if col["fseg"]:
            celdas.append((fila, col["fseg"], f'=IF({F}{fila}="","",{F}{fila}+180)'))
        for i in (1, 2, 3):
            cc, ce = base.col(TAB_RECHAZOS, f"CONTESTA {i}"), base.col(TAB_RECHAZOS, f"EFECTIVIDAD {i}")
            if cc and ce:
                celdas.append((fila, ce, f'=IF(${F}{fila}="","",IF({letra(cc)}{fila}="Si Responde",1,0))'))
        if c_id:
            id_de_fila[fila] = ids.formato(siguiente_id)
            celdas.append((fila, c_id, id_de_fila[fila]))
            siguiente_id += 1
        if c_rein and l_nom:
            celdas.append((fila, c_rein, ids.formula_reingreso(fila, l_nom, l_tel)))
        nuevas_marcas.append((fila, col["nombre"]))
        contexto["fila_de_sol"].setdefault(id(p["_sol"]), []).append(fila)
        historial.append(dict(Origen="App · bitácoras", Hoja=TAB_RECHAZOS, Celda=f"Fila {fila}", Cliente=p["Cliente"],
                              Campo="(fila nueva)", Antes="", Después=p["Después"], Motivo=p["Motivo"]))
        fila += 1

    base.escribir(TAB_RECHAZOS, celdas)
    for hoja, cs in otras.items():
        base.escribir(hoja, cs)
    base.marcar(TAB_RECHAZOS, nuevas_marcas, AZUL_NUEVO)
    for _, p in ignoradas.iterrows():
        historial.append(dict(Origen="Ignorado", Hoja=hoja_de(p), Celda="", Cliente=p["Cliente"], Campo=p["Campo"],
                              Antes=p["Antes"], Después=p["Después"], Motivo="No se aplicó: " + p["Motivo"]))
    base.anotar_historial(historial)

    if capturas is not None and not capturas.empty:
        base.reemplazar(TAB_CAPTURAS, CAP_COLS + CAP_EXTRA,
                        capturas[CAP_COLS + CAP_EXTRA].fillna("").values.tolist(), oculta=True)
    if registros_fi:       # copia de la Bitácora de F&I: así no se pierde si la próxima vez no se sube
        base.reemplazar(TAB_FI, FI_COLS, filas_fi(registros_fi), oculta=True)
        base.dar_formato(TAB_FI, FORMATO_FI)
    # Pasos complementarios: si alguno falla, lo principal ya quedó escrito; se avisa qué faltó en lugar de callarlo.
    avisos = []

    def paso(que, funcion):
        try:
            funcion()
        except Exception as e:
            log.warning("Actualizar: no se pudo %s", que, exc_info=e)
            avisos.append(f"No se pudo {que}: {e}")

    def llenar_registro():                   # rechazos nuevos sin Registro en Sale-U: con los leads guardados
        from cuadres.saleu_sheet import llenar_vacios
        llenar_vacios(base)

    def fijos_y_ligas():                     # columnas fijas de los rechazos nuevos y liga Finalizados ↔ Rechazos
        from cuadres import optimizar
        optimizar.llenar_fijos(base)
        optimizar.ligar_finalizados(base)

    def dashboard_resumen():                 # el Resumen toma financieras/asesores nuevos
        from cuadres import dashboard
        dashboard.construir(base)

    paso("llenar el Registro en Sale-U de los rechazos nuevos", llenar_registro)
    base.reemplazar(TAB_BITACORA, BITACORA_COLS,
                    filas_bitacora_completa(base, solicitudes, contexto["fila_de_sol"], id_de_fila))
    base.dar_formato(TAB_BITACORA, FORMATO_BITACORA)
    paso("dar color a la Bitacora", lambda: estilo_bitacora(base))
    paso("llenar las columnas fijas y ligar Finalizados con Rechazos", fijos_y_ligas)
    paso("actualizar el Resumen a «Formalizado»", lambda: migrar_formalizado(base))
    paso("rehacer el dashboard del Resumen", dashboard_resumen)
    aviso = " · ".join(avisos) or None
    return dict(respaldo=respaldo, aviso=aviso,
                faltantes=int((elegidas["Tipo"] == "Dato faltante").sum()),
                correcciones=int(elegidas["Tipo"].isin(["Corrección", "Comentario"]).sum()),
                nuevos=len(nuevos), ignorados=len(ignoradas), celdas=len(celdas) + sum(map(len, otras.values())))
