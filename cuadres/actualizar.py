# -*- coding: utf-8 -*-
"""
Actualizar la base de seguimiento con las bitácoras de Cetelem y la "Copia de Financiamiento".

  proponer(base, solicitudes)  → tabla de cambios propuestos (nada se escribe todavía)
  aplicar(base, propuestas, …) → respalda, escribe lo elegido, anota el Historial y
                                  reescribe las pestañas "Capturas Cetelem" y "Bitacora"
"""
import datetime as dt

import pandas as pd

from cuadres.base import (TAB_RECHAZOS, TAB_BITACORA, TAB_CAPTURAS, TAB_HISTORIAL, norm, letra)
from cuadres.fuentes import (name_key, similar, to_folio, to_phone, to_num, to_date, folio_date,
                             BITACORA_COLS, CAP_COLS, CAP_EXTRA, filas_bitacora)

AZUL_NUEVO = "#CFE2F3"     # mismo azul claro de "Agregado desde la bitácora" en la leyenda del Resumen

# Encabezados de MG Colima Rechazos que usa la app
C = dict(fecha="Fecha ingreso", folio="Folio", nombre="Nombre de Cliente", cel="Celular", perfil="Perfil",
         fin="Financiera", plazo="Plazo", unidad="Unidad", asesor="Asesor de Ventas", gap="GAP",
         acc="Accesorios", status="Status 1", mens="Mensualidad", correo="Correo",
         obs="Condición / Observaciones", saleu="Cliente con Registro en SalesU", fseg="Fecha de Seguimiento",
         com="COMENTARIOS")


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
    return {(norm(r["Cliente"]), norm(r["Campo"]), norm(r["Después"])) for _, r in h.iterrows()}


def _catalogo_asesores(base):
    if "Catalogos" not in base.pestanas():
        return []
    cat = base.leer("Catalogos")
    return [str(v).strip() for v in cat.iloc[:, 0].tolist() if str(v).strip() not in ("", "None", "nan")] if not cat.empty else []


def _asesor(corto, catalogo):
    c = norm(corto)
    if not c:
        return None
    exact = [a for a in catalogo if norm(a) == c]
    if exact:
        return exact[0]
    mismo = [a for a in catalogo if norm(a).split()[0] == c.split()[0]]
    return mismo[0] if len(mismo) == 1 else str(corto).strip()


def proponer(base, solicitudes):
    """Devuelve (propuestas: DataFrame, contexto: dict). No escribe nada."""
    filas = leer_rechazos(base)
    por_folio = {s["folio"]: s for s in solicitudes if s["folio"]}
    ignor = _ignorados(base)
    catalogo = _catalogo_asesores(base)
    prop = []

    def add(tipo, row, campo, antes, despues, motivo, aplicar=True, payload=None, sol=None):
        cli = row["nombre"] if row else (payload or {}).get(C["nombre"], "")
        if (norm(cli), norm(campo), norm(despues)) in ignor:
            return
        prop.append(dict(Aplicar=aplicar, Tipo=tipo, Fila=row["fila"] if row else None, Cliente=cli, Campo=campo,
                         Antes=_fmt(antes), Después=_fmt(despues), Motivo=motivo,
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
            if s["est_fin"] == "Contrato" and "YA COMPRO" not in norm(row["com"]):
                t = f"Ya compró según la bitácora: contrato con {s['fin_banco']}" + (f", factura {s['factura']}" if s["factura"] else "")
                add("Comentario", row, C["com"], row["com"], (row["com"] + " | " if row["com"] else "") + t,
                    "La bitácora registra contrato; no requiere seguimiento de rechazo")

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
        if s["est_fin"] == "Contrato":
            com.append(f"Ya compró según la bitácora: contrato con {s['fin_banco']}" + (f", factura {s['factura']}" if s["factura"] else ""))
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

    df = pd.DataFrame(prop, columns=["Aplicar", "Tipo", "Fila", "Cliente", "Campo", "Antes", "Después", "Motivo",
                                     "_valor", "_payload", "_sol"])
    orden = {"Nuevo rechazo": 0, "Corrección": 1, "Comentario": 2}
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


def aplicar(base, propuestas, solicitudes, contexto, capturas=None):
    """Escribe en la base las propuestas con Aplicar=True. Devuelve un resumen."""
    marca = propuestas["Aplicar"].astype(bool) if not propuestas.empty else pd.Series([], dtype=bool)
    elegidas, ignoradas = propuestas[marca.values], propuestas[~marca.values]
    col = {k: base.col(TAB_RECHAZOS, v) for k, v in C.items()}
    faltan = [C[k] for k in ("fecha", "folio", "nombre") if not col[k]]
    if faltan:
        raise ValueError(f"No encuentro las columnas {', '.join(faltan)} en '{TAB_RECHAZOS}'.")
    respaldo = base.respaldar(TAB_RECHAZOS) if len(elegidas) else None

    celdas, historial, nuevas_marcas = [], [], []
    for _, p in elegidas[elegidas["Tipo"] != "Nuevo rechazo"].iterrows():
        c = base.col(TAB_RECHAZOS, p["Campo"])
        if not c:
            continue
        celdas.append((int(p["Fila"]), c, p["_valor"]))
        historial.append(dict(Origen="App · bitácoras", Hoja=TAB_RECHAZOS, Celda=f"{letra(c)}{int(p['Fila'])}",
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
        nuevas_marcas.append((fila, col["nombre"]))
        contexto["fila_de_sol"].setdefault(id(p["_sol"]), []).append(fila)
        historial.append(dict(Origen="App · bitácoras", Hoja=TAB_RECHAZOS, Celda=f"Fila {fila}", Cliente=p["Cliente"],
                              Campo="(fila nueva)", Antes="", Después=p["Después"], Motivo=p["Motivo"]))
        fila += 1

    base.escribir(TAB_RECHAZOS, celdas)
    base.marcar(TAB_RECHAZOS, nuevas_marcas, AZUL_NUEVO)
    for _, p in ignoradas.iterrows():
        historial.append(dict(Origen="Ignorado", Hoja=TAB_RECHAZOS, Celda="", Cliente=p["Cliente"], Campo=p["Campo"],
                              Antes=p["Antes"], Después=p["Después"], Motivo="No se aplicó: " + p["Motivo"]))
    base.anotar_historial(historial)

    if capturas is not None and not capturas.empty:
        base.reemplazar(TAB_CAPTURAS, CAP_COLS + CAP_EXTRA,
                        capturas[CAP_COLS + CAP_EXTRA].fillna("").values.tolist())
    base.reemplazar(TAB_BITACORA, BITACORA_COLS, filas_bitacora(solicitudes, contexto["fila_de_sol"]))
    return dict(respaldo=respaldo, correcciones=int((elegidas["Tipo"] != "Nuevo rechazo").sum()),
                nuevos=len(nuevos), ignorados=len(ignoradas), celdas=len(celdas))
