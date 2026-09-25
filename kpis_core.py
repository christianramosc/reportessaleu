# -*- coding: utf-8 -*-
"""
kpis_core.py  (v2)
KPIs de capacitación Ventas / Posventa de MG Colima a partir del archivo PLAN DE ACCIÓN.

- Lee las hojas "KPIS-VENTAS" y "KPIS POSVENTA" (las tablas se localizan por su título).
- Recalcula totales y avances desde las filas de cada colaborador.
- Revisa la calidad de los datos (totales del Excel, letras, sumas) y devuelve alertas.
- Genera: PDF (resumen + detalle por área), una imagen PNG del resumen por área
  y un Excel con el detalle por colaborador.

Uso en Colab:  !python kpis_core.py   (pide el archivo si no existe y descarga todo)
Uso en Streamlit: import kpis_core as kc; kc.generar(bytes_del_excel)
"""
import os
import re
import zipfile
import unicodedata
from io import BytesIO
from datetime import datetime

import openpyxl
import pandas as pd
import matplotlib
matplotlib.use("Agg")
import matplotlib.pyplot as plt
from matplotlib.backends.backend_pdf import PdfPages
from matplotlib.patches import FancyBboxPatch, Circle, Wedge, Rectangle, Polygon

# ---------------------------------------------------------------- CONFIG
ARCHIVO = "PLAN_DE_ACCIÓN.xlsx"
SALIDA = "KPIs_Capacitacion_MG_Colima.pdf"
HOJAS = {"VENTAS": "KPIS-VENTAS", "POSVENTA": "KPIS POSVENTA"}

MODULOS = {  # clave -> (nombre, color, color claro, icono)
    "DOT":        ("DOT",        "#1E6FD9", "#EAF3FD", "grupo"),
    "CAPSULAS":   ("Cápsulas",   "#138A4B", "#E9F6EE", "pantalla"),
    "LOL":        ("LOL",        "#5B3FA8", "#F1EDFA", "grupo"),
    "PRESENCIAL": ("Presencial", "#E8820C", "#FDF3E7", "presentador"),
}
ORDEN = ["DOT", "CAPSULAS", "LOL", "PRESENCIAL"]
AZUL_OSCURO = "#12305C"
GRIS = "#D5DBE3"
TEXTO = "#1B2A41"
TEXTO_2 = "#5A6B82"

# semáforo de avance
META_VERDE, META_AMBAR = 0.80, 0.50
VERDE, AMBAR, ROJO = "#2E9E5B", "#E0A21B", "#D64545"
VERDE_CL, AMBAR_CL, ROJO_CL = "#E3F4EA", "#FCF1D8", "#FBE3E3"

FIG_W, FIG_H = 11, 8.5
plt.rcParams["font.family"] = "DejaVu Sans"


def semaforo(v, claro=False):
    if v >= META_VERDE:
        return VERDE_CL if claro else VERDE
    if v >= META_AMBAR:
        return AMBAR_CL if claro else AMBAR
    return ROJO_CL if claro else ROJO


def ahora():
    try:
        from zoneinfo import ZoneInfo
        return datetime.now(ZoneInfo("America/Mexico_City"))
    except Exception:  # noqa: BLE001
        return datetime.now()


def pct(v):
    return f"{v * 100:.0f}%"


def entero(v):
    return f"{v:,.0f}"


# ---------------------------------------------------------------- LECTURA
def _norm(s):
    s = unicodedata.normalize("NFKD", str(s)).encode("ascii", "ignore").decode()
    return s.upper().strip()


def _num(v):
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    try:
        return float(str(v).strip().replace(",", "."))
    except ValueError:
        return 0.0  # 'N/A', vacío, etc.


def leer_hoja(ws):
    """Devuelve ({modulo: [filas]}, {modulo: totales_del_excel})."""
    tablas, totales_excel = {}, {}
    for row in ws.iter_rows():
        for c in row:
            if not isinstance(c.value, str):
                continue
            m = re.search(r"\((CAPSULAS|DOT|LOL|PRESENCIAL)\)", _norm(c.value))
            if not m or m.group(1) in tablas:
                continue
            clave, col, fila = m.group(1), c.column, c.row + 2  # título, encabezado, datos
            filas = []
            while fila <= ws.max_row:
                nombre = ws.cell(fila, col).value
                if nombre is None:
                    break
                if "TOTAL" in _norm(nombre):
                    totales_excel[clave] = {
                        "total": _num(ws.cell(fila, col + 2).value),
                        "concluidas": _num(ws.cell(fila, col + 3).value),
                        "avance": _num(ws.cell(fila, col + 6).value),
                    }
                    break
                total = _num(ws.cell(fila, col + 2).value)
                concl = _num(ws.cell(fila, col + 3).value)
                filas.append({
                    "nombre": " ".join(str(nombre).split()),
                    "letra": str(ws.cell(fila, col + 1).value or "").strip().upper(),
                    "total": total,
                    "concluidas": concl,
                    "pendientes": _num(ws.cell(fila, col + 4).value),
                    "noshow": _num(ws.cell(fila, col + 5).value),
                    "avance": min(concl / total, 1.0) if total else 0.0,
                })
                fila += 1
            tablas[clave] = filas
    return tablas, totales_excel


def resumir(tablas):
    res = {}
    for k in ORDEN:
        f = tablas.get(k, [])
        n = len(f)
        res[k] = {
            "total": sum(x["total"] for x in f),
            "concluidas": sum(x["concluidas"] for x in f),
            "pendientes": sum(x["pendientes"] for x in f),
            "noshow": sum(x["noshow"] for x in f),
            "avance": (sum(x["avance"] for x in f) / n) if n else 0.0,
        }
    letras = {}
    for k in ORDEN:
        for x in tablas.get(k, []):
            d = letras.setdefault(x["letra"], {"av": [], "total": 0, "personas": set()})
            d["av"].append(x["avance"])
            d["total"] += x["total"]
            d["personas"].add(x["nombre"])
    por_letra = {l: {"avance": sum(d["av"]) / len(d["av"]), "total": d["total"],
                     "personas": len(d["personas"])}
                 for l, d in letras.items() if l}
    general = sum(res[k]["avance"] for k in ORDEN) / len(ORDEN)
    return res, por_letra, general


def detalle_df(tablas):
    """Tabla ancha: una fila por colaborador, avance y concluidas/total por módulo."""
    nombres = []
    for k in ORDEN:
        for x in tablas.get(k, []):
            if x["nombre"] not in nombres:
                nombres.append(x["nombre"])
    filas = []
    for nm in nombres:
        f = {"Colaborador": nm, "Letra": ""}
        avs = []
        for k in ORDEN:
            x = next((r for r in tablas.get(k, []) if r["nombre"] == nm), None)
            nom = MODULOS[k][0]
            if x:
                f["Letra"] = f["Letra"] or x["letra"]
                f[nom] = x["avance"]
                f[f"{nom} (concl/total)"] = f"{x['concluidas']:.0f}/{x['total']:.0f}"
                avs.append(x["avance"])
            else:
                f[nom] = None
                f[f"{nom} (concl/total)"] = ""
        f["Promedio"] = sum(avs) / len(avs) if avs else 0.0
        filas.append(f)
    return pd.DataFrame(filas)


def validar(area, tablas, totales_excel):
    """Lista de alertas (texto) sobre la calidad de los datos."""
    alertas = []
    # 1) totales del Excel vs recalculados
    for k in ORDEN:
        f, ex = tablas.get(k, []), totales_excel.get(k)
        if not ex:
            continue
        nom = MODULOS[k][0]
        tot = sum(x["total"] for x in f)
        con = sum(x["concluidas"] for x in f)
        av = sum(x["avance"] for x in f) / len(f) if f else 0
        if abs(ex["total"] - tot) > 0.5:
            alertas.append(f"{nom}: el Excel suma {entero(ex['total'])} asignaciones, "
                           f"pero las filas suman {entero(tot)}. Revisa el rango de la fórmula.")
        if abs(ex["concluidas"] - con) > 0.5:
            alertas.append(f"{nom}: el Excel suma {entero(ex['concluidas'])} concluidas, "
                           f"pero las filas suman {entero(con)}.")
        if abs(ex["avance"] - av) > 0.005:
            extra = " (mayor a 100%)" if ex["avance"] > 1 else ""
            alertas.append(f"{nom}: el Excel marca {pct(ex['avance'])} de avance{extra}; "
                           f"el promedio real de {len(f)} personas es {pct(av)}. "
                           "Revisa el divisor de la fórmula.")
    # 2) letra distinta para la misma persona
    letras = {}
    for k in ORDEN:
        for x in tablas.get(k, []):
            letras.setdefault(x["nombre"], {}).setdefault(x["letra"], []).append(MODULOS[k][0])
    for nm, d in letras.items():
        if len(d) > 1:
            det = "; ".join(f"{l} en {', '.join(m)}" for l, m in d.items())
            alertas.append(f"{nm} tiene letras distintas: {det}.")
    # 3) persona que no aparece en todos los módulos
    for nm in letras:
        faltan = [MODULOS[k][0] for k in ORDEN if all(x["nombre"] != nm for x in tablas.get(k, []))]
        if faltan:
            alertas.append(f"{nm} no aparece en: {', '.join(faltan)}.")
    # 4) filas inconsistentes
    for k in ORDEN:
        for x in tablas.get(k, []):
            nom = MODULOS[k][0]
            if x["concluidas"] > x["total"]:
                alertas.append(f"{nom} / {x['nombre']}: concluidas ({entero(x['concluidas'])}) "
                               f"mayor que el total ({entero(x['total'])}).")
            elif abs(x["concluidas"] + x["pendientes"] - x["total"]) > 0.5:
                alertas.append(f"{nom} / {x['nombre']}: concluidas + pendientes = "
                               f"{entero(x['concluidas'] + x['pendientes'])}, pero el total es "
                               f"{entero(x['total'])}.")
    return alertas


# ---------------------------------------------------------------- DIBUJO
def caja(ax, x, y, w, h, fc, ec, r=0.012, lw=1.0):
    ax.add_patch(FancyBboxPatch((x, y), w, h, boxstyle=f"round,pad=0,rounding_size={r}",
                                fc=fc, ec=ec, lw=lw, transform=ax.transAxes, zorder=1))


def dona(fig, rect, pct, color, texto_centro, sub="Avance", grosor=0.22, fs=15):
    a = fig.add_axes(rect)
    a.set_aspect("equal")
    a.axis("off")
    a.set_xlim(-1.05, 1.05)
    a.set_ylim(-1.05, 1.05)
    a.add_patch(Wedge((0, 0), 1, 0, 360, width=grosor, fc=GRIS, ec="none"))
    p = max(0.0, min(pct, 1.0))
    if p > 0:
        a.add_patch(Wedge((0, 0), 1, 90 - 360 * p, 90, width=grosor, fc=color, ec="none"))
    a.text(0, 0.1, texto_centro, ha="center", va="center", fontsize=fs, weight="bold", color=TEXTO)
    a.text(0, -0.28, sub, ha="center", va="center", fontsize=fs * 0.45, color=TEXTO_2)
    return a


def icono(fig, rect, color, tipo):
    a = fig.add_axes(rect)
    a.set_aspect("equal")
    a.axis("off")
    a.set_xlim(-1, 1)
    a.set_ylim(-1, 1)
    a.add_patch(Circle((0, 0), 1, fc=color, ec="none"))
    w = "white"
    if tipo == "grupo":
        for dx, s in [(-0.38, 0.8), (0.38, 0.8), (0, 1.0)]:
            a.add_patch(Circle((dx, 0.18 * s + 0.02), 0.17 * s, fc=w, ec=color, lw=1.2))
            a.add_patch(Wedge((dx, -0.42), 0.34 * s, 0, 180, fc=w, ec=color, lw=1.2))
    elif tipo == "pantalla":
        a.add_patch(Rectangle((-0.55, -0.3), 1.1, 0.75, fc="none", ec=w, lw=2.2))
        a.add_patch(Rectangle((-0.3, -0.5), 0.6, 0.07, fc=w, ec="none"))
        a.add_patch(Polygon([(-0.12, -0.1), (-0.12, 0.25), (0.2, 0.07)], fc=w, ec="none"))
    else:  # presentador
        a.add_patch(Rectangle((-0.1, -0.05), 0.65, 0.5, fc="none", ec=w, lw=2))
        a.add_patch(Circle((-0.38, 0.18), 0.17, fc=w, ec="none"))
        a.add_patch(Wedge((-0.38, -0.5), 0.33, 0, 180, fc=w, ec="none"))
        a.plot([0.05, 0.4], [0.05, 0.3], color=w, lw=2)



def _lienzo():
    fig = plt.figure(figsize=(FIG_W, FIG_H))
    ax = fig.add_axes([0, 0, 1, 1])
    ax.axis("off")
    ax.set_xlim(0, 1)
    ax.set_ylim(0, 1)
    ax.set_autoscale_on(False)
    return fig, ax


def _encabezado(ax, titulo, derecha):
    ax.add_patch(Rectangle((0, 0.915), 1, 0.085, fc=AZUL_OSCURO, ec="none"))
    ax.text(0.035, 0.957, titulo, color="white", fontsize=16, weight="bold", va="center")
    lineas = derecha.split("\n")
    for i, t in enumerate(lineas):
        yy = 0.957 + (len(lineas) - 1) * 0.0125 - i * 0.025
        ax.text(0.965, yy, t, color="#C9D6EA", fontsize=9.5, va="center", ha="right")


def _panel(ax, x, y, w, h, titulo):
    caja(ax, x, y, w, h, "white", GRIS)
    ax.add_patch(Rectangle((x, y + h - 0.065), w, 0.065, fc=AZUL_OSCURO, ec="none"))
    ax.text(x + 0.02, y + h - 0.0325, titulo, color="white", fontsize=12, weight="bold", va="center")


def pagina_resumen(area, res, por_letra, general, n_colab):
    fig, ax = _lienzo()
    asign = sum(res[k]["total"] for k in ORDEN)
    concl = sum(res[k]["concluidas"] for k in ORDEN)
    _encabezado(ax, f"AVANCE DE CAPACITACIÓN {area} · MG COLIMA",
                f"{n_colab} colaboradores  |  {entero(concl)} de {entero(asign)} concluidas\n"
                f"Generado el {ahora():%d/%m/%Y}")

    # ---------- tarjetas de módulo
    mx, gap, y0, h = 0.03, 0.015, 0.545, 0.345
    w = (1 - 2 * mx - 3 * gap) / 4
    for i, k in enumerate(ORDEN):
        nombre, col, claro, tipo = MODULOS[k]
        r = res[k]
        x = mx + i * (w + gap)
        caja(ax, x, y0, w, h, claro, col + "55")
        ax.add_patch(Rectangle((x + 0.002, y0 + 0.002), w - 0.004, 0.05, fc=col + "22", ec="none"))
        icono(fig, [x + 0.01, y0 + h - 0.085, 0.07, 0.075], col, tipo)
        ax.text(x + 0.085, y0 + h - 0.0475, nombre, fontsize=15, weight="bold", color=TEXTO, va="center")
        dona(fig, [x + w / 2 - 0.07, y0 + 0.118, 0.14, 0.132], r["avance"], col, pct(r["avance"]), fs=15)
        stats = [(r["concluidas"], "Concluidas"), (r["pendientes"], "Pendientes"), (r["noshow"], "No show")]
        for j, (v, lbl) in enumerate(stats):
            cx = x + w * (j + 0.5) / 3
            ax.text(cx, y0 + 0.083, entero(v), fontsize=14, weight="bold", color=TEXTO, ha="center")
            ax.text(cx, y0 + 0.062, lbl, fontsize=7.5, color=TEXTO_2, ha="center")
            if j < 2:
                dx = x + w * (j + 1) / 3
                ax.plot([dx, dx], [y0 + 0.06, y0 + 0.105], color=col + "55", lw=1)
        ax.text(x + w / 2, y0 + 0.026, f"Total de {nombre}: {entero(r['total'])}",
                fontsize=10.5, weight="bold", color=TEXTO, ha="center", va="center")

    # ---------- avance por letra (barra con semáforo)
    lx, ly, lw_, lh = 0.03, 0.06, 0.455, 0.46
    _panel(ax, lx, ly, lw_, lh, "AVANCE GENERAL POR LETRA")
    orden_l = sorted(por_letra, key=lambda l: (l != "P", l))
    n = max(len(orden_l), 1)
    cw = (lw_ - 0.04) / n
    for i, l in enumerate(orden_l):
        d = por_letra[l]
        cx0 = lx + 0.02 + i * cw
        caja(ax, cx0 + 0.005, ly + 0.03, cw - 0.01, lh - 0.12, "#F6F8FB", "#E3E8EF")
        mid = cx0 + cw / 2
        ax.text(mid, ly + 0.3, l, fontsize=16, weight="bold", color=TEXTO, ha="center")
        ax.text(mid, ly + 0.21, pct(d["avance"]), fontsize=22, weight="bold", color=TEXTO, ha="center")
        bw, bx, by = cw - 0.05, cx0 + 0.025, ly + 0.155
        caja(ax, bx, by, bw, 0.026, GRIS, "none", r=0.01)
        if d["avance"] > 0:
            caja(ax, bx, by, max(bw * min(d["avance"], 1), 0.012), 0.026, semaforo(d["avance"]),
                 "none", r=0.01)
        ax.text(mid, ly + 0.098, f"{d['personas']} personas", fontsize=9.5, color=TEXTO_2, ha="center")
        ax.text(mid, ly + 0.066, f"{entero(d['total'])} asignaciones", fontsize=9.5,
                color=TEXTO_2, ha="center")

    # ---------- avance promedio general (dona con semáforo)
    gx, gy, gw, gh = 0.515, 0.06, 0.455, 0.46
    _panel(ax, gx, gy, gw, gh, "AVANCE PROMEDIO GENERAL")
    dona(fig, [gx + 0.01, gy + 0.05, 0.2, 0.28], general, semaforo(general), pct(general),
         sub="Avance total", grosor=0.2, fs=24)
    lx2, ly2 = gx + 0.235, gy + 0.05
    caja(ax, lx2, ly2, 0.2, 0.28, "#F6F8FB", "#E3E8EF")
    for i, k in enumerate(ORDEN):
        nombre, col, _, _ = MODULOS[k]
        yy = ly2 + 0.245 - i * 0.065
        ax.add_patch(Circle((lx2 + 0.022, yy), 0.008, fc=col, ec="none"))
        ax.text(lx2 + 0.04, yy, nombre, fontsize=11, color=TEXTO, va="center")
        ax.text(lx2 + 0.185, yy, pct(res[k]["avance"]), fontsize=11, weight="bold", color=TEXTO,
                va="center", ha="right")
        if i < 3:
            ax.plot([lx2 + 0.012, lx2 + 0.188], [yy - 0.0325] * 2, color="#E3E8EF", lw=1)

    # pie: método + leyenda de semáforo
    ax.text(0.03, 0.025, "Avance por módulo = promedio del avance individual (concluidas / total). "
            "Avance total = promedio de los 4 módulos.", fontsize=7.5, color=TEXTO_2, va="center")
    xs = 0.62
    for c, t in [(VERDE, f"≥ {META_VERDE:.0%}"), (AMBAR, f"{META_AMBAR:.0%} a {META_VERDE:.0%}"),
                 (ROJO, f"< {META_AMBAR:.0%}")]:
        ax.add_patch(Circle((xs, 0.025), 0.006, fc=c, ec="none"))
        ax.text(xs + 0.01, 0.025, t, fontsize=7.5, color=TEXTO_2, va="center")
        xs += 0.085
    return fig


def pagina_detalle(area, df):
    """Tabla tipo mapa de calor: colaborador x módulo, con concluidas/total y avance."""
    fig, ax = _lienzo()
    _encabezado(ax, f"DETALLE POR COLABORADOR · {area}", f"{len(df)} colaboradores")
    df = df.sort_values(["Letra", "Promedio"], ascending=[True, False]).reset_index(drop=True)

    cols = [("Colaborador", 0.30), ("Letra", 0.06)] + [(MODULOS[k][0], 0.115) for k in ORDEN] + \
           [("Promedio", 0.12)]
    x0, ytop, ybot = 0.03, 0.87, 0.07
    alto = min(0.06, (ytop - ybot) / (len(df) + 1))
    # encabezado de tabla
    x = x0
    for nom, w in cols:
        ax.add_patch(Rectangle((x, ytop - alto), w, alto, fc="#E8EDF4", ec="white", lw=2))
        ax.text(x + (0.01 if nom == "Colaborador" else w / 2), ytop - alto / 2, nom, fontsize=10,
                weight="bold", color=TEXTO, va="center", ha="left" if nom == "Colaborador" else "center")
        x += w
    fs = 9.5 if alto > 0.045 else 8.5
    for i, r in df.iterrows():
        y = ytop - alto * (i + 2)
        x = x0
        for nom, w in cols:
            if nom == "Colaborador":
                ax.add_patch(Rectangle((x, y), w, alto, fc="#F6F8FB" if i % 2 else "white", ec="white", lw=2))
                ax.text(x + 0.01, y + alto / 2, str(r[nom]).title(), fontsize=fs, color=TEXTO, va="center")
            elif nom == "Letra":
                ax.add_patch(Rectangle((x, y), w, alto, fc="#F6F8FB" if i % 2 else "white", ec="white", lw=2))
                ax.text(x + w / 2, y + alto / 2, r[nom], fontsize=fs, weight="bold", color=TEXTO,
                        va="center", ha="center")
            else:
                v = r[nom]
                if v is None or pd.isna(v):
                    ax.add_patch(Rectangle((x, y), w, alto, fc="#F0F0F0", ec="white", lw=2))
                    ax.text(x + w / 2, y + alto / 2, "—", fontsize=fs, color=TEXTO_2, va="center", ha="center")
                else:
                    ax.add_patch(Rectangle((x, y), w, alto, fc=semaforo(v, claro=True), ec="white", lw=2))
                    sub = r.get(f"{nom} (concl/total)", "")
                    txt = pct(v) if nom == "Promedio" else f"{pct(v)}  ({sub})"
                    ax.text(x + w / 2, y + alto / 2, txt, fontsize=fs,
                            weight="bold" if nom == "Promedio" else "normal",
                            color=semaforo(v) if nom == "Promedio" else TEXTO, va="center", ha="center")
            x += w
    ax.text(0.03, 0.035, "Celdas: avance (concluidas/total). Promedio = promedio de los módulos "
            "asignados. Colores según semáforo de avance.", fontsize=7.5, color=TEXTO_2)
    return fig


# ---------------------------------------------------------------- EXPORTACIÓN
def _png(fig, dpi=200):
    buf = BytesIO()
    fig.savefig(buf, format="png", dpi=dpi, facecolor="white")
    return buf.getvalue()


def generar(archivo):
    """archivo: ruta, bytes o BytesIO del Excel.
    Devuelve dict con pdf, imagenes {nombre: png}, excel (detalle), zip, areas {area: datos}."""
    if isinstance(archivo, (bytes, bytearray)):
        archivo = BytesIO(archivo)
    wb = openpyxl.load_workbook(archivo, data_only=True)
    areas, imagenes, figs = {}, {}, []
    for area, hoja in HOJAS.items():
        if hoja not in wb.sheetnames:
            raise ValueError(f"No se encontró la hoja '{hoja}'. Hojas del archivo: "
                             f"{', '.join(wb.sheetnames)}.")
        tablas, tot_excel = leer_hoja(wb[hoja])
        faltan = [MODULOS[k][0] for k in ORDEN if not tablas.get(k)]
        if faltan:
            raise ValueError(f"En la hoja '{hoja}' no se encontraron las tablas de: {', '.join(faltan)}.")
        res, por_letra, general = resumir(tablas)
        df = detalle_df(tablas)
        f1 = pagina_resumen(area, res, por_letra, general, len(df))
        f2 = pagina_detalle(area, df)
        nombre_png = f"KPIs_{area.title()}_resumen.png"
        imagenes[nombre_png] = _png(f1)
        figs += [f1, f2]
        areas[area] = {"modulos": res, "por_letra": por_letra, "general": general,
                       "colaboradores": len(df), "detalle": df,
                       "alertas": validar(area, tablas, tot_excel), "png": nombre_png}

    buf_pdf = BytesIO()
    with PdfPages(buf_pdf) as pdf:
        for i, f in enumerate(figs, 1):
            f.text(0.965, 0.025, f"Página {i} de {len(figs)}", fontsize=7.5, color=TEXTO_2,
                   ha="right", va="center")
            pdf.savefig(f)
            plt.close(f)

    buf_xl = BytesIO()
    with pd.ExcelWriter(buf_xl, engine="openpyxl") as xw:
        for area, d in areas.items():
            d["detalle"].to_excel(xw, sheet_name=area.title(), index=False)
            resumen = pd.DataFrame([{"Módulo": MODULOS[k][0], **{c.title(): v for c, v in d["modulos"][k].items()}}
                                    for k in ORDEN])
            resumen.to_excel(xw, sheet_name=f"Resumen {area.title()}", index=False)

    buf_zip = BytesIO()
    with zipfile.ZipFile(buf_zip, "w", zipfile.ZIP_DEFLATED) as z:
        for nombre, data in imagenes.items():
            z.writestr(nombre, data)
        z.writestr(SALIDA, buf_pdf.getvalue())
        z.writestr("KPIs_Capacitacion_detalle.xlsx", buf_xl.getvalue())
    return {"pdf": buf_pdf.getvalue(), "imagenes": imagenes, "excel": buf_xl.getvalue(),
            "zip": buf_zip.getvalue(), "areas": areas}


def main(archivo=ARCHIVO, carpeta="."):
    out = generar(archivo)
    os.makedirs(carpeta, exist_ok=True)
    rutas = []
    for nombre, data in [(SALIDA, out["pdf"]), ("KPIs_Capacitacion_detalle.xlsx", out["excel"]),
                         *out["imagenes"].items()]:
        ruta = os.path.join(carpeta, nombre)
        with open(ruta, "wb") as f:
            f.write(data)
        rutas.append(ruta)
    for area, d in out["areas"].items():
        print(f"\n{area}: avance total {pct(d['general'])}")
        for k in ORDEN:
            r = d["modulos"][k]
            print(f"  {MODULOS[k][0]:<11} {pct(r['avance']):>5}  concl {r['concluidas']:.0f}  "
                  f"pend {r['pendientes']:.0f}  no show {r['noshow']:.0f}  total {r['total']:.0f}")
        if d["alertas"]:
            print(f"  Alertas ({len(d['alertas'])}):")
            for a in d["alertas"]:
                print(f"   - {a}")
    print("\nGenerados:", *rutas, sep="\n  ")
    return rutas


if __name__ == "__main__":
    try:  # Google Colab: pide el archivo si no existe y descarga los resultados
        from google.colab import files
        if not os.path.exists(ARCHIVO):
            ARCHIVO = next(iter(files.upload()))
        for r in main(ARCHIVO):
            files.download(r)
    except ImportError:
        main()
