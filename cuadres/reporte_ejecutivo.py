# -*- coding: utf-8 -*-
"""
Resumen ejecutivo de una página (PDF) con lo más importante del Sheet de seguimiento:

  · indicadores del año: rechazos, seguimiento al día, contacto efectivo, registro en Sale-U, vencidos, Revision
  · gráficas: rechazos por mes y financiera · bitácora por estatus
  · tablas: seguimientos vencidos más atrasados · asesores (rechazos, por registrar en Sale-U, vencidos)
  · «Prioridades de la semana»: frases armadas con los números

  pdf(base, anio=None, hoy=None) → bytes del PDF
Mismo estilo que el PDF «Clientes por registrar en Sale-U» (cuadres/reporte_saleu.py).
"""
import datetime as dt
import io

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (BaseDocTemplate, Frame, PageTemplate, Paragraph, Spacer, Table, TableStyle)

from cuadres import resumen
from cuadres.base import TAB_BITACORA, norm
from cuadres.fuentes import to_date
from cuadres.reporte_saleu import (Cifras, Imagen, NAVY, NAVY_2, TINTA, TINTA_2, TINTA_3, LINEA, MARGEN, UTIL,
                                   _corto, _e, _estilos, _fuentes, _icono, _ico, _mpl, _titulo)

ANCHO, ALTO = letter
SERIE = ["#2a78d6", "#eb6834", "#1baf7a"]                        # CETELEM, BBVA, otras
ESTADO = dict(bien="#0ca30c", aviso="#fab219", serio="#ec835a", critico="#d03b3b", neutro="#8a8984")
MESES = ["Ene", "Feb", "Mar", "Abr", "May", "Jun", "Jul", "Ago", "Sep", "Oct", "Nov", "Dic"]


def _pct(a, b):
    return f"{a / b:.0%}" if b else "—"


# =============================================================================
# Datos
# =============================================================================
def datos(base, anio=None, hoy=None):
    hoy = hoy or dt.date.today()
    anio = anio or hoy.year
    k = resumen.indicadores(base, anio)
    faltan, _ = resumen.faltantes_saleu(base, hoy=hoy)
    bit = pd.DataFrame()
    if TAB_BITACORA in base.pestanas():
        b = base.leer(TAB_BITACORA)
        if not b.empty and "Estatus final" in b.columns:
            b["_f"] = b["Fecha"].map(to_date)
            bit = b[b["_f"].map(lambda d: bool(d) and d.year == anio)].copy()
            bit["_m"] = bit["_f"].map(lambda d: d.month)
            bit["_e"] = bit["Estatus final"].map(lambda v: str(v or "").strip() or "Sin estatus")
    return dict(k=k, faltan=faltan, bit=bit, anio=anio, hoy=hoy)


# =============================================================================
# Gráficas
# =============================================================================
def _ejes(ax):
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom"].set_color(LINEA)
    ax.tick_params(axis="y", length=0, labelsize=7, colors=TINTA_3)
    ax.tick_params(axis="x", length=0, labelsize=7.5)
    ax.grid(axis="y", color=LINEA, linewidth=0.6)
    ax.set_axisbelow(True)


def grafica_rechazos(k, w, h):
    plt = _mpl()
    pm = k["por_mes"]
    ult = max([m for m in pm["mes"]] + [1]) if not pm.empty else 1
    meses = list(range(1, ult + 1))
    fig, ax = plt.subplots(figsize=(w / 72, h / 72))
    abajo = [0] * len(meses)
    for fin, color in zip(["CETELEM", "BBVA", "Otras"], SERIE):
        v = [int(pm[(pm["mes"] == m) & (pm["fin"] == fin)]["rechazos"].sum()) for m in meses]
        ax.bar(range(len(meses)), v, bottom=abajo, color=color, width=0.62, edgecolor="white", linewidth=1,
               label=fin)
        abajo = [a + b for a, b in zip(abajo, v)]
    for i, t in enumerate(abajo):
        if t:
            ax.text(i, t + 0.4, str(t), ha="center", va="bottom", fontsize=7, color=TINTA, fontweight="bold")
    ax.set_xticks(range(len(meses)), [MESES[m - 1] for m in meses])
    ax.set_ylim(0, max(abajo + [1]) * 1.18)
    _ejes(ax)
    return _titulo(fig, "Rechazos por mes", f"{k['rechazos']} en {k['anio']}, por financiera", ax)


def grafica_bitacora(bit, w, h):
    plt = _mpl()
    fig, ax = plt.subplots(figsize=(w / 72, h / 72))
    orden = [("Formalizado", SERIE[2]), ("Aprobado", SERIE[0]), ("Contrapropuesta", ESTADO["aviso"]),
             ("Análisis", ESTADO["neutro"]), ("Rechazado", ESTADO["critico"])]
    tot = bit["_e"].value_counts() if not bit.empty else pd.Series(dtype=int)
    n = int(tot.sum())
    nombres = [e for e, _ in orden if tot.get(e, 0)]
    valores = [int(tot.get(e, 0)) for e in nombres]
    col = [c for e, c in orden if tot.get(e, 0)]
    y = range(len(nombres))[::-1]
    ax.barh(list(y), valores, color=col, height=0.6)
    for yi, v in zip(y, valores):
        ax.text(v + max(valores + [1]) * 0.02, yi, f"{v}  ({v / n:.0%})", va="center", fontsize=7.5, color=TINTA)
    ax.set_yticks(list(y), nombres)
    ax.set_xlim(0, max(valores + [1]) * 1.32)
    for s in ("top", "right", "bottom"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color(LINEA)
    ax.xaxis.set_visible(False)
    ax.tick_params(axis="y", length=0, labelsize=7.8)
    return _titulo(fig, "Bitácora por estatus final", f"{n} solicitudes del año (todas las financieras)", ax)


# =============================================================================
# Tablas
# =============================================================================
def _tabla(enc, filas, anchos, der=()):
    st = _estilos()
    datos = [[Paragraph(_e(h), st["enc_d"] if i in der else st["enc"]) for i, h in enumerate(enc)]]
    for f in filas:
        datos.append([Paragraph(_e(v), st["celda_d"] if i in der else (st["nombre"] if i == 0 else st["celda"]))
                      for i, v in enumerate(f)])
    t = Table(datos, colWidths=anchos)
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor(LINEA)),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor(LINEA)),
        ("TOPPADDING", (0, 0), (-1, -1), 3), ("BOTTOMPADDING", (0, 0), (-1, -1), 3),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]))
    return t


def _bloque(titulo, icono, color, contenido, ancho):
    st = _estilos()
    t = Table([[Paragraph(f'{_ico(icono, color, 9.5)}  <b>{_e(titulo)}</b>', st["txt"])], [contenido]],
              colWidths=[ancho])
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("TOPPADDING", (0, 0), (-1, -1), 0), ("BOTTOMPADDING", (0, 0), (-1, 0), 4)]))
    return t


# =============================================================================
# PDF
# =============================================================================
class _Doc(BaseDocTemplate):
    def __init__(self, buf, anio, hoy):
        super().__init__(buf, pagesize=letter, leftMargin=MARGEN, rightMargin=MARGEN, topMargin=40, bottomMargin=34,
                         title=f"Resumen ejecutivo {anio}", author="MG Colima")
        self.anio, self.hoy = anio, hoy
        marco = Frame(MARGEN, 34, UTIL, ALTO - 34 - 92, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        self.addPageTemplates([PageTemplate("una", [marco], onPage=self._fondo)])

    def _fondo(self, c, doc):
        c.saveState()
        alto = 78
        c.setFillColor(colors.HexColor(NAVY))
        c.rect(0, ALTO - alto, ANCHO, alto, stroke=0, fill=1)
        c.setFillColor(colors.white)
        for (x, y, r, a) in ((ANCHO - 60, ALTO - 10, 70, 0.06), (ANCHO - 150, ALTO - 75, 40, 0.05)):
            c.setFillAlpha(a)
            c.circle(x, y, r, stroke=0, fill=1)
        c.setFillAlpha(1)
        c.setFillColor(colors.HexColor(NAVY_2))
        c.circle(MARGEN + 19, ALTO - 39, 19, stroke=0, fill=1)
        _icono(c, MARGEN + 19, ALTO - 44, "meta", 14, "#FFFFFF")
        c.setFillColor(colors.white)
        c.setFont("Archivo-ExtraBold", 19)
        c.drawString(MARGEN + 48, ALTO - 36, f"Resumen ejecutivo {self.anio} — MG Colima")
        c.setFont("Archivo", 9)
        c.setFillColor(colors.HexColor("#C9D6EA"))
        c.drawString(MARGEN + 48, ALTO - 52, f"Rechazos, seguimiento, Sale-U y bitácora · corte {self.hoy:%d/%m/%Y}")
        c.setFont("Archivo", 7)
        c.setFillColor(colors.HexColor(TINTA_3))
        c.drawString(MARGEN, 18, "Fuente: Google Sheet de seguimiento (MG Colima Rechazos, Bitacora y Revision)")
        c.restoreState()


def _prioridades(d):
    k, faltan = d["k"], d["faltan"]
    out = []
    if len(k["vencidos"]):
        peor = k["vencidos"].iloc[0]
        out.append(f"<b>{len(k['vencidos'])} seguimientos vencidos</b> (el más atrasado: {_e(str(peor['nombre']).title())}, "
                   f"{int(peor['dias'])} días). Otros {len(k['por_vencer'])} vencen en los próximos 15 días.")
    if not faltan.empty:
        ap = int((faltan["Estatus"] == resumen.APROBADO).sum())
        cred = int((faltan["Falta"] == resumen.FALTA_CREDITO).sum())
        out.append(f"<b>{len(faltan)} por registrar en Sale-U</b>: {len(faltan) - cred} dar de alta y {cred} registrar "
                   f"crédito" + (f"; {ap} ya están aprobados (la venta más cercana)." if ap else "."))
    if k["contactados"]:
        out.append(f"<b>Contacto efectivo {_pct(k['contestaron'], k['contactados'])}</b>: "
                   f"{k['sin_contesta']} contactos sin dato de si contestó; captúralo en CONTESTA.")
    if k["pendientes"]:
        out.append(f"<b>{k['pendientes']} pendientes en Revision</b> por resolver.")
    return out


def pdf(base, anio=None, hoy=None):
    """Bytes del PDF de una página."""
    _fuentes()
    st = _estilos()
    d = datos(base, anio, hoy)
    k, faltan, bit = d["k"], d["faltan"], d["bit"]
    buf = io.BytesIO()
    doc = _Doc(buf, d["anio"], d["hoy"])
    al_dia = k["al_dia"] / k["ya_toca"] if k["ya_toca"] else 1
    fin = " · ".join(f"{f} {n}" for f, n in sorted(k["por_fin"].items(), key=lambda x: -x[1]))
    el = [
        Cifras([("Rechazos en el año", k["rechazos"], fin, "rechazado", SERIE[1]),
                ("Seguimiento al día", _pct(k["al_dia"], k["ya_toca"]),
                 f"{k['al_dia']} de {k['ya_toca']} con 180 días cumplidos", "calendario",
                 ESTADO["bien"] if al_dia >= 0.9 else ESTADO["aviso"]),
                ("Contacto efectivo", _pct(k["contestaron"], k["contactados"]),
                 f"{k['contestaron']} de {k['contactados']} contestaron", "telefono", NAVY_2)], alto=70),
        Spacer(1, 8),
        Cifras([("Registrados en Sale-U", _pct(k["registrados"], k["saleu_total"]),
                 f"{k['falta_alta']} dar de alta · {k['falta_credito']} registrar crédito", "aprobado", SERIE[0]),
                ("Seguimientos vencidos", len(k["vencidos"]), f"{len(k['por_vencer'])} vencen en 15 días", "reloj",
                 ESTADO["critico"] if len(k["vencidos"]) else ESTADO["bien"]),
                ("Pendientes en Revision", k["pendientes"], "por resolver", "info",
                 ESTADO["aviso"] if k["pendientes"] else ESTADO["bien"])], alto=70),
        Spacer(1, 12)]
    w, h = (UTIL - 14) / 2, 168
    g = Table([[Imagen(grafica_rechazos(k, w, h), w, h), Imagen(grafica_bitacora(bit, w, h), w, h)]],
              colWidths=[w + 7, w + 7])
    g.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    el += [g, Spacer(1, 10)]

    venc = k["vencidos"].drop_duplicates("nombre").head(6)
    t_venc = _tabla(["Cliente", "Asesor", "Días"],
                    [[str(r["nombre"]).title(), _corto(r["asesor"]), str(int(r["dias"]))] for _, r in venc.iterrows()]
                    or [["Sin seguimientos vencidos", "", ""]], [w * 0.52, w * 0.36, w * 0.12], der=(2,))
    ases = k["por_asesor"].head(6)
    venc_as = k["vencidos"]["asesor"].replace("", "(sin asesor)").value_counts() if len(k["vencidos"]) else {}
    falt_as = faltan["Vendedor"].replace("", "(sin asesor)").value_counts() if not faltan.empty else {}
    t_as = _tabla(["Asesor", "Rechazos", "Por registrar", "Vencidos"],
                  [[_corto(a), str(int(n)), str(int(falt_as.get(a, 0))), str(int(venc_as.get(a, 0)))]
                   for a, n in ases.items()], [w * 0.43, w * 0.19, w * 0.22, w * 0.16], der=(1, 2, 3))
    t = Table([[_bloque(f"Seguimientos más atrasados ({len(k['vencidos'])} vencidos)", "reloj", ESTADO["critico"],
                        t_venc, w),
                _bloque("Asesores con más rechazos", "asesores", NAVY_2, t_as, w)]], colWidths=[w + 14, w])
    t.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                           ("VALIGN", (0, 0), (-1, -1), "TOP")]))
    el += [t, Spacer(1, 12)]
    prio = _prioridades(d)
    if prio:
        filas = [[Paragraph(f'{_ico("idea", NAVY_2, 10)}  <b>Prioridades</b>', st["txt"])]]
        filas += [[Paragraph(f'<font color="{NAVY_2}">•</font>  {p}', st["txt"])] for p in prio]
        caja = Table(filas, colWidths=[UTIL])
        caja.setStyle(TableStyle([("BACKGROUND", (0, 0), (-1, -1), colors.HexColor("#EEF3FA")),
                                  ("LINEBEFORE", (0, 0), (0, -1), 3, colors.HexColor(NAVY_2)),
                                  ("LEFTPADDING", (0, 0), (-1, -1), 12), ("TOPPADDING", (0, 0), (-1, -1), 2.5),
                                  ("BOTTOMPADDING", (0, 0), (-1, -1), 2.5), ("TOPPADDING", (0, 0), (-1, 0), 7),
                                  ("BOTTOMPADDING", (0, -1), (-1, -1), 7)]))
        el.append(caja)
    doc.build(el)
    return buf.getvalue()
