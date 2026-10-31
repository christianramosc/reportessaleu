# -*- coding: utf-8 -*-
"""
Reporte PDF «Clientes por registrar en Sale-U»: Aprobados y Rechazados que faltan en Sale-U según el Sheet
(ver resumen.faltantes_saleu).

  pdf(faltan, por_confirmar) → bytes del PDF

Portada con cifras (cuántos falta dar de alta y a cuántos les falta el crédito), gráficas por asesor y de antigüedad, y después el detalle en dos partes —«Dar de alta en Sale-U» y
«Registrar el crédito»—, cada una por asesor (Aprobados primero), y los aprobados por confirmar. Letra Archivo e iconos de Font Awesome
Free (assets/fonts, licencias SIL OFL y CC BY 4.0).
"""
import datetime as dt
import html
import io
import os

import pandas as pd
from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import ParagraphStyle
from reportlab.lib.utils import ImageReader
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.platypus import (BaseDocTemplate, CondPageBreak, Flowable, Frame, KeepTogether, PageTemplate,
                                Paragraph, Spacer, Table, TableStyle, NextPageTemplate, PageBreak)

from cuadres.resumen import APROBADO, RECHAZADO, FALTA_ALTA, FALTA_CREDITO

FUENTES = os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), "assets", "fonts")
ANCHO, ALTO = letter
MARGEN = 40
UTIL = ANCHO - 2 * MARGEN

# Colores: texto en tintas neutras; azul y naranja (slots 1 y 2 de la paleta categórica, validados) para el estatus
NAVY, NAVY_2 = "#12305C", "#1D4A86"
TINTA, TINTA_2, TINTA_3 = "#1B2A41", "#5B6B82", "#8A96A8"
SUPERFICIE, LINEA = "#F5F7FB", "#E3E8EF"
COLOR = {APROBADO: "#2a78d6", RECHAZADO: "#eb6834", "confirmar": "#8a8984"}
FALTA = {FALTA_ALTA: ("#d03b3b", "usuario_mas", "Dar de alta en Sale-U",
                      "No aparecen en Sale-U: ni en el reporte Fuente ni en el de Crédito. Dalos de alta como lead y "
                      "registra su crédito."),
         FALTA_CREDITO: ("#C98A00", "dinero", "Registrar el crédito en Sale-U",
                         "Ya son lead en Sale-U (salen en el reporte Fuente), pero su crédito no está registrado: "
                         "captura la solicitud en el lead que ya existe, sin darlos de alta otra vez.")}
TINTE = {APROBADO: "#E3EEFB", RECHAZADO: "#FDEBE3", "confirmar": "#EFEFEC"}

ICONO = dict(auto="", usuario_mas="", aprobado="", rechazado="", asesores="",
             telefono="", calendario="", reloj="", espera="", idea="",
             asesor="", persona="", correo="", dinero="", info="", meta="")


# =============================================================================
# Fuentes
# =============================================================================
_LISTO = {}


def _fuentes():
    if _LISTO:
        return
    for nombre, archivo in [("Archivo", "Archivo-Regular"), ("Archivo-Medium", "Archivo-Medium"),
                            ("Archivo-SemiBold", "Archivo-SemiBold"), ("Archivo-Bold", "Archivo-Bold"),
                            ("Archivo-ExtraBold", "Archivo-ExtraBold"), ("FA", "fa-solid-900")]:
        pdfmetrics.registerFont(TTFont(nombre, os.path.join(FUENTES, archivo + ".ttf")))
    pdfmetrics.registerFontFamily("Archivo", normal="Archivo", bold="Archivo-Bold", italic="Archivo",
                                  boldItalic="Archivo-Bold")
    from matplotlib import font_manager
    for archivo in ("Archivo-Regular", "Archivo-SemiBold", "Archivo-Bold"):
        font_manager.fontManager.addfont(os.path.join(FUENTES, archivo + ".ttf"))
    _LISTO["ok"] = True


def _e(t):
    return html.escape(str(t if t is not None else ""), quote=False)


def _ico(nombre, color=TINTA_3, tam=None):
    t = f' size="{tam}"' if tam else ""
    return f'<font name="FA" color="{color}"{t}>{ICONO[nombre]}</font>'


ESTILO = {}


def _estilos():
    if ESTILO:
        return ESTILO
    base = dict(fontName="Archivo", textColor=colors.HexColor(TINTA))
    ESTILO.update(
        h2=ParagraphStyle("h2", fontName="Archivo-ExtraBold", fontSize=15, leading=19, textColor=colors.HexColor(NAVY),
                          spaceBefore=4, spaceAfter=6),
        txt=ParagraphStyle("txt", fontSize=9, leading=12.5, **base),
        sub=ParagraphStyle("sub", fontSize=8, leading=10.5, fontName="Archivo", textColor=colors.HexColor(TINTA_2)),
        nombre=ParagraphStyle("nombre", fontName="Archivo-SemiBold", fontSize=8.8, leading=11,
                              textColor=colors.HexColor(TINTA)),
        celda=ParagraphStyle("celda", fontSize=8.3, leading=10.5, **base),
        celda_d=ParagraphStyle("celda_d", fontSize=8.3, leading=10.5, alignment=2, **base),
        enc=ParagraphStyle("enc", fontName="Archivo-SemiBold", fontSize=7.5, leading=9,
                           textColor=colors.HexColor(TINTA_2)),
        enc_d=ParagraphStyle("enc_d", fontName="Archivo-SemiBold", fontSize=7.5, leading=9, alignment=2,
                             textColor=colors.HexColor(TINTA_2)),
    )
    return ESTILO


def _fecha(d):
    return d.strftime("%d/%m/%Y") if isinstance(d, (dt.date, dt.datetime)) else "—"


def _corto(nombre):
    p = [x for x in str(nombre).split() if x.upper() not in ("DE", "DEL", "LA", "LAS", "LOS", "Y")]
    if not p:
        return "(sin asesor)"
    return " ".join([p[0], p[-2]]).title() if len(p) >= 4 else " ".join(p[:2]).title()


# =============================================================================
# Piezas dibujadas
# =============================================================================
def _redondeado(c, x, y, w, h, r, relleno, borde=None):
    c.setFillColor(colors.HexColor(relleno))
    if borde:
        c.setStrokeColor(colors.HexColor(borde))
        c.setLineWidth(0.6)
    c.roundRect(x, y, w, h, r, stroke=1 if borde else 0, fill=1)


def _icono(c, x, y, nombre, tam, color):
    c.setFillColor(colors.HexColor(color))
    c.setFont("FA", tam)
    c.drawCentredString(x, y, ICONO[nombre])


class Cifras(Flowable):
    """Fila de tarjetas: icono en círculo con tinte, número grande, nombre y detalle."""

    def __init__(self, items, alto=78):
        super().__init__()
        self.items, self.alto = items, alto

    def wrap(self, *a):
        return UTIL, self.alto

    def draw(self):
        c = self.canv
        n, sep = len(self.items), 10
        w = (UTIL - sep * (n - 1)) / n
        for i, (nombre, valor, detalle, icono, color) in enumerate(self.items):
            x = i * (w + sep)
            _redondeado(c, x, 0, w, self.alto, 9, SUPERFICIE)
            c.setFillColor(colors.HexColor(color))
            c.roundRect(x, self.alto - 4, w, 4, 2, stroke=0, fill=1)
            c.saveState()
            c.setFillColor(colors.HexColor(color))
            c.setFillAlpha(0.14)
            c.circle(x + 22, self.alto - 27, 13, stroke=0, fill=1)
            c.restoreState()
            _icono(c, x + 22, self.alto - 31.5, icono, 12, color)
            c.setFillColor(colors.HexColor(TINTA))
            c.setFont("Archivo-ExtraBold", 24)
            c.drawString(x + 42, self.alto - 36, str(valor))
            c.setFont("Archivo-SemiBold", 8.5)
            c.drawString(x + 12, 22, nombre)
            c.setFont("Archivo", 7.5)
            c.setFillColor(colors.HexColor(TINTA_2))
            c.drawString(x + 12, 11, detalle)


class Aviso(Flowable):
    """Caja con icono y acento a la izquierda."""

    def __init__(self, icono, titulo, texto, color=NAVY_2):
        super().__init__()
        self.icono, self.titulo, self.texto, self.color = icono, titulo, texto, color
        self.p = Paragraph(f"<b>{_e(titulo)}</b>  {_e(texto)}", _estilos()["txt"])

    def wrap(self, aw, ah):
        _, h = self.p.wrap(UTIL - 52, ah)
        self.alto = max(h + 18, 36)
        return UTIL, self.alto

    def draw(self):
        c = self.canv
        _redondeado(c, 0, 0, UTIL, self.alto, 8, "#EEF3FA")
        c.setFillColor(colors.HexColor(self.color))
        c.roundRect(0, 0, 4, self.alto, 2, stroke=0, fill=1)
        _icono(c, 24, self.alto / 2 - 5, self.icono, 13, self.color)
        self.p.drawOn(c, 42, (self.alto - self.p.height) / 2)


class Pastilla(Flowable):
    def __init__(self, texto, color, tinte, ancho=58):
        super().__init__()
        self.texto, self.color, self.tinte, self.ancho = texto, color, tinte, ancho

    def wrap(self, *a):
        return self.ancho, 15

    def draw(self):
        c = self.canv
        _redondeado(c, 0, 0, self.ancho, 15, 7.5, self.tinte)
        c.setFillColor(colors.HexColor(self.color))
        c.circle(9, 7.5, 2.6, stroke=0, fill=1)
        c.setFillColor(colors.HexColor(TINTA))
        c.setFont("Archivo-SemiBold", 7.2)
        c.drawString(16, 4.6, self.texto)


class EncabezadoAsesor(Flowable):
    def __init__(self, nombre, n_ap, n_re, color=NAVY):
        super().__init__()
        self.nombre, self.n_ap, self.n_re, self.color = nombre, n_ap, n_re, color

    def wrap(self, *a):
        return UTIL, 34

    def draw(self):
        c = self.canv
        _redondeado(c, 0, 0, UTIL, 34, 8, SUPERFICIE)
        c.setFillColor(colors.HexColor(self.color))
        c.circle(18, 17, 11, stroke=0, fill=1)
        _icono(c, 18, 13, "asesor", 10.5, "#FFFFFF")
        c.setFillColor(colors.HexColor(TINTA))
        c.setFont("Archivo-Bold", 10.5)
        c.drawString(36, 13.5, self.nombre.title() if self.nombre else "(sin asesor)")
        x = UTIL - 10
        for n, est in ((self.n_re, RECHAZADO), (self.n_ap, APROBADO)):
            if not n:
                continue
            texto = f"{n} {est.lower()}{'s' if n != 1 else ''}"
            w = pdfmetrics.stringWidth(texto, "Archivo-SemiBold", 7.5) + 22
            _redondeado(c, x - w, 9.5, w, 15, 7.5, TINTE[est])
            c.setFillColor(colors.HexColor(COLOR[est]))
            c.circle(x - w + 9, 17, 2.6, stroke=0, fill=1)
            c.setFillColor(colors.HexColor(TINTA))
            c.setFont("Archivo-SemiBold", 7.5)
            c.drawString(x - w + 15, 14, texto)
            x -= w + 6


class Imagen(Flowable):
    def __init__(self, png, w, h):
        super().__init__()
        self.img, self.w, self.h = ImageReader(io.BytesIO(png)), w, h

    def wrap(self, *a):
        return self.w, self.h

    def draw(self):
        self.canv.drawImage(self.img, 0, 0, self.w, self.h, mask="auto")


# =============================================================================
# Gráficas (matplotlib)
# =============================================================================
def _mpl():
    import matplotlib
    matplotlib.use("Agg")
    import matplotlib.pyplot as plt
    plt.rcParams.update({"font.family": "Archivo", "font.size": 8, "axes.edgecolor": LINEA,
                         "axes.labelcolor": TINTA_2, "xtick.color": TINTA_2, "ytick.color": TINTA,
                         "axes.titleweight": "bold", "svg.fonttype": "none"})
    return plt


def _png(fig):
    buf = io.BytesIO()
    fig.savefig(buf, format="png", dpi=220, transparent=True)
    import matplotlib.pyplot as plt
    plt.close(fig)
    return buf.getvalue()


def _titulo(fig, titulo, subtitulo, ax):
    """Título, subtítulo y leyenda alineados a la izquierda de la figura (no del eje)."""
    fig.tight_layout(pad=0.4, rect=(0, 0, 1, 0.80))
    fig.text(0.01, 0.975, titulo, fontsize=10.5, fontweight="bold", color=TINTA, va="top")
    fig.text(0.01, 0.905, subtitulo, fontsize=7.5, color=TINTA_2, va="top")
    h, l = ax.get_legend_handles_labels()
    fig.legend(h, l, loc="upper left", bbox_to_anchor=(0.0, 0.855), ncol=2, frameon=False, fontsize=7.5,
               handlelength=1, handleheight=0.8, handletextpad=0.4, labelcolor=TINTA_2, borderaxespad=0)
    return _png(fig)


def grafica_asesores(faltan, w_pt, h_pt):
    plt = _mpl()
    t = (faltan.assign(Vendedor=faltan["Vendedor"].map(_corto))
         .pivot_table(index="Vendedor", columns="Estatus", values="Cliente", aggfunc="count", fill_value=0))
    for est in (APROBADO, RECHAZADO):
        if est not in t.columns:
            t[est] = 0
    t["total"] = t[APROBADO] + t[RECHAZADO]
    t = t.sort_values(["total", APROBADO]).tail(9)
    fig, ax = plt.subplots(figsize=(w_pt / 72, h_pt / 72))
    y = range(len(t))
    alto = min(0.55, 0.9 / max(1, len(t)) * 4)
    ax.barh(y, t[APROBADO], height=alto, color=COLOR[APROBADO], edgecolor="white", linewidth=1.5, label="Aprobados")
    ax.barh(y, t[RECHAZADO], left=t[APROBADO], height=alto, color=COLOR[RECHAZADO], edgecolor="white",
            linewidth=1.5, label="Rechazados")
    for i, total in enumerate(t["total"]):
        ax.text(total + 0.12, i, str(int(total)), va="center", ha="left", fontsize=8, color=TINTA, fontweight="bold")
    ax.set_yticks(list(y), t.index)
    ax.tick_params(axis="y", length=0, labelsize=7.8)
    ax.xaxis.set_visible(False)
    for s in ("top", "right", "bottom"):
        ax.spines[s].set_visible(False)
    ax.spines["left"].set_color(LINEA)
    ax.set_xlim(0, max(t["total"].max(), 1) * 1.18)
    return _titulo(fig, "Por asesor", "Clientes por registrar de cada asesor", ax)


def grafica_espera(faltan, hoy, w_pt, h_pt):
    """Línea de tiempo: cada cliente en la fecha de su solicitud, en el carril de su estatus."""
    plt = _mpl()
    import matplotlib.dates as mdates
    datos = faltan[faltan["Fecha"].notna()].copy()
    fig, ax = plt.subplots(figsize=(w_pt / 72, h_pt / 72))
    carril = {APROBADO: 1.0, RECHAZADO: 0.0}
    for est, y0 in carril.items():
        sub = datos[datos["Estatus"] == est].sort_values("Fecha")
        semanas = {}
        xs, ys = [], []
        for f in sub["Fecha"]:
            k = (f - dt.timedelta(days=f.weekday()))
            n = semanas.get(k, 0)
            semanas[k] = n + 1
            xs.append(pd.Timestamp(f))
            ys.append(y0 + (n % 4) * 0.13 - 0.2)          # se apilan los de la misma semana
        ax.scatter(xs, ys, s=58, color=COLOR[est], edgecolor="white", linewidth=1.5, zorder=3,
                   label=f"{est}s")
    ax.axvline(pd.Timestamp(hoy), color=TINTA_3, linewidth=1)
    ax.text(pd.Timestamp(hoy), 1.55, " hoy", fontsize=7, color=TINTA_2, va="top")
    if not datos.empty:
        viejo = datos.sort_values("Fecha").iloc[0]
        ax.annotate(f"{_corto(viejo['Cliente'])}\n{int(viejo['Dias'])} días", (pd.Timestamp(viejo["Fecha"]),
                    carril[viejo["Estatus"]] - 0.2), xytext=(4, -26), textcoords="offset points", fontsize=7,
                    color=TINTA_2, arrowprops=dict(arrowstyle="-", color=TINTA_3, lw=0.8))
        ini = min(datos["Fecha"]) - dt.timedelta(days=20)
        ax.set_xlim(pd.Timestamp(ini), pd.Timestamp(hoy + dt.timedelta(days=10)))
    ax.set_ylim(-0.75, 1.65)
    ax.set_yticks([0, 1], ["Rechazados", "Aprobados"])
    ax.tick_params(axis="y", length=0, labelsize=7.8)
    ax.xaxis.set_major_locator(mdates.MonthLocator())
    ax.xaxis.set_major_formatter(mdates.DateFormatter("%b"))
    meses = {"Jan": "Ene", "Apr": "Abr", "Aug": "Ago", "Dec": "Dic"}
    fig.canvas.draw()
    ax.set_xticks(ax.get_xticks(), [meses.get(t.get_text(), t.get_text()) for t in ax.get_xticklabels()])
    ax.tick_params(axis="x", labelsize=7.5, length=0)
    ax.grid(axis="x", color=LINEA, linewidth=0.8)
    for s in ("top", "right", "left"):
        ax.spines[s].set_visible(False)
    return _titulo(fig, "¿Desde cuándo esperan?", "Cada punto es un cliente, en la fecha de su solicitud", ax)


# =============================================================================
# Páginas
# =============================================================================
class _Doc(BaseDocTemplate):
    def __init__(self, buf, corte, total):
        super().__init__(buf, pagesize=letter, leftMargin=MARGEN, rightMargin=MARGEN, topMargin=40, bottomMargin=40,
                         title="Clientes por registrar en Sale-U", author="MG Colima")
        self.corte, self.total = corte, total
        portada = Frame(MARGEN, 40, UTIL, ALTO - 40 - 182, id="portada", leftPadding=0, rightPadding=0,
                        topPadding=0, bottomPadding=0)
        resto = Frame(MARGEN, 40, UTIL, ALTO - 40 - 62, id="resto", leftPadding=0, rightPadding=0, topPadding=0,
                      bottomPadding=0)
        self.addPageTemplates([PageTemplate("portada", [portada], onPage=self._portada),
                               PageTemplate("resto", [resto], onPage=self._interior)])

    def _pie(self, c):
        c.setFont("Archivo", 7)
        c.setFillColor(colors.HexColor(TINTA_3))
        c.drawString(MARGEN, 22, "MG Colima · Fuente: Google Sheet de seguimiento (MG Colima Rechazos y Bitacora)")
        c.drawRightString(ANCHO - MARGEN, 22, f"Página {c.getPageNumber()}")

    def _portada(self, c, doc):
        c.saveState()
        alto = 162
        c.setFillColor(colors.HexColor(NAVY))
        c.rect(0, ALTO - alto, ANCHO, alto, stroke=0, fill=1)
        c.setFillColor(colors.white)                            # círculos decorativos
        for (x, y, r, a) in ((ANCHO - 70, ALTO - 30, 95, 0.06), (ANCHO - 160, ALTO - 150, 60, 0.05),
                             (ANCHO - 20, ALTO - 140, 40, 0.07)):
            c.setFillAlpha(a)
            c.circle(x, y, r, stroke=0, fill=1)
        c.setFillAlpha(1)
        c.setFillColor(colors.HexColor(NAVY_2))
        c.circle(MARGEN + 26, ALTO - 66, 26, stroke=0, fill=1)
        _icono(c, MARGEN + 26, ALTO - 73, "auto", 19, "#FFFFFF")
        c.setFillColor(colors.white)
        c.setFont("Archivo-ExtraBold", 23)
        c.drawString(MARGEN + 66, ALTO - 62, "Clientes por registrar en Sale-U")
        c.setFont("Archivo", 10)
        c.setFillColor(colors.HexColor("#C9D6EA"))
        c.drawString(MARGEN + 66, ALTO - 80, "Aprobados y Rechazados de MG Colima que todavía no aparecen en Sale-U")
        chip = f"Corte {self.corte:%d/%m/%Y}"
        w = pdfmetrics.stringWidth(chip, "Archivo-SemiBold", 8.5) + 30
        c.setFillColor(colors.white)
        c.setFillAlpha(0.12)
        c.roundRect(MARGEN + 66, ALTO - 118, w, 20, 10, stroke=0, fill=1)
        c.setFillAlpha(1)
        _icono(c, MARGEN + 80, ALTO - 111.5, "calendario", 8.5, "#FFFFFF")
        c.setFillColor(colors.white)
        c.setFont("Archivo-SemiBold", 8.5)
        c.drawString(MARGEN + 91, ALTO - 111.5, chip)
        self._pie(c)
        c.restoreState()

    def _interior(self, c, doc):
        c.saveState()
        c.setFillColor(colors.HexColor(NAVY))
        c.rect(0, ALTO - 34, ANCHO, 34, stroke=0, fill=1)
        _icono(c, MARGEN + 6, ALTO - 22, "auto", 10, "#FFFFFF")
        c.setFillColor(colors.white)
        c.setFont("Archivo-Bold", 9.5)
        c.drawString(MARGEN + 20, ALTO - 22, "Clientes por registrar en Sale-U")
        c.setFont("Archivo", 8.5)
        c.setFillColor(colors.HexColor("#C9D6EA"))
        c.drawRightString(ANCHO - MARGEN, ALTO - 22, f"Corte {self.corte:%d/%m/%Y} · {self.total} clientes")
        self._pie(c)
        c.restoreState()


def _tabla_clientes(df, color_estatus=True):
    st = _estilos()
    datos = [[Paragraph("Estatus", st["enc"]), Paragraph("Cliente", st["enc"]), Paragraph("Contacto", st["enc"]),
              Paragraph("Solicitud", st["enc"]), Paragraph("Mensualidad", st["enc_d"])]]
    sub = []
    for _, r in df.iterrows():
        est = r["Estatus"]
        clave = est if color_estatus else "confirmar"
        detalle = " · ".join(x for x in (r["Unidad"], r["Financiera"]) if x)
        contacto = (f'{_ico("telefono")}  {_e(r["Telefono"]) or "sin celular"}'
                    + (f'<br/><font size="7" color="{TINTA_2}">{_ico("correo", tam=6.5)}  {_e(r["Correo"])}</font>'
                       if r["Correo"] else ""))
        dias = r["Dias"]
        solicitud = (f'{_ico("calendario")}  {_fecha(r["Fecha"])}'
                     + (f'<br/><font size="7" color="{TINTA_2}">hace {int(dias)} días</font>' if dias == dias and
                        dias is not None else ""))
        mens = r["Mensualidad"]
        datos.append([
            Pastilla(est.upper(), COLOR[clave], TINTE[clave]),
            Paragraph(f"{_e(r['Cliente'])}" + (f'<br/><font name="Archivo" size="7" color="{TINTA_2}">'
                                               f'{_e(detalle)}</font>' if detalle else ""), st["nombre"]),
            Paragraph(contacto, st["celda"]), Paragraph(solicitud, st["celda"]),
            Paragraph(f"${mens:,.2f}" if mens == mens and mens else "—", st["celda_d"])])
        como = str(r.get("Cómo está en Sale-U") or "").strip()
        fuente = str(r.get("Fuente") or "").strip()
        if como or fuente:                  # renglón debajo del cliente: cómo encontrarlo en Sale-U y su fuente
            texto = (f'{_ico("info", NAVY_2, 6.8)}  <font color="{NAVY_2}"><b>En Sale-U:</b></font> {_e(como or "—")}'
                     + (f'  ·  <font color="{NAVY_2}"><b>Fuente:</b></font> {_e(fuente)}' if fuente else ""))
            datos.append(["", Paragraph(texto, st["celda"]), "", "", ""])
            sub.append(len(datos) - 1)
    t = Table(datos, colWidths=[68, 186, 128, 82, 68], repeatRows=1)
    estilo = [
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor(LINEA)),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor(LINEA)),
        ("TOPPADDING", (0, 1), (-1, -1), 6), ("BOTTOMPADDING", (0, 1), (-1, -1), 6),
        ("LEFTPADDING", (0, 0), (-1, -1), 4), ("RIGHTPADDING", (0, 0), (-1, -1), 4)]
    for i in sub:
        estilo += [("SPAN", (1, i), (4, i)), ("TOPPADDING", (0, i), (-1, i), 0), ("BOTTOMPADDING", (0, i), (-1, i), 5),
                   ("LINEBELOW", (0, i - 1), (-1, i - 1), 0, colors.white),
                   ("BACKGROUND", (1, i), (4, i), colors.HexColor("#F3F6FB"))]
    t.setStyle(TableStyle(estilo))
    return t


def pdf(faltan, por_confirmar=None, hoy=None):
    """Bytes del PDF."""
    _fuentes()
    st = _estilos()
    hoy = hoy or dt.date.today()
    por_confirmar = por_confirmar if por_confirmar is not None else faltan.iloc[0:0]
    buf = io.BytesIO()
    doc = _Doc(buf, hoy, len(faltan))
    if "Falta" not in faltan.columns:
        faltan = faltan.assign(Falta=FALTA_ALTA)
    n_ap = int((faltan["Estatus"] == APROBADO).sum()) if not faltan.empty else 0
    n_re = int((faltan["Estatus"] == RECHAZADO).sum()) if not faltan.empty else 0
    n_alta = int((faltan["Falta"] == FALTA_ALTA).sum()) if not faltan.empty else 0
    n_cred = int((faltan["Falta"] == FALTA_CREDITO).sum()) if not faltan.empty else 0
    el = [NextPageTemplate("resto"),
          Cifras([("Por registrar", len(faltan), "Aprobados + Rechazados", "meta", NAVY_2),
                  ("Dar de alta", n_alta, "no están en Sale-U", "usuario_mas", FALTA[FALTA_ALTA][0]),
                  ("Registrar crédito", n_cred, "ya son lead en Sale-U", "dinero", FALTA[FALTA_CREDITO][0]),
                  ("Aprobados", n_ap, "ya tienen crédito", "aprobado", COLOR[APROBADO]),
                  ("Rechazados", n_re, "para seguimiento", "rechazado", COLOR[RECHAZADO])]),
          Spacer(1, 12)]
    if faltan.empty:
        el += [Aviso("aprobado", "Todo al día.", "No hay Aprobados ni Rechazados pendientes de registrar en Sale-U "
                                                 "según el Sheet.", COLOR[APROBADO])]
    else:
        w = (UTIL - 14) / 2
        h = 236
        graf = Table([[Imagen(grafica_asesores(faltan, w, h), w, h), Imagen(grafica_espera(faltan, hoy, w, h), w, h)]],
                     colWidths=[w + 7, w + 7])
        graf.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                                  ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        el += [graf, Spacer(1, 10)]
        if n_ap:
            el.append(Aviso("idea", "Empieza por los Aprobados.",
                            f"{n_ap} cliente(s) ya tienen el crédito aprobado y no están registrados en Sale-U: son la "
                            "venta más cercana. Los Rechazados se registran para darles seguimiento a los 180 días."))
        primera = True
        for falta in (FALTA_ALTA, FALTA_CREDITO):
            parte = faltan[faltan["Falta"] == falta]
            if parte.empty:
                continue
            color, icono, titulo, texto = FALTA[falta]
            el += [PageBreak() if primera else CondPageBreak(200),
                   Paragraph(f'{_ico(icono, color, 13)}  {_e(titulo)} ({len(parte)})', st["h2"]),
                   Paragraph(_e(texto + " Por asesor, Aprobados primero, con su contacto, la fecha de su solicitud "
                                        "y la mensualidad que se le cotizó."), st["sub"]), Spacer(1, 8)]
            primera = False
            orden = (parte.assign(_ap=(parte["Estatus"] == APROBADO).astype(int), _v=parte["Vendedor"].replace("", "~"))
                     .groupby("_v").agg(ap=("_ap", "sum"), n=("Cliente", "count"))
                     .sort_values(["ap", "n"], ascending=False).index)
            for v in orden:
                sub = parte[parte["Vendedor"].replace("", "~") == v]
                cab = EncabezadoAsesor("" if v == "~" else v, int((sub["Estatus"] == APROBADO).sum()),
                                       int((sub["Estatus"] == RECHAZADO).sum()))
                el += [CondPageBreak(110), KeepTogether([cab, Spacer(1, 4), _tabla_clientes(sub)]), Spacer(1, 14)]
    if por_confirmar is not None and not por_confirmar.empty:
        el += [CondPageBreak(160), Paragraph(f'{_ico("espera", "#8a8984", 13)}  Aprobados por confirmar '
                                             f'({len(por_confirmar)})', st["h2"]),
               Paragraph(_e("La Bitácora de F&I no trae el dato de Sale-U de estos aprobados. Revísalos en Sale-U y, "
                            "si faltan, regístralos."), st["sub"]), Spacer(1, 6),
               _tabla_clientes(por_confirmar, color_estatus=False), Spacer(1, 14)]
    el += [CondPageBreak(70), Aviso("info", "Cómo se armó.",
                                    "Rechazados de «MG Colima Rechazos» y Aprobados de la Bitacora, según su Registro en "
                                    "Sale-U, que marca el cuadre de Sale-U: «No Registrado» = no está ni en el reporte Fuente "
                                    "ni en el de Crédito (dar de alta); «Falta crédito» = está en Fuente pero no en Crédito "
                                    "(registrar el crédito). Si un cliente aparece registrado en cualquiera de las dos "
                                    "pestañas, no se lista; cada cliente sale una vez con su estatus más reciente.",
                                    TINTA_3)]
    doc.build(el)
    return buf.getvalue()


def excel(faltan, por_confirmar=None, hoy=None):
    """Bytes del Excel con la misma lista del PDF (una hoja por lo que falta, más los por confirmar)."""
    from cuadres import salidas as sal
    hoy = hoy or dt.date.today()
    cols = ["Falta", "Estatus", "Cliente", "Cómo está en Sale-U", "Fuente", "Telefono", "Vendedor", "Financiera", "Unidad",
            "Fecha", "Dias", "Mensualidad", "Correo", "Origen", "ID"]
    vista = lambda df: df.assign(Fecha=df["Fecha"].map(_fecha)) if not df.empty else df
    hojas = [("Por registrar", vista(faltan), cols)]
    if "Falta" in faltan.columns:
        for f in (FALTA_ALTA, FALTA_CREDITO):
            hojas.append((f, vista(faltan[faltan["Falta"] == f]), cols[1:]))
    if por_confirmar is not None and not por_confirmar.empty:
        hojas.append(("Por confirmar", vista(por_confirmar), cols[1:]))
    return sal.excel(None, "Clientes por registrar en Sale-U – MG Colima",
                     f"Corte {hoy:%d/%m/%Y} · {len(faltan)} clientes", hojas,
                     {FALTA_ALTA: "FFC7CE", FALTA_CREDITO: "FFEB9C"})
