# -*- coding: utf-8 -*-
"""
PDF de seguimientos vencidos por asesor: una página (o más) por asesor con sus clientes que ya cumplieron la fecha
de seguimiento y no tienen el primer contacto, del más atrasado al más reciente.

  asesores(vencidos)            → nombres con vencidos (más vencidos primero)
  pdf(vencidos, asesor=None)    → bytes del PDF (asesor=None: todos, cada uno en su página)
Mismo estilo que los demás PDF (cuadres/reporte_saleu.py).
"""
import datetime as dt
import io

from reportlab.lib import colors
from reportlab.lib.pagesizes import letter
from reportlab.platypus import (BaseDocTemplate, Frame, PageBreak, PageTemplate, Paragraph, Spacer, Table,
                                TableStyle)

from cuadres.reporte_saleu import (NAVY, NAVY_2, TINTA_2, TINTA_3, LINEA, MARGEN, UTIL, _e, _estilos, _fuentes,
                                   _icono, _ico)
from cuadres.salidas import telefono

ANCHO, ALTO = letter
ROJO = "#d03b3b"


def _nombre(a):
    return str(a or "").strip() or "(sin asesor)"


def asesores(vencidos):
    if vencidos is None or vencidos.empty:
        return []
    return list(vencidos["asesor"].map(_nombre).value_counts().index)


class _Doc(BaseDocTemplate):
    def __init__(self, buf, hoy):
        super().__init__(buf, pagesize=letter, leftMargin=MARGEN, rightMargin=MARGEN, topMargin=40, bottomMargin=34,
                         title="Seguimientos vencidos", author="MG Colima")
        self.hoy = hoy
        marco = Frame(MARGEN, 34, UTIL, ALTO - 34 - 76, leftPadding=0, rightPadding=0, topPadding=0, bottomPadding=0)
        self.addPageTemplates([PageTemplate("p", [marco], onPage=self._fondo)])

    def _fondo(self, c, doc):
        c.saveState()
        c.setFillColor(colors.HexColor(NAVY))
        c.rect(0, ALTO - 60, ANCHO, 60, stroke=0, fill=1)
        c.setFillColor(colors.HexColor(NAVY_2))
        c.circle(MARGEN + 15, ALTO - 30, 15, stroke=0, fill=1)
        _icono(c, MARGEN + 15, ALTO - 34, "reloj", 12, "#FFFFFF")
        c.setFillColor(colors.white)
        c.setFont("Archivo-ExtraBold", 16)
        c.drawString(MARGEN + 40, ALTO - 28, "Seguimientos vencidos")
        c.setFont("Archivo", 8.5)
        c.setFillColor(colors.HexColor("#C9D6EA"))
        c.drawString(MARGEN + 40, ALTO - 42, f"Clientes rechazados que ya cumplieron su fecha de seguimiento y no tienen "
                                             f"el primer contacto · corte {self.hoy:%d/%m/%Y}")
        c.setFont("Archivo", 7)
        c.setFillColor(colors.HexColor(TINTA_3))
        c.drawString(MARGEN, 18, "MG Colima · Fuente: Google Sheet de seguimiento (MG Colima Rechazos)")
        c.drawRightString(ANCHO - MARGEN, 18, f"Página {c.getPageNumber()}")
        c.restoreState()


def _bloque(asesor, df, st):
    el = [Paragraph(f'{_ico("asesor", NAVY_2, 12)}  {_e(asesor.title())}', st["h2"]),
          Paragraph(_e(f"{len(df)} cliente(s) con seguimiento vencido. Contáctalos y captura en el Sheet la fecha, el "
                       "canal y si contestaron."), st["sub"]), Spacer(1, 8)]
    enc = ["Cliente", "Celular", "Solicitud", "Tocaba el", "Atraso", "Financiera · unidad"]
    datos = [[Paragraph(_e(h), st["enc_d"] if h == "Atraso" else st["enc"]) for h in enc]]
    for _, r in df.sort_values("dias", ascending=False).iterrows():
        fecha = r["fecha"].strftime("%d/%m/%Y") if isinstance(r["fecha"], (dt.date, dt.datetime)) else ""
        fseg = r["fseg"].strftime("%d/%m/%Y") if isinstance(r["fseg"], (dt.date, dt.datetime)) else ""
        dias = int(r["dias"])
        datos.append([
            Paragraph(_e(str(r["nombre"]).title()) + (f'<br/><font size="7" color="{TINTA_2}">{_e(r["id"])}</font>'
                                                       if r.get("id") else ""), st["nombre"]),
            Paragraph(f'{_ico("telefono")}  {_e(telefono(r["tel"])) or "sin celular"}', st["celda"]),
            Paragraph(_e(fecha), st["celda"]), Paragraph(_e(fseg), st["celda"]),
            Paragraph(f'<font color="{ROJO}"><b>{dias} {"día" if dias == 1 else "días"}</b></font>', st["celda_d"]),
            Paragraph(_e(" · ".join(x for x in (str(r.get("fin") or ""), str(r.get("unidad") or "")) if x)),
                      st["celda"])])
    t = Table(datos, colWidths=[150, 92, 62, 62, 52, 114], repeatRows=1)
    t.setStyle(TableStyle([
        ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
        ("LINEBELOW", (0, 0), (-1, 0), 0.8, colors.HexColor(LINEA)),
        ("LINEBELOW", (0, 1), (-1, -1), 0.4, colors.HexColor(LINEA)),
        ("TOPPADDING", (0, 1), (-1, -1), 5), ("BOTTOMPADDING", (0, 1), (-1, -1), 5),
        ("LEFTPADDING", (0, 0), (-1, -1), 3), ("RIGHTPADDING", (0, 0), (-1, -1), 3)]))
    return el + [t]


def pdf(vencidos, asesor=None, hoy=None):
    """Bytes del PDF. asesor=None: todos los asesores, cada uno en su página."""
    _fuentes()
    st = _estilos()
    hoy = hoy or dt.date.today()
    df = vencidos.assign(_a=vencidos["asesor"].map(_nombre)) if vencidos is not None else None
    lista = [asesor] if asesor else asesores(vencidos)
    buf = io.BytesIO()
    doc = _Doc(buf, hoy)
    el = []
    for i, a in enumerate(lista):
        sub = df[df["_a"] == a] if df is not None else None
        if sub is None or sub.empty:
            continue
        if el:
            el.append(PageBreak())
        el += _bloque(a, sub, st)
    if not el:
        el = [Paragraph("Sin seguimientos vencidos.", st["txt"])]
    doc.build(el)
    return buf.getvalue()
