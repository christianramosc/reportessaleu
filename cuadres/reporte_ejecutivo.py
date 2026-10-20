# -*- coding: utf-8 -*-
"""
Resumen ejecutivo de una página (PDF) con lo más importante del Sheet de seguimiento:

  · indicadores del año: rechazos, seguimiento al día, contacto efectivo, registro en Sale-U, vencidos, Revision
  · gráficas (sin nombres): seguimientos más atrasados (días por ID), agenda de seguimientos, embudo de contacto y
    registro en Sale-U. tablero(d) da los datos de esas y de las demás gráficas del Inicio.
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
from cuadres.reporte_saleu import (Cifras, Imagen, NAVY, NAVY_2, TINTA, TINTA_2, TINTA_3, LINEA, MARGEN, UTIL,
                                   _estilos, _fuentes, _icono, _ico, _mpl, _titulo)

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
    return dict(k=k, faltan=faltan, anio=anio, hoy=hoy)


# =============================================================================
# Tablero (datos de las gráficas, sin nombres de clientes ni de asesores)
# =============================================================================
RANGOS = [(0, 7, "0–7 días"), (8, 15, "8–15 días"), (16, 30, "16–30 días"), (31, 60, "31–60 días"),
          (61, 10 ** 6, "Más de 60")]
COLOR_RANGO = ["#fab219", "#ec835a", "#e0603f", "#d03b3b", "#8f1d1d"]
SEMANAS = 8


def _rango(dias):
    return next(r for a, b, r in RANGOS if a <= dias <= b)


def tablero(d):
    """DataFrames de las gráficas del Inicio y del PDF."""
    k, hoy = d["k"], d["hoy"]
    v = k["vencidos"]
    venc = pd.DataFrame({
        "Seguimiento": [str(r.get("id") or "").strip() or f"Fila {r.get('fila', '')}" for _, r in v.iterrows()],
        "Días": [int(x) for x in v["dias"]] if len(v) else [],
        "Financiera": [str(r.get("fin") or "").title() or "—" for _, r in v.iterrows()],
        "Tocaba el": [r["fseg"].strftime("%d/%m/%Y") if r.get("fseg") else "" for _, r in v.iterrows()]})
    venc["Atraso"] = venc["Días"].map(_rango) if len(venc) else []
    venc = venc.drop_duplicates("Seguimiento").sort_values("Días", ascending=False).reset_index(drop=True)
    antig = pd.DataFrame({"Atraso": [r for _, _, r in RANGOS],
                          "Seguimientos": [int((venc["Atraso"] == r).sum()) for _, _, r in RANGOS],
                          "Color": COLOR_RANGO})

    # agenda: vencidos + lo que toca en las próximas semanas (lunes a domingo)
    lunes = hoy - dt.timedelta(days=hoy.weekday())
    sh = k.get("sin_hacer")
    filas = [dict(Semana="Vencidos", Orden=-1, Seguimientos=len(venc), Tipo="Vencido")]
    for i in range(SEMANAS):
        ini = lunes + dt.timedelta(weeks=i)
        n = 0 if sh is None or sh.empty else int(sh["fseg"].map(
            lambda f: bool(f) and max(ini, hoy) <= f < ini + dt.timedelta(days=7)).sum())
        filas.append(dict(Semana="Esta semana" if i == 0 else f"{ini:%d/%m}", Orden=i, Seguimientos=n,
                          Tipo="Esta semana" if i == 0 else "Próximo"))
    agenda = pd.DataFrame(filas)

    embudo = pd.DataFrame({"Paso": ["Rechazos del año", "Ya les toca (180 días)", "Con primer contacto",
                                    "Contestaron"],
                           "Clientes": [k["rechazos"], k["ya_toca"], k["contactados"], k["contestaron"]]})
    contacto = pd.DataFrame({"Resultado": ["Contestó", "No contestó", "Sin dato"],
                             "Clientes": [k["contestaron"], k.get("no_contestaron", 0), k["sin_contesta"]],
                             "Color": [ESTADO["bien"], ESTADO["serio"], ESTADO["neutro"]]})
    canal = pd.DataFrame(sorted(k.get("por_canal", {}).items(), key=lambda x: -x[1]), columns=["Canal", "Clientes"])
    saleu = pd.DataFrame({"Estado": ["Registrado", "Falta crédito", "Dar de alta"],
                          "Clientes": [k.get("registrados", 0), k.get("falta_credito", 0), k.get("falta_alta", 0)],
                          "Color": [ESTADO["bien"], ESTADO["aviso"], ESTADO["critico"]]})
    fin = pd.DataFrame(sorted(k["por_fin"].items(), key=lambda x: -x[1]), columns=["Financiera", "Rechazos"])
    al_dia = pd.DataFrame({"Estado": ["Con seguimiento", "Sin seguimiento"],
                           "Clientes": [k["al_dia"], k["ya_toca"] - k["al_dia"]],
                           "Color": [ESTADO["bien"], ESTADO["critico"]]})
    return dict(vencidos=venc, antiguedad=antig, agenda=agenda, embudo=embudo, contacto=contacto, canal=canal,
                saleu=saleu, financiera=fin, al_dia=al_dia)


# =============================================================================
# Gráficas del PDF (matplotlib)
# =============================================================================
def _ejes(ax, eje="y"):
    for s in ("top", "right", "left" if eje == "y" else "bottom"):
        ax.spines[s].set_visible(False)
    ax.spines["bottom" if eje == "y" else "left"].set_color(LINEA)
    ax.tick_params(length=0, labelsize=7, colors=TINTA_2)
    ax.grid(axis=eje, color=LINEA, linewidth=0.6)
    ax.set_axisbelow(True)


def grafica_vencidos(t, w, h, top=12):
    plt = _mpl()
    v = t["vencidos"].head(top)
    fig, ax = plt.subplots(figsize=(w / 72, h / 72))
    if v.empty:
        ax.text(0.5, 0.5, "Sin seguimientos vencidos", ha="center", va="center", color=ESTADO["bien"], fontsize=9)
        ax.axis("off")
    else:
        y = list(range(len(v)))[::-1]
        col = [COLOR_RANGO[[r for _, _, r in RANGOS].index(x)] for x in v["Atraso"]]
        ax.barh(y, v["Días"], color=col, height=0.62)
        for yi, d in zip(y, v["Días"]):
            ax.text(d + v["Días"].max() * 0.02, yi, f"{d} d", va="center", fontsize=6.8, color=TINTA)
        ax.set_yticks(y, v["Seguimiento"])
        ax.set_xlim(0, v["Días"].max() * 1.18)
        _ejes(ax, "x")
        ax.xaxis.set_visible(False)
    n = len(t["vencidos"])
    return _titulo(fig, "Seguimientos más atrasados", f"Días de atraso por ID · {n} vencido(s)"
                   + (f", se muestran {top}" if n > top else ""), ax)


def grafica_agenda(t, w, h):
    plt = _mpl()
    a = t["agenda"]
    fig, ax = plt.subplots(figsize=(w / 72, h / 72))
    col = [ESTADO["critico"] if x == "Vencido" else (NAVY_2 if x == "Esta semana" else SERIE[0])
           for x in a["Tipo"]]
    ax.bar(range(len(a)), a["Seguimientos"], color=col, width=0.62)
    for i, n in enumerate(a["Seguimientos"]):
        if n:
            ax.text(i, n + max(a["Seguimientos"].max(), 1) * 0.03, str(n), ha="center", va="bottom", fontsize=7,
                    color=TINTA, fontweight="bold")
    ax.set_xticks(range(len(a)), [x.replace("Esta semana", "Esta\nsemana") for x in a["Semana"]], fontsize=6.3)
    ax.set_ylim(0, max(a["Seguimientos"].max(), 1) * 1.22)
    _ejes(ax)
    return _titulo(fig, "Agenda de seguimientos", f"Vencidos y lo que toca en las próximas {SEMANAS} semanas", ax)


def grafica_embudo(t, w, h):
    plt = _mpl()
    e = t["embudo"]
    fig, ax = plt.subplots(figsize=(w / 72, h / 72))
    tope = max(e["Clientes"].max(), 1)
    y = list(range(len(e)))[::-1]
    col = [NAVY, NAVY_2, SERIE[0], ESTADO["bien"]]
    for yi, (_, r), c in zip(y, e.iterrows(), col):
        ax.barh(yi, r["Clientes"], left=(tope - r["Clientes"]) / 2, color=c, height=0.7)
        pct = f"  ({r['Clientes'] / e['Clientes'].iloc[0]:.0%})" if e["Clientes"].iloc[0] else ""
        angosta = r["Clientes"] < tope * 0.3
        ax.text((tope + r["Clientes"]) / 2 + tope * 0.02 if angosta else tope / 2, yi, f"{r['Clientes']}{pct}",
                ha="left" if angosta else "center", va="center", fontsize=7.5,
                color=TINTA if angosta else "white", fontweight="bold")
    ax.set_yticks(y, e["Paso"])
    ax.set_xlim(0, tope)
    for s in ("top", "right", "bottom", "left"):
        ax.spines[s].set_visible(False)
    ax.xaxis.set_visible(False)
    ax.tick_params(length=0, labelsize=7.2)
    return _titulo(fig, "Embudo de contacto", f"Rechazos de {t.get('anio', '')}".strip() or "Del año", ax)


def grafica_dona(df, campo, w, h, titulo, subtitulo):
    plt = _mpl()
    fig, ax = plt.subplots(figsize=(w / 72, h / 72))
    df = df[df["Clientes"] > 0]
    total = int(df["Clientes"].sum())
    if total:
        ax.pie(df["Clientes"], colors=df["Color"], startangle=90, counterclock=False,
               wedgeprops=dict(width=0.34, edgecolor="white", linewidth=1.5))
        ax.text(0, 0.05, str(total), ha="center", va="center", fontsize=13, fontweight="bold", color=TINTA)
        ax.text(0, -0.22, "clientes", ha="center", va="center", fontsize=6.5, color=TINTA_2)
        for (_, r) in df.iterrows():
            ax.bar(0, 0, color=r["Color"], label=f"{r[campo]}  {r['Clientes']} ({r['Clientes'] / total:.0%})")
    ax.set_aspect("equal")
    ax.axis("off")
    fig.tight_layout(pad=0.4, rect=(0.38, 0, 1, 0.82))
    fig.text(0.01, 0.975, titulo, fontsize=10.5, fontweight="bold", color=TINTA, va="top")
    fig.text(0.01, 0.905, subtitulo, fontsize=7.5, color=TINTA_2, va="top")
    hh, ll = ax.get_legend_handles_labels()
    fig.legend(hh, ll, loc="center left", bbox_to_anchor=(0.0, 0.42), frameon=False, fontsize=7.2,
               handlelength=1, handleheight=0.8, labelcolor=TINTA_2, labelspacing=0.9)
    from cuadres.reporte_saleu import _png
    return _png(fig)


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
        c.drawString(MARGEN + 48, ALTO - 52, f"Rechazos, seguimiento y Sale-U · corte {self.hoy:%d/%m/%Y}")
        c.setFont("Archivo", 7)
        c.setFillColor(colors.HexColor(TINTA_3))
        c.drawString(MARGEN, 18, "Fuente: Google Sheet de seguimiento (MG Colima Rechazos, Bitacora y Revision) · sin nombres de clientes")
        c.restoreState()


def prioridades(d):
    """Frases de «Prioridades» (con <b> para negritas)."""
    return _prioridades(d)


def _prioridades(d):
    k, faltan = d["k"], d["faltan"]
    out = []
    if len(k["vencidos"]):
        peor = k["vencidos"].iloc[0]
        out.append(f"<b>{len(k['vencidos'])} seguimientos vencidos</b> (el más atrasado lleva {int(peor['dias'])} días). "
                   f"Otros {len(k['por_vencer'])} vencen en los próximos 15 días.")
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


def pdf(base, anio=None, hoy=None, d=None):
    """Bytes del PDF de una página (d: lo que devuelve datos(), si ya se calculó)."""
    _fuentes()
    st = _estilos()
    d = d or datos(base, anio, hoy)
    k = d["k"]
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
    t = tablero(d)
    t["anio"] = d["anio"]
    w, h = (UTIL - 14) / 2, 150

    def fila(a, b):
        g = Table([[Imagen(a, w, h), Imagen(b, w, h)]], colWidths=[w + 14, w])
        g.setStyle(TableStyle([("LEFTPADDING", (0, 0), (-1, -1), 0), ("RIGHTPADDING", (0, 0), (-1, -1), 0),
                               ("VALIGN", (0, 0), (-1, -1), "TOP")]))
        return g
    el += [fila(grafica_vencidos(t, w, h, top=8), grafica_agenda(t, w, h)), Spacer(1, 8),
           fila(grafica_embudo(t, w, h),
                grafica_dona(t["saleu"], "Estado", w, h, "Registro en Sale-U",
                             "Rechazados y aprobados del año")), Spacer(1, 10)]
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
