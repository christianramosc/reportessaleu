# -*- coding: utf-8 -*-
"""
Resumen del Google Sheet de seguimiento: los indicadores principales y quién falta registrar en Sale-U.

  indicadores(base, anio) → dict con los KPI del año y las tablas para las gráficas:
      rechazos del año, seguimiento al día (de los que ya cumplieron 180 días), contacto efectivo,
      registrados en Sale-U, seguimientos vencidos y pendientes en Revision.
  faltantes_saleu(base)   → (faltan, por_confirmar): clientes Aprobados y Rechazados sin registro en Sale-U,
      solo con lo que dice el Sheet (columna Registro en Sale-U, la marca el cuadre de Sale-U):
        · «No Registrado» → falta darlo de alta en Sale-U (no está ni en el reporte Fuente ni en el de Crédito).
        · «Falta crédito» → ya es lead en Sale-U (reporte Fuente), falta registrar su crédito.
        · Rechazados: «MG Colima Rechazos»; Aprobados: pestaña Bitacora con Estatus final «Aprobado».
          Si un aprobado no trae el dato, va a «por confirmar».
      Un cliente que en cualquier fuente aparece registrado (Registrado / SI) no se lista. Cada cliente sale una
      vez (mismo celular o nombre), con su estatus más reciente: si fue rechazado y después aprobado, es Aprobado.
"""
import datetime as dt

import pandas as pd

from cuadres.base import TAB_RECHAZOS, TAB_BITACORA, norm
from cuadres.fuentes import name_key, similar, to_date, to_phone, to_num, txt
from cuadres.salidas import telefono

DIAS_SEGUIMIENTO = 180
TAB_REVISION = "Revision"
APROBADO, RECHAZADO = "Aprobado", "Rechazado"
FALTA_ALTA, FALTA_CREDITO = "Dar de alta", "Registrar crédito"
COL_SALEU = "Cómo está en Sale-U"
COLS_FALTAN = ["Falta", "Estatus", "Cliente", "Telefono", "Vendedor", "Financiera", "Unidad", "Fecha", "Mensualidad",
               "Correo", "Dias", "Origen", "ID", COL_SALEU, "Fuente"]


def codigo_saleu(v):
    """Registro en Sale-U del Sheet (o SI/NO de F&I) → «SI», «CRED» (falta crédito), «NO» o ""."""
    n = norm(v)
    if n in ("REGISTRADO", "SI"):
        return "SI"
    if n == "FALTA CREDITO":
        return "CRED"
    if n in ("NO REGISTRADO", "NO"):
        return "NO"
    return ""


def _col(df, *nombres):
    for n in nombres:
        c = next((c for c in df.columns if norm(c) == norm(n)), None)
        if c:
            return c
    return None


def _vacio(v):
    return v is None or (isinstance(v, float) and v != v) or str(v).strip() in ("", "None", "nan")


def _rechazos(base):
    """Filas de «MG Colima Rechazos» con los campos del resumen."""
    if TAB_RECHAZOS not in base.pestanas():
        return pd.DataFrame()
    df = base.leer(TAB_RECHAZOS)
    if df.empty:
        return df
    c = dict(fecha=_col(df, "Fecha ingreso"), nombre=_col(df, "Nombre de Cliente"), tel=_col(df, "Celular"),
             fin=_col(df, "Financiera"), asesor=_col(df, "Asesor de Ventas"), unidad=_col(df, "Unidad"),
             mens=_col(df, "Mensualidad"), correo=_col(df, "Correo"), saleu=_col(df, "Registro en Sale-U"),
             hecho=_col(df, "FECHA REALIZADO 1"), contesta=_col(df, "CONTESTA 1"), com=_col(df, "COMENTARIOS"),
             id=_col(df, "ID"))
    g = lambda r, k: r[c[k]] if c[k] else None
    filas = []
    for _, r in df.iterrows():
        if _vacio(g(r, "nombre")):
            continue
        fecha = to_date(g(r, "fecha"))
        filas.append(dict(
            fila=int(r["_fila"]), fecha=fecha, nombre=" ".join(str(g(r, "nombre")).split()),
            tel=to_phone(g(r, "tel")), fin=(txt(g(r, "fin")) or "").upper(), asesor=(txt(g(r, "asesor")) or "").upper(),
            unidad=txt(g(r, "unidad")) or "", mens=to_num(g(r, "mens")), correo=txt(g(r, "correo")) or "",
            saleu=norm(g(r, "saleu")), hecho=to_date(g(r, "hecho")) or (None if _vacio(g(r, "hecho")) else True),
            contesta=norm(g(r, "contesta")), com=norm(g(r, "com")), id=txt(g(r, "id")) or "",
            fseg=fecha + dt.timedelta(days=DIAS_SEGUIMIENTO) if fecha else None))
    return pd.DataFrame(filas)


# =============================================================================
# Indicadores
# =============================================================================
def anios(base):
    r = _rechazos(base)
    hoy = dt.date.today().year
    if r.empty:
        return [hoy]
    return sorted({f.year for f in r["fecha"] if f} | {hoy}, reverse=True)


def indicadores(base, anio=None, hoy=None):
    hoy = hoy or dt.date.today()
    anio = anio or hoy.year
    r = _rechazos(base)
    vacio = pd.DataFrame(columns=["nombre", "tel", "asesor", "fseg", "dias"])
    if r.empty:
        return dict(anio=anio, rechazos=0, por_fin={}, ya_toca=0, al_dia=0, contactados=0, contestaron=0,
                    sin_contesta=0, registrados=0, sin_registro=0, saleu_total=0, sin_registro_rech=0,
                    sin_registro_aprob=0, aprob_sin_dato=0, vencidos=vacio, por_vencer=vacio,
                    pendientes=0, por_mes=pd.DataFrame(), por_asesor=pd.Series(dtype=int), revision=pd.DataFrame())
    del_anio = r[r["fecha"].map(lambda f: bool(f) and f.year == anio)]
    hecho = del_anio["hecho"].map(lambda v: v is not None)
    toca = del_anio["fseg"].map(lambda f: bool(f) and f <= hoy)
    contactados = del_anio[hecho]
    contestaron = int((contactados["contesta"] == "SI RESPONDE").sum())
    sin_contesta = int((contactados["contesta"] == "").sum())

    fin = del_anio["fin"].map(lambda f: f if f in ("CETELEM", "BBVA") else "Otras")
    por_mes = (pd.DataFrame({"mes": del_anio["fecha"].map(lambda f: f.month), "fin": fin})
               .groupby(["mes", "fin"]).size().rename("rechazos").reset_index())

    # Vencidos y por vencer: toda la hoja (como el Resumen del Sheet), sin primer seguimiento
    sin_hacer = r[r["hecho"].isna() & r["fseg"].notna()].copy()
    sin_hacer["dias"] = sin_hacer["fseg"].map(lambda f: (hoy - f).days)
    vencidos = sin_hacer[sin_hacer["dias"] >= 0].sort_values("dias", ascending=False)
    por_vencer = sin_hacer[(sin_hacer["dias"] < 0) & (sin_hacer["dias"] >= -15)].sort_values("dias", ascending=False)

    revision = pd.DataFrame()
    if TAB_REVISION in base.pestanas():
        rev = base.leer(TAB_REVISION)
        c_est = _col(rev, "Estado")
        if c_est:
            revision = rev[rev[c_est].map(lambda v: norm(v) == "PENDIENTE")].drop(columns="_fila", errors="ignore")
    return dict(
        anio=anio, rechazos=len(del_anio), por_fin=fin.value_counts().to_dict(),
        ya_toca=int(toca.sum()), al_dia=int((toca & hecho).sum()),
        contactados=len(contactados), contestaron=contestaron, sin_contesta=sin_contesta,
        **_registro_saleu(base, del_anio, anio),
        vencidos=vencidos, por_vencer=por_vencer, pendientes=len(revision), revision=revision, por_mes=por_mes,
        por_asesor=del_anio["asesor"].replace("", "(sin asesor)").value_counts())


def _aprobados(base, anio):
    """Aprobados del año en la Bitacora (Estatus final «Aprobado»): [(clave de nombre, SI / NO / "")]."""
    if TAB_BITACORA not in base.pestanas():
        return []
    bit = base.leer(TAB_BITACORA)
    c_est, c_cli, c_fecha = _col(bit, "Estatus final"), _col(bit, "Cliente"), _col(bit, "Fecha")
    c_su = _col(bit, "Registro en Sale-U", "Sale-U")
    if not (c_est and c_cli and c_fecha):
        return []
    out = []
    for _, x in bit.iterrows():
        f = to_date(x[c_fecha])
        if norm(x[c_est]) != "APROBADO" or not f or f.year != anio or _vacio(x[c_cli]):
            continue
        out.append((name_key(x[c_cli]), codigo_saleu(x[c_su]) if c_su else ""))
    return out


def _registro_saleu(base, del_anio, anio):
    """Registro en Sale-U por cliente (como la lista «Faltan en Sale-U»): Rechazados del año (MG Colima Rechazos,
    sin los que ya compraron) y Aprobados del año (Bitacora). Un cliente con varias filas cuenta una vez; está
    registrado si alguna fila lo dice. Si también está en Rechazos, cuenta como rechazo."""
    nivel = {"SI": 2, "CRED": 1}
    rech = {}                                         # clave de nombre → 2 registrado · 1 falta crédito · 0 falta alta
    for _, x in del_anio.iterrows():
        if "YA COMPRO" in x["com"]:
            continue
        k = name_key(x["nombre"])
        clave = next((c for c in rech if similar(k, c)), k)
        rech[clave] = max(rech.get(clave, 0), nivel.get(codigo_saleu(x["saleu"]), 0))
    aprob = {}
    for k, su in _aprobados(base, anio):
        if any(similar(k, c) for c in rech):
            continue
        clave = next((c for c in aprob if similar(k, c)), k)
        anterior = aprob.get(clave, "")
        aprob[clave] = next((v for v in ("SI", "CRED", "NO") if v in (su, anterior)), "")
    reg_r = sum(1 for v in rech.values() if v == 2)
    cred_r, alta_r = sum(1 for v in rech.values() if v == 1), sum(1 for v in rech.values() if v == 0)
    reg_a = sum(1 for v in aprob.values() if v == "SI")
    cred_a, alta_a = sum(1 for v in aprob.values() if v == "CRED"), sum(1 for v in aprob.values() if v == "NO")
    return dict(registrados=reg_r + reg_a, sin_registro=cred_r + alta_r + cred_a + alta_a,
                saleu_total=len(rech) + len(aprob), sin_registro_rech=cred_r + alta_r, sin_registro_aprob=cred_a + alta_a,
                falta_alta=alta_r + alta_a, falta_credito=cred_r + cred_a, aprob_sin_dato=len(aprob) - reg_a - cred_a - alta_a)


# =============================================================================
# Faltan en Sale-U
# =============================================================================
class _Clientes:
    """Agrupa registros de la misma persona: mismo celular (con nombre parecido) o el mismo nombre."""

    def __init__(self):
        self.grupos = []

    def grupo(self, nombre, tel):
        key = name_key(nombre)
        for g in self.grupos:
            if (tel and tel in g["tels"] and any(similar(key, k) or _palabras(key, k) for k in g["keys"])) \
                    or any(similar(key, k) for k in g["keys"]):
                g["keys"].add(key)
                if tel:
                    g["tels"].add(tel)
                return g
        g = dict(keys={key}, tels={tel} if tel else set(), regs=[])
        self.grupos.append(g)
        return g


def _palabras(a, b):
    pa, pb = set(a.split()), set(b.split())
    return bool(pa and pb) and len(pa & pb) >= min(len(pa), len(pb)) / 2


def _asesor_completo(base):
    from cuadres.actualizar import _catalogo_asesores, _asesor
    catalogo = _catalogo_asesores(base)
    return lambda v: ((_asesor(v, catalogo) or v) if catalogo and v else (v or "")).upper()


def _como_en_saleu(base):
    """{clave de nombre: (dónde, texto, Fuente del lead)} con lo último que anotó el cuadre de Sale-U en el Historial (sin reportes a
    la mano)."""
    from cuadres.base import TAB_HISTORIAL
    if TAB_HISTORIAL not in base.pestanas():
        return {}
    h = base.leer(TAB_HISTORIAL)
    if h.empty or "Origen" not in h.columns:
        return {}
    out = {}
    for _, r in h.iterrows():
        if str(r.get("Origen") or "") != "App · cuadre Sale-U" or "Sale-U" not in str(r.get("Campo") or ""):
            continue
        m = str(r.get("Motivo") or "")
        if m.startswith("Está en el reporte Fuente ("):
            dato = m[len("Está en el reporte Fuente ("):].split(") pero")[0]
            fuente = dato.split(", fuente ", 1)[1] if ", fuente " in dato else ""
            out[name_key(r["Cliente"])] = ("fuente", "Lead en Fuente: " + dato.split(", fuente ", 1)[0], fuente)
        elif m.startswith("No aparece"):
            out[name_key(r["Cliente"])] = ("no", "No aparecía en Sale-U por teléfono ni por nombre", "")
    return out


def _buscar_en_reportes(fuente, credito):
    """Función (nombre, tel) → (dónde: "credito" / "fuente" / "no" / None, texto de cómo está en Sale-U, Fuente del
    lead en Sale-U), con los reportes del último cuadre."""
    from cuadres import saleu as cu
    from cuadres.saleu_fuente import Buscador, describir_lead, describir_credito
    b = Buscador(fuente) if fuente is not None and len(fuente) else None
    cred = credito if credito is not None and len(credito) else None

    def buscar(nombre, tel):
        if cred is not None:
            nom, t = cu.normalizar_nombre(nombre), cu.normalizar_tel(tel or "")
            for _, s in cred.iterrows():
                if (t and s["Tel_norm"] == t) or (cu.similitud_nombre(nom, s["Nombre_norm"])[0] >= cu.UMBRAL_NOMBRE_FUERTE
                                                  and cu.similitud_nombre(nom, s["Nombre_norm"])[1] >= 2):
                    return "credito", "Crédito: " + describir_credito(
                        s["Nombre"], s["Telefono"], s["Asesor"], s["Estatus_SaleU"], s["Financiera"],
                        s["Fecha_Alta_Credito"]), " ".join(str(s.get("Fuente") or "").split())
        if b is not None:
            s, como = b.buscar(nombre, tel)
            return (("fuente", describir_lead(s, como), " ".join(str(s.get("Fuente") or "").split())) if s is not None
                    else ("no", "No aparece en Sale-U por teléfono ni por nombre", ""))
        return None, None, None
    return buscar


def faltantes_saleu(base, hoy=None, fuente=None, credito=None):
    """(faltan, por_confirmar): DataFrames con COLS_FALTAN, ordenados Aprobados primero y por vendedor.
    fuente / credito: DataFrames de los reportes del último cuadre (opcionales). Con ellos «Cómo está en Sale-U» y
    lo que falta salen de los reportes (con la misma regla del cuadre: un lead de Fuente → Registrar crédito; en
    Crédito → ya no se lista; nunca baja un «Falta crédito» del Sheet). Sin ellos, lo que dice el Sheet y lo que
    anotó el cuadre en el Historial."""
    hoy = hoy or dt.date.today()
    asesor = _asesor_completo(base)
    clientes = _Clientes()

    r = _rechazos(base)
    for _, x in r.iterrows() if not r.empty else []:
        if "YA COMPRO" in x["com"]:
            continue                                  # ya compró: es formalizado, no rechazado
        clientes.grupo(x["nombre"], x["tel"])["regs"].append(dict(
            estatus=RECHAZADO, saleu=codigo_saleu(x["saleu"]),
            cliente=x["nombre"], tel=x["tel"], vendedor=asesor(x["asesor"]), fin=x["fin"], unidad=x["unidad"],
            fecha=x["fecha"], mens=x["mens"], correo=x["correo"], fuente="MG Colima Rechazos", id=x["id"]))

    if TAB_BITACORA in base.pestanas():
        bit = base.leer(TAB_BITACORA)
        c = dict(est=_col(bit, "Estatus final"), cli=_col(bit, "Cliente"), tel=_col(bit, "Teléfono"),
                 ven=_col(bit, "Vendedor"), fin=_col(bit, "Financiera final", "Banco solicitud"),
                 uni=_col(bit, "Unidad"), fecha=_col(bit, "Fecha"), mens=_col(bit, "Mensualidad"),
                 correo=_col(bit, "Correo"), su=_col(bit, "Registro en Sale-U", "Sale-U"), fuente=_col(bit, "Fuente en la bitácora"))
        g = lambda x, k: x[c[k]] if c[k] else None
        for _, x in bit.iterrows() if c["est"] and c["cli"] else []:
            nombre = txt(g(x, "cli"))
            if not nombre:
                continue
            est = norm(g(x, "est"))
            su = codigo_saleu(g(x, "su"))
            if est != "APROBADO" and su != "SI":
                continue                              # solo interesan aprobados (y cualquier «sí está en Sale-U»)
            clientes.grupo(nombre, to_phone(g(x, "tel")))["regs"].append(dict(
                estatus=APROBADO if est == "APROBADO" else "otro", saleu=su,
                cliente=" ".join(nombre.split()).upper(), tel=to_phone(g(x, "tel")), vendedor=asesor(txt(g(x, "ven"))),
                fin=(txt(g(x, "fin")) or "").upper(), unidad=txt(g(x, "uni")) or "", fecha=to_date(g(x, "fecha")),
                mens=to_num(g(x, "mens")), correo=txt(g(x, "correo")) or "", fuente=txt(g(x, "fuente")) or "Bitacora",
                id=""))

    buscar = _buscar_en_reportes(fuente, credito) if fuente is not None or credito is not None else None
    historial = _como_en_saleu(base)
    faltan, confirmar = [], []
    for grupo in clientes.grupos:
        regs = grupo["regs"]
        if any(x["saleu"] == "SI" for x in regs):
            continue                                  # ya está en Sale-U según alguna fuente
        aprobados = [x for x in regs if x["estatus"] == APROBADO]
        rechazados = [x for x in regs if x["estatus"] == RECHAZADO]
        if aprobados:
            x = max(aprobados, key=lambda y: y["fecha"] or dt.date.min)
            destino = faltan if any(y["saleu"] in ("NO", "CRED") for y in aprobados) or any(
                y["saleu"] in ("NO", "CRED") for y in rechazados) else confirmar
        elif rechazados and any(y["saleu"] in ("NO", "CRED") for y in rechazados):
            x, destino = max(rechazados, key=lambda y: y["fecha"] or dt.date.min), faltan
        else:
            continue
        ids = ", ".join(dict.fromkeys(y["id"] for y in rechazados if y["id"]))
        donde, como, fuente_su = (buscar(x["cliente"], x["tel"]) if buscar
                                  else historial.get(name_key(x["cliente"]), (None, None, None)))
        if donde == "credito":
            continue                                  # ya tiene crédito en Sale-U (el Sheet aún no lo dice)
        cred = any(y["saleu"] == "CRED" for y in regs) or donde == "fuente"   # nunca baja de «Falta crédito»
        destino.append(dict(
            Falta=FALTA_CREDITO if cred else FALTA_ALTA,
            Estatus=x["estatus"], Cliente=x["cliente"], Telefono=telefono(x["tel"]), Vendedor=x["vendedor"] or "",
            Financiera=x["fin"], Unidad=x["unidad"], Fecha=x["fecha"], Mensualidad=x["mens"], Correo=x["correo"],
            Dias=(hoy - x["fecha"]).days if x["fecha"] else None,
            Origen=x["fuente"] + (" · antes rechazado" if x["estatus"] == APROBADO and rechazados else ""), ID=ids,
            Fuente=fuente_su or "",
            **{COL_SALEU: como or ""}))

    def tabla(filas):
        df = pd.DataFrame(filas, columns=COLS_FALTAN)
        if df.empty:
            return df
        df["_o"] = df["Estatus"].map({APROBADO: 0, RECHAZADO: 1})
        df["_f"] = df["Falta"].map({FALTA_ALTA: 0, FALTA_CREDITO: 1})
        return (df.sort_values(["_f", "_o", "Vendedor", "Fecha"], na_position="last").drop(columns=["_o", "_f"])
                .reset_index(drop=True))
    return tabla(faltan), tabla(confirmar)
