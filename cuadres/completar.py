# -*- coding: utf-8 -*-
"""
Completar datos faltantes de la base.

Si un cliente de "MG Colima Rechazos" o "MG Colima Créditos Finalizados" cuadra con un registro de
otra fuente (bitácoras de Cetelem / F&I o el reporte de Sale-U) por **teléfono** o por **nombre**,
las celdas que estén VACÍAS se llenan con el dato de esa fuente. Nunca se sobreescribe un dato.

  · Por teléfono: mismo celular (10 dígitos) y nombre parecido (al menos la mitad de las palabras), así
    un familiar con el mismo teléfono (p. ej. mismo apellido) no mezcla sus datos.
  · Por nombre: el mismo nombre sin importar orden, acentos ni puntos (≥ 90 % de parecido).

registros: lista de dicts con las llaves de CAMPOS que tenga la fuente, más
  nombre, tel, fecha, fin (financiera) y origen (texto para el Historial).
"""
import re

from cuadres.base import TAB_RECHAZOS, TAB_FINALIZADOS, norm
from cuadres.fuentes import name_key, similar, to_phone, to_folio, to_num, to_date, financiera_igual, vacio

# Columnas que se pueden completar en cada pestaña (llave → encabezado)
COLUMNAS = {
    TAB_RECHAZOS: dict(nombre="Nombre de Cliente", fecha="Fecha ingreso", tel="Celular", folio="Folio",
                       fin="Financiera", correo="Correo", plazo="Plazo", mens="Mensualidad", unidad="Unidad",
                       asesor="Asesor de Ventas", gap="GAP", acc="Accesorios", perfil="Perfil"),
    # en Rechazos el registro en Sale-U lo maneja el cuadre de Sale-U (con su periodo); aquí solo Finalizados
    TAB_FINALIZADOS: dict(nombre="NOMBRE", tel="TELEFONO", correo="CORREO", saleu="CLIENTE CON REGISTRO EN SALE-U"),
}
CAMPOS = ["tel", "correo", "folio", "plazo", "mens", "unidad", "asesor", "gap", "acc", "perfil", "saleu"]


def _palabras(key):
    return set((key or "").split())


def filas_base(base):
    """Filas de las pestañas de la base con sus valores actuales."""
    out = []
    pestanas = set(base.pestanas())
    for hoja, cols in COLUMNAS.items():
        if hoja not in pestanas:
            continue
        df = base.leer(hoja)
        enc = {k: next((c for c in df.columns if norm(c) == norm(v)), None) for k, v in cols.items()}
        if not enc["nombre"]:
            continue
        for _, r in df.iterrows():
            nombre = r[enc["nombre"]]
            if vacio(nombre) or not str(nombre).strip():
                continue
            valores = {k: r[c] for k, c in enc.items() if c and k not in ("nombre",)}
            out.append(dict(hoja=hoja, fila=int(r["_fila"]), nombre=str(nombre).strip(), key=name_key(nombre),
                            tel=to_phone(valores.get("tel")), fecha=to_date(valores.get("fecha")),
                            fin=norm(valores.get("fin")) or None, valores=valores,
                            columnas={k: c for k, c in enc.items() if c}))
    return out


def _candidatos(fila, por_tel, registros):
    """Registros que cuadran con la fila, del más confiable al menos."""
    vistos, cands = set(), []
    from cuadres import saleu as cu
    nom = cu.normalizar_nombre(fila["nombre"])
    for r in por_tel.get(fila["tel"], []) if fila["tel"] else []:
        # mismo teléfono y nombre parecido (≥ 50 % de las palabras): un familiar con el mismo teléfono no cuenta
        if cu.similitud_nombre(nom, cu.normalizar_nombre(r.get("nombre")))[0] >= cu.UMBRAL_NOMBRE_CON_TEL:
            cands.append((0 if similar(fila["key"], r["key"]) else 1, r, "teléfono"))
            vistos.add(id(r))
    for r in registros:
        if id(r) not in vistos and similar(fila["key"], r["key"]):
            cands.append((2, r, "nombre"))

    def orden(c):
        nivel, r, _ = c
        misma_fin = bool(fila["fin"] and r.get("fin") and financiera_igual(fila["fin"], r["fin"]))
        dias = abs((fila["fecha"] - r["fecha"]).days) if fila["fecha"] and r.get("fecha") else 9999
        return (nivel, r.get("prioridad", 0), not misma_fin, dias)
    return sorted(cands, key=orden)


def _valor_para(campo, fila, r):
    """Valor del registro para ese campo, o None si no aplica a esta fila."""
    v = r.get(campo)
    if vacio(v):
        return None
    if campo == "folio":
        # un folio CCK solo es de Cetelem, y solo si la fila también es de Cetelem
        if not (financiera_igual(r.get("fin") or "", "CETELEM") and financiera_igual(fila["fin"] or "", "CETELEM")):
            return None
        return to_folio(v)
    if campo == "tel":
        return to_phone(v)
    if campo in ("plazo",):
        n = to_num(v)
        return int(n) if n else None
    if campo == "mens":
        n = to_num(v)
        return round(n, 2) if n else None
    if campo in ("gap", "acc"):
        return v if norm(v) in ("SI", "NO") else None
    if campo == "saleu":
        return v if v in ("Registrado", "No Registrado") else None
    if campo == "perfil":
        return "MORAL" if norm(v) == "MORAL" else None
    if campo == "correo":
        s = str(v).strip().lower()
        return s if re.fullmatch(r"[^@\s]+@[^@\s]+\.[^@\s]+", s) else None
    s = str(v).strip()
    return s or None


ETIQUETA = {"tel": "el teléfono", "correo": "el correo", "folio": "el folio", "plazo": "el plazo",
            "mens": "la mensualidad", "unidad": "la unidad", "asesor": "el asesor", "gap": "el GAP",
            "acc": "los accesorios", "perfil": "el perfil", "saleu": "el registro en Sale-U"}


def proponer(base, registros, campos=CAMPOS, filas=None):
    """Lista de dicts (Hoja, Fila, Cliente, Campo, Antes, Después, Motivo, _valor) para las celdas vacías."""
    for r in registros:
        r.setdefault("key", name_key(r.get("nombre")))
        r["tel"] = to_phone(r.get("tel"))
    por_tel = {}
    for r in registros:
        if r["tel"]:
            por_tel.setdefault(r["tel"], []).append(r)
    out = []
    for fila in filas if filas is not None else filas_base(base):
        vacios = [c for c in campos if c in fila["columnas"] and vacio(fila["valores"].get(c))]
        if not vacios:
            continue
        cands = _candidatos(fila, por_tel, registros)
        if not cands:
            continue
        for campo in vacios:
            # el registro en Sale-U lo dice mejor el propio reporte de Sale-U que la bitácora
            lista = sorted(cands, key=lambda c: -c[1].get("prioridad", 0)) if campo == "saleu" else cands
            for _, r, como in lista:
                v = _valor_para(campo, fila, r)
                if v is None:
                    continue
                out.append(dict(Hoja=fila["hoja"], Fila=fila["fila"], Cliente=fila["nombre"],
                                Campo=fila["columnas"][campo], Antes="", Después=v,
                                Motivo=f"Faltaba {ETIQUETA[campo]}; cuadra por {como} con {r['origen']}",
                                _valor=v))
                break
    return out


# =============================================================================
# Fuentes → registros
# =============================================================================
def registros_de_solicitudes(solicitudes, asesor=None):
    """Solicitudes consolidadas de las bitácoras (fuentes.consolidar). asesor: función nombre corto → completo."""
    out = []
    for s in solicitudes:
        for key in s["keys"] or {s["key"]}:
            out.append(dict(
                nombre=s["cliente"], key=key, tel=s["telefono"], fecha=s["fecha"],
                fin=s["banco"] if norm(s["banco"]) not in ("", "OTROS") else s.get("canal2"),
                folio=s["folio"], correo=s["correo"], plazo=s["plazo"], mens=s["mensualidad"], unidad=s["unidad"],
                asesor=asesor(s["vendedor"]) if asesor and s["vendedor"] else s["vendedor"],
                gap=s["gap"], acc=s["accesorios"], perfil=s.get("tipo_persona"),
                saleu={"SI": "Registrado", "NO": "No Registrado"}.get(s.get("sale_u")),
                origen=f"la bitácora ({(s['fuentes'] or '').split(',')[0]})"))
    return out


def registros_de_bitacora(df, asesor=None):
    """Pestaña Bitacora de la base (BITACORA_COLS)."""
    out = []
    if df is None or df.empty:
        return out
    g = lambda r, c: r[c] if c in r.index else None
    for _, r in df.iterrows():
        nombre = g(r, "Cliente")
        if vacio(nombre):
            continue
        banco = g(r, "Banco solicitud")
        ven = g(r, "Vendedor")
        out.append(dict(
            nombre=str(nombre), tel=g(r, "Teléfono"), fecha=to_date(g(r, "Fecha")),
            fin=banco if norm(banco) not in ("", "OTROS") else g(r, "Canal 2"), folio=g(r, "Folio"),
            correo=g(r, "Correo"), plazo=g(r, "Plazo"), mens=g(r, "Mensualidad"), unidad=g(r, "Unidad"),
            asesor=asesor(ven) if asesor and not vacio(ven) else ven,
            saleu={"SI": "Registrado", "NO": "No Registrado"}.get(norm(g(r, "Sale-U"))),
            origen=f"la bitácora ({str(g(r, 'Fuente en la bitácora') or '').split(',')[0]})"))
    return out


def asesor_saleu(v):
    """'JessicaElizabethMachucaRoque36' / 'Evelia_Lizette_Acosta_Herrera8765' → 'Jessica Elizabeth Machuca Roque'."""
    s = re.sub(r"\d+$", "", str(v or "")).replace("_", " ")
    s = re.sub(r"(?<=[a-záéíóúñ])(?=[A-ZÁÉÍÓÚÑ])", " ", s)
    return re.sub(r"\s+", " ", s).strip()


def registros_de_saleu(su, asesor=None):
    """Reporte de Sale-U (saleu.leer_saleu)."""
    out = []
    for _, s in su.iterrows():
        f = s.get("_fecha")
        out.append(dict(
            nombre=s["Nombre"], tel=s["Tel_norm"] or s["Telefono"], fecha=f.date() if hasattr(f, "date") and f == f else None,
            fin=s["Financiera"] or None, plazo=s.get("Plazos_SaleU"), mens=s.get("Mensualidad_SaleU"),
            asesor=(asesor or (lambda x: x))(asesor_saleu(s["Asesor"])) if s["Asesor"] else None,
            saleu="Registrado", prioridad=1,          # si la bitácora trae el dato, manda la bitácora
            origen=f"Sale-U ({s['Nombre']})"))
    return out
