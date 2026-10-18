# -*- coding: utf-8 -*-
"""
Registro en Sale-U → columna "Registro en Sale-U" de MG Colima Rechazos (antes «Cliente con Registro en SalesU»)
y de los Aprobados de la pestaña Bitacora (mismo criterio).

Con el reporte de Sale-U (reporteLeadsconCredito) se revisa cada rechazo de la base cuya fecha
de ingreso cae dentro del periodo del reporte:
  · encontrado por teléfono o por nombre (basta uno de los dos)          → "Registrado"
  · no encontrado, pero sí está en el reporte Fuente (opcional)          → "Falta crédito" (falta registrar el crédito)
  · no está en ninguno                                                   → "No Registrado" (falta darlo de alta)
Nunca se baja de nivel (Registrado > Falta crédito > No Registrado): lo que ya dice que SÍ está registrado
(p. ej. marcado a mano) se deja así y se lista para revisar.
Los rechazos fuera del periodo del reporte no se tocan (no hay con qué compararlos).

Además (proponer_datos): si un cliente de Rechazos o de Créditos Finalizados cuadra con Sale-U o con la
bitácora por teléfono o por nombre, se llenan sus celdas vacías (teléfono, correo, folio, plazo,
mensualidad, asesor…, y en Finalizados el registro en Sale-U). Nunca se sobreescribe un dato.
"""

import datetime as dt
import re

import pandas as pd

from cuadres import saleu as cu
from cuadres.base import TAB_RECHAZOS, TAB_HISTORIAL, letra, norm, canonico
from cuadres.actualizar import leer_rechazos, C

REGISTRADO, FALTA_CREDITO, NO_REGISTRADO = "Registrado", "Falta crédito", "No Registrado"
VALORES = [REGISTRADO, FALTA_CREDITO, NO_REGISTRADO]


def es_falta_credito(valor):
    return norm(valor) in ("FALTA CREDITO", "FALTA CRÉDITO")


def ya_registrado(valor):
    """'Registrado', 'Sí', 'SI', 'Si registrado'… (no 'No Registrado')."""
    n = norm(valor)
    return n in ("REGISTRADO", "REGISTRADA") or n == "SI" or n.startswith("SI ")


# =============================================================================
# El Sheet (MG Colima Rechazos + Créditos Finalizados) como base del cuadre de Sale-U
# =============================================================================
ESTATUS_FINALIZADO = "FINALIZADO"
COL_FECHA_ENTREGA = "FECHA DE ENTREGA"     # en Créditos Finalizados; con fecha, entra al cuadre solo en su periodo


def solicitudes_del_sheet(base):
    """Clientes de «MG Colima Rechazos» (todas las financieras) y «MG Colima Créditos Finalizados» con la forma de
    fuentes.solicitudes_de_bitacora(), para cuadrarlos contra Sale-U. Los de Finalizados entran por su FECHA DE
    ENTREGA; los que no la tienen entran siempre al cuadre (columna _siempre), sin importar el periodo."""
    from cuadres.fuentes import (solicitudes_de_bitacora, to_folio, to_phone, to_num, to_date, txt, mail, vacio,
                                 ETIQUETA_MES)
    from cuadres.base import TAB_FINALIZADOS
    from cuadres.ids import COL_ID
    cols = list(solicitudes_de_bitacora(None).columns) + ["Hoja", "Fila", "ID", "_siempre"]
    filas = []

    def fila(hoja, n, nombre, tel, fecha=None, folio=None, vendedor=None, unidad=None, fin=None, estatus="",
             plazo=None, mens=None, correo=None, siempre=False, id_=""):
        f = to_date(fecha)
        fin = (txt(fin) or "").upper() or "SIN FINANCIERA"
        tel, folio, plazo, mens = to_phone(tel), to_folio(folio), to_num(plazo), to_num(mens)
        filas.append(dict(
            Folio=str(folio) if folio else "", Nombre=" ".join(str(nombre).split()).upper(),
            Telefono=str(tel) if tel else "", Vendedor=(txt(vendedor) or "").upper() or "(sin vendedor)",
            Vehiculo=txt(unidad) or "", Monto_Financiar="", Fecha=f,
            Mes=f"{ETIQUETA_MES[f.month][:3]} {f.year}" if f else "SIN FECHA",
            _anio=f.year if f else 0, _mes=f.month if f else 0, Est_sol=estatus, Fin_sol=fin, Est_fin=estatus,
            Fin_fin=fin, Fuente=f"{hoja} fila {n}", Plazo=int(plazo) if plazo else None,
            Mensualidad=round(mens, 2) if mens else None, Correo=mail(correo) or "",
            Hoja=hoja, Fila=n, ID=str(id_ or "").strip(), _siempre=siempre))

    pestanas = set(base.pestanas())
    if TAB_RECHAZOS in pestanas:
        df = base.leer(TAB_RECHAZOS)
        col = {k: next((c for c in df.columns if norm(c) == norm(v)), None) for k, v in C.items()}
        c_id = next((c for c in df.columns if norm(c) == norm(COL_ID)), None)
        for _, r in df.iterrows():
            g = lambda k: r[col[k]] if col[k] else None
            if vacio(g("nombre")) or not str(g("nombre")).strip():
                continue
            fila("Rechazos", int(r["_fila"]), g("nombre"), g("cel"), g("fecha"), g("folio"), g("asesor"),
                 g("unidad"), g("fin"), norm(g("status")) or "RECHAZADO", g("plazo"), g("mens"), g("correo"),
                 id_=r[c_id] if c_id else "")
    if TAB_FINALIZADOS in pestanas:
        df = base.leer(TAB_FINALIZADOS)
        col = lambda n: next((c for c in df.columns if norm(c) == n), None)
        c_nom, c_tel, c_cor, c_ase = col("NOMBRE"), col("TELEFONO"), col("CORREO"), col("ASESOR")
        c_fec = col(norm(COL_FECHA_ENTREGA))
        for _, r in df.iterrows() if c_nom else []:
            if vacio(r[c_nom]) or not str(r[c_nom]).strip():
                continue
            entrega = to_date(r[c_fec]) if c_fec else None
            fila("Finalizados", int(r["_fila"]), r[c_nom], r[c_tel] if c_tel else None, fecha=entrega,
                 vendedor=r[c_ase] if c_ase else None, correo=r[c_cor] if c_cor else None,
                 estatus=ESTATUS_FINALIZADO, siempre=entrega is None)
    return pd.DataFrame(filas, columns=cols)


def periodo(su):
    """Periodo del reporte: el «Rango» de su encabezado (el mismo que usa el cuadre); si no lo trae,
    de la primera a la última fecha de alta de crédito."""
    rango = su.attrs.get("rango")
    if rango:
        return rango[0].date(), rango[1].date()
    fechas = su["_fecha"].dropna()
    if fechas.empty:
        return None, None
    return fechas.min().date(), fechas.max().date()


def _aprobados_bitacora(base):
    """Aprobados de la pestaña Bitacora con la forma de leer_rechazos() (hoja, fila, fecha, nombre, cel, saleu):
    su Registro en Sale-U se cuadra con el reporte igual que el de los rechazos."""
    from cuadres.base import TAB_BITACORA
    from cuadres.fuentes import COL_SALEU_BITACORA, to_date, to_phone
    if TAB_BITACORA not in base.pestanas():
        return []
    df = base.leer(TAB_BITACORA)
    col = lambda n: next((c for c in df.columns if norm(c) == norm(n)), None)
    c_est, c_nom, c_tel, c_fec, c_su = (col("Estatus final"), col("Cliente"), col("Teléfono"), col("Fecha"),
                                        col(COL_SALEU_BITACORA))
    if not (c_est and c_nom and c_su):
        return []                      # Bitacora con el formato anterior (sin «Registro en Sale-U»)
    out = []
    for _, r in df.iterrows():
        if not norm(r[c_est]).startswith("APROBAD") or not r[c_nom]:
            continue
        out.append(dict(hoja=TAB_BITACORA, fila=int(r["_fila"]), fecha=to_date(r[c_fec]) if c_fec else None,
                        nombre=" ".join(str(r[c_nom]).split()), cel=to_phone(r[c_tel]) if c_tel else None,
                        saleu=r[c_su]))
    return out


def proponer(base, su, ini=None, fin=None, fuente=None):
    """su: DataFrame de cu.leer_saleu(); fuente: DataFrame de saleu_fuente.leer_fuente() (opcional). Devuelve
    DataFrame con lo que se escribiría: los rechazos de MG Colima Rechazos y los Aprobados de la Bitacora (columna
    Hoja) cuya fecha cae en el periodo del reporte de Crédito."""
    from cuadres.saleu_fuente import Buscador, alta_txt
    buscador = Buscador(fuente) if fuente is not None and len(fuente) else None
    p_ini, p_fin = periodo(su)
    ini, fin = ini or p_ini, fin or p_fin
    if not ini:
        raise ValueError("El reporte de Sale-U no trae fechas de alta de crédito.")
    por_tel = {}
    for _, s in su.iterrows():
        if s["Tel_norm"]:
            por_tel.setdefault(s["Tel_norm"], []).append(s)
    filas = []
    for row in leer_rechazos(base) + _aprobados_bitacora(base):
        nom = cu.normalizar_nombre(row["nombre"])
        tel = str(row["cel"]) if row["cel"] else ""
        if not row["fecha"] or not (ini <= row["fecha"] <= fin):
            # fuera del periodo del reporte de Crédito no se sabe si tiene crédito; solo un «No Registrado» que
            # sí está en Fuente sube a «Falta crédito» (ya es lead)
            if buscador is None or norm(row["saleu"]) != norm(NO_REGISTRADO):
                continue
            lead, como_lead = buscador.buscar(row["nombre"], tel)
            if lead is not None:
                filas.append(dict(Hoja=row.get("hoja", TAB_RECHAZOS), Fila=row["fila"], Cliente=row["nombre"],
                                  Fecha_ingreso=row["fecha"], Antes=row["saleu"], Después=FALTA_CREDITO, Cambia=True,
                                  Como=f"Fuente, por {como_lead} (fuera del periodo de Crédito)", Revisar="",
                                  Cliente_SaleU="", Estatus_SaleU="", Financiera_SaleU="",
                                  Cliente_Fuente=" ".join(lead["Nombre"].split()), Alta_Fuente=alta_txt(lead)))
            continue
        hallado, como = None, ""
        if por_tel.get(tel):                    # mismo teléfono: basta (si el nombre no se parece, se avisa)
            s = max(por_tel[tel], key=lambda x: cu.similitud_nombre(nom, x["Nombre_norm"]))
            r, _ = cu.similitud_nombre(nom, s["Nombre_norm"])
            hallado, como = s, ("teléfono" if r >= cu.UMBRAL_NOMBRE_CON_TEL else "teléfono (nombre distinto)")
        if hallado is None:
            mejor = (0, 0, None)
            for _, s in su.iterrows():
                r, k = cu.similitud_nombre(nom, s["Nombre_norm"])
                if (r, k) > mejor[:2]:
                    mejor = (r, k, s)
            if mejor[0] >= cu.UMBRAL_NOMBRE_FUERTE and mejor[1] >= 2:
                hallado, como = mejor[2], "nombre (teléfono distinto)"
        actual = row["saleu"] or ""
        revisar, lead = "", None
        if hallado is None and buscador is not None:
            lead, como_lead = buscador.buscar(row["nombre"], tel)
        if hallado is not None:
            nuevo = REGISTRADO
        elif ya_registrado(actual):
            nuevo = REGISTRADO          # lo marcaron como registrado: se respeta (solo se uniforma la palabra)
            revisar = ("Marcado como registrado; está en Fuente pero no en el reporte de Crédito" if lead is not None
                       else "Marcado como registrado, pero no aparece en este reporte de Sale-U")
            como = "se respetó lo que ya decía"
        elif lead is not None:
            nuevo, como = FALTA_CREDITO, f"Fuente, por {como_lead}"
        elif es_falta_credito(actual):
            nuevo, como = FALTA_CREDITO, "se respetó lo que ya decía"   # ya se vio en Fuente antes
        else:
            nuevo = NO_REGISTRADO
        filas.append(dict(
            Hoja=row.get("hoja", TAB_RECHAZOS), Fila=row["fila"], Cliente=row["nombre"], Fecha_ingreso=row["fecha"], Antes=actual, Después=nuevo,
            Cambia=actual != nuevo, Como=como, Revisar=revisar,
            Cliente_SaleU=hallado["Nombre"] if hallado is not None else "",
            Estatus_SaleU=hallado["Estatus_SaleU"] if hallado is not None else "",
            Financiera_SaleU=hallado["Financiera"] if hallado is not None else "",
            Cliente_Fuente=" ".join(lead["Nombre"].split()) if lead is not None else "",
            Alta_Fuente=alta_txt(lead) if lead is not None else ""))
    df = pd.DataFrame(filas, columns=COLS_PROPUESTA)
    return df, (ini, fin)


COLS_PROPUESTA = ["Hoja", "Fila", "Cliente", "Fecha_ingreso", "Antes", "Después", "Cambia", "Como", "Revisar",
                  "Cliente_SaleU", "Estatus_SaleU", "Financiera_SaleU", "Cliente_Fuente", "Alta_Fuente"]


def _motivo(r, periodo_txt):
    if r["Revisar"] and r["Después"] == REGISTRADO:
        return f"Ya decía que estaba registrado; se uniformó la palabra (no aparece en el reporte {periodo_txt})"
    if r["Después"] == REGISTRADO:
        return f"Encontrado en Sale-U por {r['Como']}: {r['Cliente_SaleU']} ({r['Estatus_SaleU']})"
    if r["Después"] == FALTA_CREDITO:
        if r.get("Cliente_Fuente"):
            return (f"Está en el reporte Fuente ({r['Cliente_Fuente']}, alta {r.get('Alta_Fuente', '')}) pero no en "
                    f"el de Crédito ({periodo_txt}): falta registrar el crédito")
        return "Ya decía Falta crédito"
    return f"No aparece en el reporte de Sale-U ({periodo_txt})"


def proponer_datos(base, su):
    """Celdas vacías de Rechazos y Finalizados que se llenan porque el cliente cuadra con Sale-U o con la
    bitácora (pestaña Bitacora) por teléfono o nombre. Si las dos traen el dato, manda la bitácora
    (la mensualidad de Sale-U es aproximada)."""
    from cuadres import completar
    from cuadres.base import TAB_BITACORA
    from cuadres.actualizar import _catalogo_asesores, _asesor
    catalogo = _catalogo_asesores(base)
    nombre = (lambda v: _asesor(v, catalogo)) if catalogo else None
    regs = completar.registros_de_saleu(su, nombre)
    if TAB_BITACORA in base.pestanas():
        regs += completar.registros_de_bitacora(base.leer(TAB_BITACORA), nombre)
    return completar.proponer(base, regs)


def aplicar_datos(base, datos):
    """Escribe lo de proponer_datos() y lo anota en el Historial."""
    if not datos:
        return 0
    por_hoja = {}
    for d in datos:
        c = base.col(d["Hoja"], d["Campo"])
        if c:
            por_hoja.setdefault(d["Hoja"], []).append((d, c))
    for hoja, items in por_hoja.items():
        base.respaldar(hoja)
        base.escribir(hoja, [(int(d["Fila"]), c, d["_valor"]) for d, c in items])
    base.anotar_historial([dict(Origen="App · cuadre Sale-U", Hoja=hoja, Celda=f"{letra(c)}{int(d['Fila'])}",
                                Cliente=d["Cliente"], Campo=d["Campo"], Antes="", Después=d["Después"],
                                Motivo=d["Motivo"])
                           for hoja, items in por_hoja.items() for d, c in items])
    return sum(map(len, por_hoja.values()))


def aplicar(base, cambios, periodo_txt):
    """Escribe solo las filas cuyo valor cambia y las anota en el Historial."""
    from cuadres.base import TAB_BITACORA
    from cuadres.fuentes import COL_SALEU_BITACORA
    cambios = cambios[cambios["Cambia"]]
    if cambios.empty:
        return 0
    if "Hoja" not in cambios.columns:
        cambios = cambios.assign(Hoja=TAB_RECHAZOS)
    nombre_col = {TAB_RECHAZOS: C["saleu"], TAB_BITACORA: COL_SALEU_BITACORA}
    cols = {}
    for hoja, g in cambios.groupby("Hoja"):
        col = base.col(hoja, nombre_col[hoja])
        if not col:
            raise ValueError(f"No encuentro la columna '{nombre_col[hoja]}' en '{hoja}'.")
        cols[hoja] = col
        if hoja == TAB_RECHAZOS:
            base.respaldar(hoja)        # la Bitacora se rehace sola con las bitácoras: no necesita respaldo
        base.escribir(hoja, [(int(r["Fila"]), col, r["Después"]) for _, r in g.iterrows()])
    base.anotar_historial([dict(
        Origen="App · cuadre Sale-U", Hoja=r["Hoja"], Celda=f"{letra(cols[r['Hoja']])}{int(r['Fila'])}",
        Cliente=r["Cliente"], Campo=nombre_col[r["Hoja"]], Antes=r["Antes"], Después=r["Después"],
        Motivo=_motivo(r, periodo_txt))
        for _, r in cambios.iterrows()])
    estilo_registro(base, set(cols))
    return len(cambios)


def estilo_registro(base, hojas=None):
    """Lista Registrado / Falta crédito / No Registrado (más «No aplica» en la Bitacora) y el color de «Falta
    crédito» en la columna Registro en Sale-U de Rechazos y Bitacora. No duplica reglas."""
    from cuadres.base import TAB_BITACORA
    from cuadres.fuentes import COL_SALEU_BITACORA
    from cuadres.actualizar import ESTILO_SALEU, NO_APLICA
    for hoja, nombre in ((TAB_RECHAZOS, C["saleu"]), (TAB_BITACORA, COL_SALEU_BITACORA)):
        if (hojas is not None and hoja not in hojas) or hoja not in base.pestanas():
            continue
        c = base.col(hoja, nombre)
        if not c:
            continue
        L = letra(c)
        for valor, fondo, tinta in ESTILO_SALEU:
            if hoja == TAB_BITACORA or valor != f'="{NO_APLICA}"':
                base.regla_color(hoja, [c], f"=${L}2{valor}", fondo, 1000, tinta)
        base.lista_valores(hoja, [c], VALORES + ([NO_APLICA] if hoja == TAB_BITACORA else []), 1000)


# =============================================================================
# Restaurar lo que una versión anterior bajó a "No Registrado"
# =============================================================================
def bajas_por_error(base):
    """Filas que el cuadre de Sale-U cambió de Registrado/Sí a "No Registrado" (versión anterior de la app)
    y que siguen así. Devuelve DataFrame con Fila, Columna, Celda, Cliente, Antes, Fecha."""
    from cuadres.fuentes import name_key, similar
    from cuadres.pendientes import _celda
    cols = ["Fila", "Columna", "Celda", "Cliente", "Antes", "Fecha"]
    if TAB_HISTORIAL not in base.pestanas():
        return pd.DataFrame(columns=cols)
    todo = base.leer(TAB_HISTORIAL)
    if todo.empty or "Origen" not in todo.columns:
        return pd.DataFrame(columns=cols)
    origen = todo["Origen"].astype(str)
    h = todo[(origen == "App · cuadre Sale-U") & (todo["Después"].astype(str) == NO_REGISTRADO)
             & (todo["Hoja"].astype(str) == TAB_RECHAZOS)
             & todo["Antes"].map(ya_registrado)]
    if h.empty:
        return pd.DataFrame(columns=cols)
    restaurados = set(todo.loc[origen == "App · restaurar Sale-U", "Celda"].astype(str))
    actuales = {r["fila"]: r for r in leer_rechazos(base)}
    filas, vistas = [], set()
    for _, r in h.iloc[::-1].iterrows():                  # la anotación más reciente de cada celda
        ref = _celda(r["Celda"])
        if not ref or str(r["Celda"]) in vistas:
            continue
        vistas.add(str(r["Celda"]))
        fila, col = ref
        act = actuales.get(fila)
        if (act is None or (act["saleu"] or "") != NO_REGISTRADO or str(r["Celda"]) in restaurados
                or not similar(name_key(act["nombre"]), name_key(r["Cliente"]))):
            continue                                       # ya cambió, se movió la fila o ya se restauró
        f = r["Fecha"]
        filas.append(dict(Fila=fila, Columna=col, Celda=str(r["Celda"]), Cliente=r["Cliente"], Antes=r["Antes"],
                          Fecha=f.strftime("%d/%m/%Y %H:%M") if isinstance(f, (dt.date, dt.datetime)) else str(f or "")))
    return pd.DataFrame(filas, columns=cols)


def restaurar(base, bajas):
    """Regresa a "Registrado" las filas de bajas_por_error() y lo anota en el Historial."""
    if bajas.empty:
        return 0
    base.respaldar(TAB_RECHAZOS)
    base.escribir(TAB_RECHAZOS, [(int(r["Fila"]), int(r["Columna"]), REGISTRADO) for _, r in bajas.iterrows()])
    base.anotar_historial([dict(
        Origen="App · restaurar Sale-U", Hoja=TAB_RECHAZOS, Celda=r["Celda"], Cliente=r["Cliente"], Campo=C["saleu"],
        Antes=NO_REGISTRADO, Después=REGISTRADO,
        Motivo=f"Decía «{r['Antes']}» y el cuadre de Sale-U del {r['Fecha']} lo bajó a No Registrado (versión anterior)")
        for _, r in bajas.iterrows()])
    return len(bajas)


# =============================================================================
# Deshacer lo que se escribió por un teléfono compartido (regla anterior: bastaba una palabra en común)
# =============================================================================
_RE_REGISTRO = re.compile(r"^Encontrado en Sale-U por teléfono: (.+) \([^()]*\)$")
_RE_DATO = re.compile(r"cuadra por teléfono con Sale-U \((.+)\)$")


def por_telefono_dudosos(base):
    """Celdas que el cuadre de Sale-U escribió solo porque el teléfono coincidía con alguien de nombre distinto
    (menos del 50 % del nombre igual) y que siguen con ese valor. DataFrame con lo necesario para deshacer."""
    from cuadres.fuentes import name_key, similar
    from cuadres.pendientes import _celda
    cols = ["Hoja", "Fila", "Columna", "Celda", "Cliente", "Campo", "Valor", "Regresar a", "Sale-U", "Fecha"]
    if TAB_HISTORIAL not in base.pestanas():
        return pd.DataFrame(columns=cols)
    h = base.leer(TAB_HISTORIAL)
    if h.empty or "Origen" not in h.columns:
        return pd.DataFrame(columns=cols)
    deshechas = {(r["Hoja"], r["Celda"]) for _, r in h[h["Origen"].astype(str) == "App · deshacer Sale-U"].iterrows()}
    h = h[h["Origen"].astype(str) == "App · cuadre Sale-U"]
    hojas, filas, vistas = {}, [], set(deshechas)
    for _, r in h.iloc[::-1].iterrows():
        motivo = str(r["Motivo"] or "")
        m = _RE_REGISTRO.match(motivo) or _RE_DATO.search(motivo)
        ref = _celda(r["Celda"])
        if not m or not ref or (r["Hoja"], r["Celda"]) in vistas:
            continue
        vistas.add((r["Hoja"], r["Celda"]))
        nom_su = m.group(1)
        sim, _ = cu.similitud_nombre(cu.normalizar_nombre(r["Cliente"]), cu.normalizar_nombre(nom_su))
        if sim >= cu.UMBRAL_NOMBRE_CON_TEL or str(r["Antes"] or "") == str(r["Después"] or ""):
            continue                                    # con la regla actual también cuadra, o no cambió nada
        hoja = str(r["Hoja"])
        if hoja not in base.pestanas():
            continue
        if hoja not in hojas:
            hojas[hoja] = {int(x["_fila"]): x for _, x in base.leer(hoja).iterrows()}
        fila, col = ref
        actual = hojas[hoja].get(fila)
        if actual is None:
            continue
        nombre_col = next((c for c in actual.index if norm(c) in ("NOMBRE DE CLIENTE", "NOMBRE")), None)
        campo_col = next((c for c in actual.index if norm(c) == norm(canonico(r["Campo"] or ""))), None)
        if (not campo_col or str(actual[campo_col] if actual[campo_col] is not None else "") != str(r["Después"])
                or (nombre_col and not similar(name_key(actual[nombre_col]), name_key(r["Cliente"])))):
            continue                                    # ya cambió o la fila se movió: no se toca
        f = r["Fecha"]
        filas.append({"Hoja": hoja, "Fila": fila, "Columna": col, "Celda": str(r["Celda"]), "Cliente": r["Cliente"],
                      "Campo": r["Campo"], "Valor": r["Después"], "Regresar a": r["Antes"] or "",
                      "Sale-U": nom_su,
                      "Fecha": f.strftime("%d/%m/%Y %H:%M") if isinstance(f, (dt.date, dt.datetime)) else str(f or "")})
    return pd.DataFrame(filas, columns=cols)


def deshacer_dudosos(base, dudosos):
    if dudosos.empty:
        return 0
    for hoja, g in dudosos.groupby("Hoja"):
        base.respaldar(hoja)
        base.escribir(hoja, [(int(r["Fila"]), int(r["Columna"]), r["Regresar a"] or None) for _, r in g.iterrows()])
    base.anotar_historial([dict(
        Origen="App · deshacer Sale-U", Hoja=r["Hoja"], Celda=r["Celda"], Cliente=r["Cliente"], Campo=r["Campo"],
        Antes=r["Valor"], Después=r["Regresar a"],
        Motivo=f"Se había escrito porque el teléfono coincide con {r['Sale-U']} en Sale-U, pero el nombre es otro")
        for _, r in dudosos.iterrows()])
    return len(dudosos)
