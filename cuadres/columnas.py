# -*- coding: utf-8 -*-
"""
Orden de las primeras columnas de las hojas de captura, para tener fecha, nombre y teléfono a la vista (congelados):

  · MG Colima Rechazos: Fecha ingreso | Nombre de Cliente | Celular | Folio …        (congela A–C)
  · Bitacora:           Fecha | Cliente | Teléfono | Folio …                          (congela A–C)
  · Créditos Finalizados: sin la columna MARCA; NOMBRE | TELEFONO …                   (congela A–B)

  reacomodar(base) → lista de lo que hizo. Mueve columnas completas (datos, formato, validaciones, colores y
  fórmulas se ajustan), así que no se pierde nada; si ya están en orden no hace nada. También corrige las celdas
  que apuntan los pendientes de «Revision» a las nuevas letras.
"""
from cuadres.base import TAB_RECHAZOS, TAB_FINALIZADOS, TAB_BITACORA, cambiar_referencias, norm

ORDEN = {
    TAB_RECHAZOS: dict(primeras=["Fecha ingreso", "Nombre de Cliente", "Celular", "Folio"], congelar=3, borrar=[]),
    TAB_BITACORA: dict(primeras=["Fecha", "Cliente", "Teléfono", "Folio"], congelar=3, borrar=[]),
    TAB_FINALIZADOS: dict(primeras=["NOMBRE", "TELEFONO"], congelar=2, borrar=["MARCA"]),
}
TAB_REVISION = "Revision"


def pendiente(base):
    """Hojas que todavía no tienen el orden nuevo."""
    out = []
    for hoja, cfg in ORDEN.items():
        if hoja not in base.pestanas():
            continue
        enc = [norm(h) for h in base.columnas(hoja)]
        if any(norm(b) in enc for b in cfg["borrar"]):
            out.append(hoja)
            continue
        cols = [base.col(hoja, n) for n in cfg["primeras"]]
        if any(c and c != i for i, c in enumerate(cols, start=1)):
            out.append(hoja)
    return out


def reacomodar(base):
    hecho, mapas = [], {}
    for hoja, cfg in ORDEN.items():
        if hoja not in base.pestanas():
            continue
        if hoja in pendiente(base) and hoja != TAB_BITACORA and hasattr(base, "respaldar"):
            base.respaldar(hoja)                 # la Bitacora se rehace sola; las de captura llevan respaldo
        n = len(base.columnas(hoja))
        pos = list(range(1, n + 1))              # pos[i]: columna original que hoy está en la posición i + 1
        borradas = []
        for nombre in cfg["borrar"]:
            c = base.col(hoja, nombre)
            if c:
                base.borrar_columna(hoja, c)
                borradas.append(pos.pop(c - 1))
                hecho.append(f"Columna {nombre} eliminada de «{hoja}»")
        movidas = []
        for i, nombre in enumerate(cfg["primeras"], start=1):
            c = base.col(hoja, nombre)
            if c and c != i:
                base.mover_columna(hoja, c, i)
                pos.insert(i - 1, pos.pop(c - 1))
                movidas.append(nombre)
        if movidas:
            hecho.append(f"«{hoja}»: {', '.join(n for n in cfg['primeras'] if base.col(hoja, n))} al inicio")
        base.congelar(hoja, 1, cfg["congelar"])
        mapa = {orig: nueva for nueva, orig in enumerate(pos, start=1)}
        mapa.update({b: None for b in borradas})
        if borradas or any(k != v for k, v in mapa.items() if v):
            mapas[hoja] = mapa
    n = _ajustar_revision(base, mapas)
    if n:
        hecho.append(f"{n} celda(s) de Revision apuntan a las columnas nuevas")
    return hecho


def _ajustar_revision(base, mapas):
    """Las celdas de los pendientes (p. ej. «D12» o «L2:L33») se escribieron con las letras viejas."""
    if not mapas or TAB_REVISION not in base.pestanas():
        return 0
    rev = base.leer(TAB_REVISION)
    c_hoja, c_celda = base.col(TAB_REVISION, "Hoja"), base.col(TAB_REVISION, "Celda")
    if rev.empty or not c_hoja or not c_celda:
        return 0
    h_hoja = next(c for c in rev.columns if norm(c) == "HOJA")
    h_celda = next(c for c in rev.columns if norm(c) == "CELDA")
    hojas = {norm(h): h for h in mapas}
    celdas = []
    for _, r in rev.iterrows():
        hoja = hojas.get(norm(r[h_hoja]))
        valor = str(r[h_celda] or "").strip()
        if hoja and valor:
            nuevo = cambiar_referencias(valor, mapas[hoja])
            if nuevo != valor:
                celdas.append((int(r["_fila"]), c_celda, nuevo))
    if celdas:
        base.escribir(TAB_REVISION, celdas)
    return len(celdas)
