# -*- coding: utf-8 -*-
"""
Ajustes de estructura y formato del Google Sheet de seguimiento (se corren una vez con un botón; repetirlos no
duplica nada).

  0. Fecha, nombre y teléfono como primeras columnas de Rechazos y Bitacora (Finalizados: nombre y teléfono, sin
     MARCA), congeladas (cuadres/columnas.py).
  1. Encabezado del registro en Sale-U con el mismo nombre en las dos pestañas («Registro en Sale-U»).
  2. MG Colima Rechazos: columnas Reingreso e ID al final (ver cuadres/ids.py).
  3. CONTESTA 1/2/3: opción «No responde» y color cuando hay fecha de contacto pero no se capturó si contestó.
  4. Columnas que siempre traen lo mismo (Tipo de Subvención, Accesorios, Status 1) agrupadas y contraídas.
  5. MG Colima Créditos Finalizados: columna FECHA DE ENTREGA (con ella entran al cuadre de Sale-U por periodo).
  6. Historial y Bitacora F&I: las fechas con formato de fecha (antes se veían como números); Capturas Cetelem y
     Bitacora F&I ocultas; en la Bitacora, «Registro en Sale-U» con los mismos colores que en Rechazos.
  7. Revision: ID de cada pendiente y los resueltos pasan a «Revision resueltos».
  8. Se rehace el Resumen.
  9. Mantenimiento (cuadres/optimizar.py): un solo respaldo por pestaña (se borran los viejos y las copias «Hoja N»),
     columnas fijas llenas, ID del rechazo previo en Finalizados, columnas calculadas y Bitacora protegidas (solo
     advertencia), encabezado congelado, vistas filtradas por asesor, «Revision resueltos» oculta y pestañas en
     orden: Resumen, Rechazos, Finalizados, Bitacora, Revision.
Antes de tocar una pestaña se guarda su respaldo oculto.
"""
from cuadres.base import TAB_RECHAZOS, TAB_FINALIZADOS, TAB_HISTORIAL, TAB_FI, letra
from cuadres import ids, pendientes
from cuadres.saleu_sheet import COL_FECHA_ENTREGA

ENCABEZADOS_SALEU = {TAB_RECHAZOS: ("Cliente con Registro en SalesU", "Registro en Sale-U"),
                     TAB_FINALIZADOS: ("CLIENTE CON REGISTRO EN SALE-U", "REGISTRO EN SALE-U")}
CONTESTA = ["Si Responde", "No responde", "No Responde. Enviar WhatsApp"]
COLOR_FALTA = "#FCE5CD"                           # mismo naranja claro de «falta correo o Sale-U»
COLUMNAS_FIJAS = ["Tipo de Subvención", "Accesorios", "Status 1"]
FILAS = ids.FILAS_FORMULA


def aplicar(base):
    """Aplica los ajustes que falten. Devuelve una lista de frases con lo que hizo."""
    hecho = []
    # 0. fecha, nombre y teléfono al inicio (congelados) y sin la columna MARCA en Finalizados
    from cuadres import columnas
    hecho += columnas.reacomodar(base)
    pestanas = set(base.pestanas())
    respaldadas = set()

    def respaldar(p):
        if p not in respaldadas:
            base.respaldar(p)
            respaldadas.add(p)

    # 1. encabezado de Sale-U
    for hoja, (viejo, nuevo) in ENCABEZADOS_SALEU.items():
        if hoja in pestanas and base.renombrar_encabezado(hoja, viejo, nuevo):
            hecho.append(f"«{viejo}» ahora se llama «{nuevo}» en {hoja}")

    if TAB_RECHAZOS in pestanas:
        # 2. Reingreso e ID
        tenia_id = bool(base.col(TAB_RECHAZOS, ids.COL_ID))
        tenia_rein = bool(base.col(TAB_RECHAZOS, ids.COL_REINGRESO))
        if not (tenia_id and tenia_rein):
            respaldar(TAB_RECHAZOS)
        if not tenia_rein:
            ids.asegurar_reingreso(base)
            hecho.append("Columna Reingreso en MG Colima Rechazos")
        antes = ids.mapa(base)
        todos = ids.asegurar(base)
        nuevos = len(todos) - len(antes)
        if nuevos:
            hecho.append(f"ID a {nuevos} rechazo(s) en MG Colima Rechazos")

        # 3. CONTESTA
        cols_contesta, nuevas = [], 0
        for i in (1, 2, 3):
            cc, cf = base.col(TAB_RECHAZOS, f"CONTESTA {i}"), base.col(TAB_RECHAZOS, f"FECHA REALIZADO {i}")
            if cc and cf:
                cols_contesta.append(cc)
                if base.regla_color(TAB_RECHAZOS, [cc], f'=AND(${letra(cf)}2<>"",${letra(cc)}2="")', COLOR_FALTA,
                                    FILAS):
                    nuevas += 1
        if nuevas:
            hecho.append("Color en CONTESTA 1, 2 y 3 cuando hay fecha de contacto pero no se capturó si contestó")
        if cols_contesta and nuevas:
            base.lista_valores(TAB_RECHAZOS, cols_contesta, CONTESTA, FILAS)
            hecho.append("Opción «No responde» en CONTESTA 1, 2 y 3")

        # 4. columnas que siempre traen lo mismo
        fijas = [c for c in (base.col(TAB_RECHAZOS, n) for n in COLUMNAS_FIJAS) if c]
        n = base.agrupar_columnas(TAB_RECHAZOS, fijas) if fijas else 0
        if n:
            hecho.append(f"{n} columna(s) agrupadas en MG Colima Rechazos ({', '.join(COLUMNAS_FIJAS)})")

    # 5. fecha de entrega en Finalizados
    if TAB_FINALIZADOS in pestanas and not base.col(TAB_FINALIZADOS, COL_FECHA_ENTREGA):
        respaldar(TAB_FINALIZADOS)
        c = base.agregar_columna(TAB_FINALIZADOS, COL_FECHA_ENTREGA)
        base.formato_columnas(TAB_FINALIZADOS, {c: "fecha"}, FILAS)
        base.validar_fecha(TAB_FINALIZADOS, [c], FILAS)
        hecho.append("Columna FECHA DE ENTREGA en MG Colima Créditos Finalizados")

    # 6. fechas que se veían como número
    if TAB_HISTORIAL in pestanas:
        base.formato_columnas(TAB_HISTORIAL, {1: "fechahora"})
        hecho.append("Fechas del Historial con formato de fecha")
    if TAB_FI in pestanas:
        from cuadres.actualizar import FORMATO_FI
        base.dar_formato(TAB_FI, FORMATO_FI)
        hecho.append("Fechas de Bitacora F&I con formato de fecha")

    # 6b. pestañas de apoyo ocultas y estilo de Sale-U en la Bitacora
    from cuadres.base import TAB_CAPTURAS, TAB_BITACORA
    from cuadres.saleu_fuente import TAB_FUENTE
    for p in (TAB_CAPTURAS, TAB_FI, TAB_FUENTE):
        if p in pestanas:
            base.ocultar(p)
    if any(p in pestanas for p in (TAB_CAPTURAS, TAB_FI)):
        hecho.append("Pestañas Capturas Cetelem y Bitacora F&I ocultas")
    if TAB_BITACORA in pestanas:
        from cuadres.actualizar import estilo_bitacora
        estilo_bitacora(base)
    from cuadres.saleu_sheet import estilo_registro
    estilo_registro(base)                      # Registrado / Falta crédito / No Registrado en Rechazos y Bitacora

    # 6c. zona horaria del Sheet (HOY() y las fechas de las fórmulas en hora de México)
    if hasattr(base, "libro") and hasattr(base.libro, "batch_update") and hasattr(base, "url"):
        try:
            base.libro.batch_update({"requests": [{"updateSpreadsheetProperties": {
                "properties": {"timeZone": "America/Mexico_City"}, "fields": "timeZone"}}]})
            hecho.append("Zona horaria del Sheet: Ciudad de México")
        except Exception:
            pass

    # 7. Revision
    if pendientes.TAB_REVISION in pestanas:
        n = pendientes.ligar_ids(base)
        if n:
            hecho.append(f"ID en {n} pendiente(s) de Revision")
        n = pendientes.archivar(base)
        if n:
            hecho.append(f"{n} renglón(es) resueltos pasaron a «{pendientes.TAB_ARCHIVO}»")

    # 8. Resumen
    from cuadres import dashboard
    dashboard.construir(base)
    hecho.append("Resumen rehecho")

    # 9. Mantenimiento
    from cuadres import optimizar
    from cuadres.dashboard import TAB_RESUMEN
    borradas = optimizar.limpiar_respaldos(base)
    if borradas:
        hecho.append(f"{len(borradas)} respaldo(s) viejo(s) borrado(s): {', '.join(borradas)}")
    n = optimizar.llenar_fijos(base)
    if n:
        hecho.append(f"{n} celda(s) de Tipo de Plan / Subvención / Accesorios / Status 1 llenadas")
    n = optimizar.ligar_finalizados(base)
    if n:
        hecho.append(f"ID del rechazo previo en {n} crédito(s) de Finalizados")
    hecho += optimizar.proteger(base)
    n = optimizar.vistas_por_asesor(base)
    hecho.append("Encabezado, fecha, nombre y teléfono congelados en Rechazos, Finalizados y Bitacora"
                 + (f"; {n} vista(s) filtrada(s) por asesor" if n else ""))
    if pendientes.TAB_ARCHIVO in base.pestanas():
        base.ocultar(pendientes.TAB_ARCHIVO)
        hecho.append(f"«{pendientes.TAB_ARCHIVO}» oculta")
    base.ordenar([TAB_RESUMEN, TAB_RECHAZOS, TAB_FINALIZADOS, TAB_BITACORA, pendientes.TAB_REVISION])
    hecho.append("Pestañas en orden: Resumen, Rechazos, Finalizados, Bitacora, Revision")
    return hecho
