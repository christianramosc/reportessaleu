# -*- coding: utf-8 -*-
"""
Base de seguimiento: lectura y escritura del Google Sheet "Seguimiento Rechazos MG Colima".

Dos implementaciones con la misma interfaz:
  · BaseGoogle  → el Google Sheet real (cuenta de servicio, gspread).
  · BaseExcel   → un .xlsx en memoria (modo sin conexión y pruebas). Al final se descarga.

Reglas de escritura:
  · Solo se escribe en las celdas que se indican (nunca se reescribe una fila completa), así
    las columnas de seguimiento que captura el equipo no se tocan.
  · Antes de aplicar cambios se guarda una copia oculta de la pestaña (respaldo).
  · Cada cambio queda en la pestaña "Historial".
"""
import io
import re
import unicodedata
from datetime import datetime, date, timedelta

import pandas as pd

TAB_RECHAZOS = "MG Colima Rechazos"
TAB_FINALIZADOS = "MG Colima Créditos Finalizados"
TAB_BITACORA = "Bitacora"
TAB_CAPTURAS = "Capturas Cetelem"
TAB_FI = "Bitacora F&I"          # copia de los registros de la Bitácora de F&I (para no perderlos entre cargas)
TAB_HISTORIAL = "Historial"
HISTORIAL_COLS = ["Fecha", "Origen", "Hoja", "Celda", "Cliente", "Campo", "Antes", "Después", "Motivo"]
RESPALDOS_A_GUARDAR = 3

_EPOCA = datetime(1899, 12, 30)        # día 0 de Google Sheets / Excel

# Estilo uniforme de las pestañas que escribe la app (el mismo que ya usaba la Bitacora)
ESTILO = dict(fuente="Calibri", tamano=9, texto="#000000", fondo="#FFFFFF", enc_texto="#FFFFFF", enc_fondo="#666666")
FORMATO_NUMERO = {"fecha": ("DATE", "dd/mm/yyyy", "DD/MM/YYYY"), "entero": ("NUMBER", "0", "0"),
                  "moneda": ("CURRENCY", "$#,##0.00", '"$"#,##0.00')}   # (tipo Google, patrón Google, patrón Excel)


def _rgb(color):
    return {k: int(color.lstrip("#")[i:i + 2], 16) / 255 for k, i in (("red", 0), ("green", 2), ("blue", 4))}


def norm(t):
    """Texto comparable: sin acentos, mayúsculas, espacios simples."""
    t = unicodedata.normalize("NFKD", str(t if t is not None else ""))
    t = "".join(ch for ch in t if not unicodedata.combining(ch))
    return re.sub(r"\s+", " ", t).strip().upper()


def a_serial(v):
    """Fecha → número de serie de Sheets (así la columna conserva su formato dd/mm/aaaa).
    También deja los valores listos para la API (sin tipos de numpy ni NaN)."""
    if hasattr(v, "item") and not isinstance(v, (str, bytes)):
        v = v.item()                                   # numpy.int64 / float64 → int / float
    if isinstance(v, float) and v != v:
        return ""
    if isinstance(v, datetime):
        return (v - _EPOCA).total_seconds() / 86400
    if isinstance(v, date):
        return (datetime(v.year, v.month, v.day) - _EPOCA).days
    return v


def de_serial(v):
    if isinstance(v, (int, float)) and 20000 < v < 80000:
        return _EPOCA + timedelta(days=float(v))
    return v


def letra(col):
    s = ""
    while col:
        col, r = divmod(col - 1, 26)
        s = chr(65 + r) + s
    return s


def _es_fecha(encabezado):
    h = norm(encabezado)
    return h.startswith("FECHA") or h in ("FECHA INGRESO",)


class _BaseComun:
    """Lógica compartida; las subclases implementan _valores, _escribir, _crear_pestana, etc."""

    origen = "App"

    # ---------------------------------------------------------------- lectura
    def leer(self, pestana):
        """DataFrame con los encabezados de la fila 1 y la columna '_fila' (número de fila en la hoja)."""
        valores = self._valores(pestana)
        if not valores:
            return pd.DataFrame(columns=["_fila"])
        enc = [str(h).strip() if h not in (None, "") else f"_col{i + 1}" for i, h in enumerate(valores[0])]
        vistos = {}
        for i, h in enumerate(enc):                       # encabezados repetidos → "X", "X (2)"
            if h in vistos:
                vistos[h] += 1
                enc[i] = f"{h} ({vistos[h]})"
            else:
                vistos[h] = 1
        filas = []
        for n, fila in enumerate(valores[1:], start=2):
            fila = list(fila) + [None] * (len(enc) - len(fila))
            if not any(v not in (None, "") for v in fila[:min(len(fila), 20)]):
                continue
            filas.append(fila[:len(enc)] + [n])
        df = pd.DataFrame(filas, columns=enc + ["_fila"])
        for c in enc:
            if _es_fecha(c):
                df[c] = df[c].map(de_serial)
        df = df.astype(object)
        return df.where(df.notna(), None)

    def columnas(self, pestana):
        valores = self._valores(pestana, solo_encabezado=True)
        return [str(h).strip() for h in (valores[0] if valores else [])]

    def col(self, pestana, nombre):
        """Número de columna (1-based) por nombre de encabezado, sin importar acentos/mayúsculas."""
        objetivo = norm(nombre)
        for i, h in enumerate(self.columnas(pestana), start=1):
            if norm(h) == objetivo:
                return i
        return None

    def ultima_fila_con_datos(self, pestana, columnas_clave=(1, 3)):
        valores = self._valores(pestana)
        ultima = 1
        for n, fila in enumerate(valores[1:], start=2):
            if any(len(fila) >= c and fila[c - 1] not in (None, "") for c in columnas_clave):
                ultima = n
        return ultima

    # ---------------------------------------------------------------- escritura
    def escribir(self, pestana, celdas):
        """celdas: lista de (fila, columna, valor). Las fórmulas van como texto que empieza con '='."""
        if celdas:
            self._escribir(pestana, [(f, c, a_serial(v)) for f, c, v in celdas])

    def reemplazar(self, pestana, encabezados, filas, oculta=False):
        """Borra la pestaña y escribe la tabla completa (se usa para Bitacora y Capturas Cetelem)."""
        self._reemplazar(pestana, [list(encabezados)] + [[a_serial(v) for v in f] for f in filas], oculta)

    def marcar(self, pestana, celdas, color):
        """Pinta el fondo de las celdas [(fila, columna)] con color '#RRGGBB'."""
        if celdas:
            self._marcar(pestana, celdas, color)

    def quitar_marca(self, pestana, celdas):
        """Deja las celdas [(fila, columna)] sin relleno y sin nota."""
        if celdas:
            self._marcar(pestana, celdas, "#FFFFFF")
            self._quitar_notas(pestana, celdas)

    def reemplazar_textos(self, pestana, cambios):
        """Cambia texto dentro de las celdas (también dentro de fórmulas). cambios: [(viejo, nuevo)].
        Devuelve cuántas celdas cambiaron."""
        if pestana not in self.pestanas():
            return 0
        celdas = []
        for n, fila in enumerate(self._formulas(pestana), start=1):
            for m, v in enumerate(fila, start=1):
                if isinstance(v, str):
                    nv = v
                    for a, b in cambios:
                        nv = nv.replace(a, b)
                    if nv != v:
                        celdas.append((n, m, nv))
        if celdas:
            self._escribir(pestana, celdas)
        return len(celdas)

    def reemplazar_en_formatos(self, pestana, cambios, filas_min=0):
        """Cambia texto dentro de las reglas de formato condicional de la pestaña (p. ej. ="Contrato").
        cambios: [(viejo, nuevo)]; una regla cuyo valor es exactamente `viejo` sin comillas también cambia.
        filas_min: alarga hasta esa fila el rango de las reglas de la pestaña (en Google).
        Devuelve cuántas reglas cambiaron."""
        if pestana not in self.pestanas():
            return 0
        return self._reemplazar_en_formatos(pestana, cambios, filas_min)

    def dar_formato(self, pestana, numeros=None):
        """Un solo estilo para toda la pestaña (ESTILO): misma letra, tamaño y color, fondo blanco y encabezado
        gris con letra blanca. numeros: {encabezado: "fecha" | "entero" | "moneda"} para el formato de esas columnas.
        No toca el formato condicional (los colores por estatus)."""
        enc = self.columnas(pestana)
        cols = {i: tipo for i, h in enumerate(enc, start=1) for nombre, tipo in (numeros or {}).items()
                if norm(h) == norm(nombre)}
        self._dar_formato(pestana, len(enc), cols)

    def anotar_historial(self, registros):
        """registros: dicts con las llaves de HISTORIAL_COLS (Fecha se llena sola)."""
        if not registros:
            return
        if TAB_HISTORIAL not in self.pestanas():
            self._reemplazar(TAB_HISTORIAL, [HISTORIAL_COLS], False)
        inicio = self.ultima_fila_con_datos(TAB_HISTORIAL, (1, 2)) + 1
        ahora = datetime.now()
        celdas = []
        for i, r in enumerate(registros):
            for j, k in enumerate(HISTORIAL_COLS, start=1):
                v = ahora if k == "Fecha" else r.get(k, "")
                if k == "Origen" and not v:
                    v = self.origen
                celdas.append((inicio + i, j, _texto_historial(v)))
        self.escribir(TAB_HISTORIAL, celdas)


def _cambiar_texto(obj, cambios):
    """Aplica los cambios a todos los textos de una estructura (dict/list/str)."""
    if isinstance(obj, str):
        for a, b in cambios:
            obj = b.strip('"') if obj == a.strip('"') else obj.replace(a, b)
        return obj
    if isinstance(obj, list):
        return [_cambiar_texto(x, cambios) for x in obj]
    if isinstance(obj, dict):
        return {k: _cambiar_texto(v, cambios) for k, v in obj.items()}
    return obj


def _texto_historial(v):
    if isinstance(v, (datetime, date)) or v is None or isinstance(v, (int, float)):
        return "" if v is None else v
    return str(v)


# =============================================================================
# Google Sheets
# =============================================================================
_MES_ANIO = re.compile(r"(?i)^\s*(enero|febrero|marzo|abril|mayo|junio|julio|agosto|septiembre|setiembre|"
                       r"octubre|noviembre|diciembre)\s+\d{4}\s*$")


def _como_texto(v):
    """Google interpreta "JUNIO 2026" como la fecha 01/06/2026; con apóstrofo se queda como texto."""
    return "'" + v if isinstance(v, str) and _MES_ANIO.match(v) else v


class BaseGoogle(_BaseComun):
    def __init__(self, info_cuenta_servicio, url):
        import gspread
        from gspread.http_client import BackOffHTTPClient
        self._gs = gspread
        # BackOffHTTPClient reintenta solo (con espera creciente) si Google responde 429 por cuota
        self.cliente = gspread.service_account_from_dict(dict(info_cuenta_servicio), http_client=BackOffHTTPClient)
        self.libro = self.cliente.open_by_url(url)
        self.url = url
        self.titulo = self.libro.title
        self._cache = {}
        self._hojas = None

    def _mapa(self):
        """{título: Worksheet}. Se pide a Google una sola vez por ejecución (cada consulta cuenta en la cuota)."""
        if self._hojas is None:
            self._hojas = {w.title: w for w in self.libro.worksheets()}
        return self._hojas

    def pestanas(self):
        return list(self._mapa())

    def _hoja(self, pestana):
        ws = self._mapa().get(pestana)
        if ws is None:                      # quizá se creó después de leer la lista
            self._hojas = None
            ws = self._mapa().get(pestana)
        if ws is None:
            raise ValueError(f"No existe la pestaña '{pestana}' en el Sheet.")
        return ws

    def _valores(self, pestana, solo_encabezado=False):
        clave = (pestana, solo_encabezado)
        if clave not in self._cache:
            from gspread.utils import ValueRenderOption, DateTimeOption
            ws = self._hoja(pestana)
            rango = "1:1" if solo_encabezado else None
            self._cache[clave] = ws.get(rango, value_render_option=ValueRenderOption.unformatted,
                                        date_time_render_option=DateTimeOption.serial_number,
                                        pad_values=True)
        return self._cache[clave]

    def _limpiar_cache(self, pestana):
        self._cache = {k: v for k, v in self._cache.items() if k[0] != pestana}

    def _escribir(self, pestana, celdas):
        from gspread.utils import ValueInputOption
        ws = self._hoja(pestana)
        max_fila = max(f for f, _, _ in celdas)
        max_col = max(c for _, c, _ in celdas)
        if max_fila > ws.row_count or max_col > ws.col_count:
            ws.resize(rows=max(ws.row_count, max_fila + 50), cols=max(ws.col_count, max_col))
        datos = [{"range": f"{letra(c)}{f}", "values": [["" if v is None else _como_texto(v)]]} for f, c, v in celdas]
        for i in range(0, len(datos), 500):            # lotes para no rebasar la cuota de la API
            ws.batch_update(datos[i:i + 500], value_input_option=ValueInputOption.user_entered)
        self._limpiar_cache(pestana)

    def _reemplazar(self, pestana, tabla, oculta):
        from gspread.utils import ValueInputOption
        ncols = max(len(f) for f in tabla)
        if pestana in self.pestanas():
            ws = self._hoja(pestana)
            ws.clear()
            if ws.row_count < len(tabla) or ws.col_count < ncols:
                ws.resize(rows=max(ws.row_count, len(tabla) + 50), cols=max(ws.col_count, ncols))
        else:
            ws = self.libro.add_worksheet(pestana, rows=len(tabla) + 50, cols=ncols)
            self._hojas = None
        tabla = [["" if v is None else _como_texto(v) for v in f] + [""] * (ncols - len(f)) for f in tabla]
        ws.update(tabla, "A1", value_input_option=ValueInputOption.user_entered)
        if oculta:
            ws.hide()
        self._limpiar_cache(pestana)

    def _formulas(self, pestana):
        from gspread.utils import ValueRenderOption
        return self._hoja(pestana).get(value_render_option=ValueRenderOption.formula, pad_values=True)

    def _reemplazar_en_formatos(self, pestana, cambios, filas_min):
        ws = self._hoja(pestana)
        meta = self.libro.fetch_sheet_metadata(
            params={"fields": "sheets(properties(sheetId),conditionalFormats)"})
        hoja = next((h for h in meta.get("sheets", []) if h.get("properties", {}).get("sheetId") == ws.id), {})
        pedidos = []
        for i, regla in enumerate(hoja.get("conditionalFormats", [])):
            nueva = _cambiar_texto(regla, cambios)
            for r in nueva.get("ranges", []):          # que el color alcance a las filas nuevas
                if filas_min and r.get("endRowIndex") is not None and r["endRowIndex"] < filas_min:
                    r["endRowIndex"] = filas_min
            if nueva == regla:
                continue
            pedidos.append({"updateConditionalFormatRule": {"sheetId": ws.id, "index": i, "rule": nueva}})
        if pedidos:
            self.libro.batch_update({"requests": pedidos})
        return len(pedidos)

    def _dar_formato(self, pestana, ncols, numeros):
        ws = self._hoja(pestana)
        filas, ncols = ws.row_count, max(ncols, 1)
        rango = lambda f0, f1, c0, c1: {"sheetId": ws.id, "startRowIndex": f0, "endRowIndex": f1,
                                        "startColumnIndex": c0, "endColumnIndex": c1}
        texto = lambda color, negrita: {"fontFamily": ESTILO["fuente"], "fontSize": ESTILO["tamano"], "bold": negrita,
                                        "italic": False, "strikethrough": False, "underline": False,
                                        "foregroundColor": _rgb(color)}
        campos = "userEnteredFormat(textFormat,backgroundColor,horizontalAlignment,verticalAlignment,wrapStrategy,numberFormat)"
        pedidos = [
            {"repeatCell": {"range": rango(1, filas, 0, ws.col_count), "fields": campos, "cell": {"userEnteredFormat": {
                "textFormat": texto(ESTILO["texto"], False), "backgroundColor": _rgb(ESTILO["fondo"]),
                "horizontalAlignment": "LEFT", "verticalAlignment": "MIDDLE", "wrapStrategy": "CLIP"}}}},
            {"repeatCell": {"range": rango(0, 1, 0, ws.col_count), "fields": campos, "cell": {"userEnteredFormat": {
                "textFormat": texto(ESTILO["enc_texto"], True), "backgroundColor": _rgb(ESTILO["enc_fondo"]),
                "horizontalAlignment": "CENTER", "verticalAlignment": "MIDDLE", "wrapStrategy": "WRAP"}}}},
            {"updateSheetProperties": {"properties": {"sheetId": ws.id, "gridProperties": {"frozenRowCount": 1}},
                                       "fields": "gridProperties.frozenRowCount"}},
        ]
        for col, tipo in numeros.items():
            t, patron, _ = FORMATO_NUMERO[tipo]
            pedidos.append({"repeatCell": {
                "range": rango(1, filas, col - 1, col), "fields": "userEnteredFormat(numberFormat,horizontalAlignment)",
                "cell": {"userEnteredFormat": {"numberFormat": {"type": t, "pattern": patron},
                                               "horizontalAlignment": "LEFT" if tipo == "fecha" else "RIGHT"}}}})
        self.libro.batch_update({"requests": pedidos})

    def _marcar(self, pestana, celdas, color):
        r, g, b = (int(color.lstrip("#")[i:i + 2], 16) / 255 for i in (0, 2, 4))
        formatos = [{"range": f"{letra(c)}{f}", "format": {"backgroundColor": {"red": r, "green": g, "blue": b}}}
                    for f, c in celdas]
        self._hoja(pestana).batch_format(formatos)

    def _quitar_notas(self, pestana, celdas):
        self._hoja(pestana).clear_notes([f"{letra(c)}{f}" for f, c in celdas])

    def respaldar(self, pestana):
        """Copia oculta de la pestaña antes de escribir; conserva las últimas RESPALDOS_A_GUARDAR."""
        prefijo = f"_resp {pestana[:20]} "
        nombre = prefijo + datetime.now().strftime("%Y-%m-%d %H%M%S")
        existentes = {w.title for w in self.libro.worksheets()}
        n = 2
        while nombre in existentes:          # dos respaldos en el mismo segundo: Google no admite nombres repetidos
            nombre = prefijo + datetime.now().strftime("%Y-%m-%d %H%M%S") + f" ({n})"
            n += 1
        copia = self._hoja(pestana).duplicate(new_sheet_name=nombre)
        copia.hide()
        viejas = sorted([w for w in self.libro.worksheets() if w.title.startswith(prefijo)], key=lambda w: w.title)
        for w in viejas[:-RESPALDOS_A_GUARDAR]:
            self.libro.del_worksheet(w)
        self._hojas = None
        return nombre

    def refrescar(self):
        self._cache = {}
        self._hojas = None


# =============================================================================
# Excel (modo sin conexión / pruebas)
# =============================================================================
class BaseExcel(_BaseComun):
    def __init__(self, contenido, nombre="Seguimiento.xlsx"):
        import openpyxl
        self._op = openpyxl
        self.libro = openpyxl.load_workbook(io.BytesIO(contenido))
        self.titulo = nombre
        self.url = None
        self.origen = "App (Excel)"

    def pestanas(self):
        return self.libro.sheetnames

    def _valores(self, pestana, solo_encabezado=False):
        ws = self.libro[pestana]
        filas = ws.iter_rows(min_row=1, max_row=1 if solo_encabezado else ws.max_row, values_only=True)
        out = []
        for f in filas:
            out.append([None if isinstance(v, str) and v.startswith("=") else v for v in f])
        return out

    def _escribir(self, pestana, celdas):
        ws = self.libro[pestana]
        for f, c, v in celdas:
            if isinstance(v, (int, float)) and _es_fecha(ws.cell(1, c).value) and 20000 < v < 80000:
                v = de_serial(v)
            ws.cell(f, c).value = v

    def _reemplazar(self, pestana, tabla, oculta):
        formatos = None
        if pestana in self.libro.sheetnames:
            idx = self.libro.sheetnames.index(pestana)
            formatos = self.libro[pestana].conditional_formatting      # se conservan los colores condicionales
            del self.libro[pestana]
            ws = self.libro.create_sheet(pestana, idx)
            ws.conditional_formatting = formatos
        else:
            ws = self.libro.create_sheet(pestana)
        for fila in tabla:
            ws.append([de_serial(v) if isinstance(v, (int, float)) and _es_fecha(tabla[0][i] if i < len(tabla[0]) else "")
                       else v for i, v in enumerate(fila)])
        if oculta:
            ws.sheet_state = "hidden"

    def _marcar(self, pestana, celdas, color):
        from openpyxl.styles import PatternFill
        relleno = (PatternFill(fill_type=None) if color.upper() in ("#FFFFFF", "FFFFFF")
                   else PatternFill("solid", fgColor="FF" + color.lstrip("#").upper()))
        ws = self.libro[pestana]
        for f, c in celdas:
            ws.cell(f, c).fill = relleno

    def _formulas(self, pestana):
        return [[c.value for c in fila] for fila in self.libro[pestana].iter_rows()]

    def _dar_formato(self, pestana, ncols, numeros):
        from openpyxl.styles import Alignment, Font, PatternFill
        ws = self.libro[pestana]
        sin = lambda c: c.lstrip("#").upper()
        letra_normal = Font(name=ESTILO["fuente"], size=ESTILO["tamano"], color="FF" + sin(ESTILO["texto"]))
        letra_enc = Font(name=ESTILO["fuente"], size=ESTILO["tamano"], bold=True, color="FF" + sin(ESTILO["enc_texto"]))
        fondo_enc = PatternFill("solid", fgColor="FF" + sin(ESTILO["enc_fondo"]))
        for fila in ws.iter_rows(min_row=1, max_row=ws.max_row, max_col=max(ws.max_column, ncols)):
            for c in fila:
                if c.row == 1:
                    c.font, c.fill = letra_enc, fondo_enc
                    c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                else:
                    c.font, c.fill = letra_normal, PatternFill(fill_type=None)
                    tipo = numeros.get(c.column)
                    c.number_format = FORMATO_NUMERO[tipo][2] if tipo else "General"
                    c.alignment = Alignment(horizontal="right" if tipo in ("entero", "moneda") else "left",
                                            vertical="center")
        ws.freeze_panes = ws.freeze_panes or "A2"

    def _reemplazar_en_formatos(self, pestana, cambios, filas_min):
        n = 0
        for rango in self.libro[pestana].conditional_formatting:
            for regla in rango.rules:
                nuevas = [_cambiar_texto(f, cambios) for f in (regla.formula or [])]
                if nuevas != list(regla.formula or []):
                    regla.formula = nuevas
                    n += 1
        return n

    def _quitar_notas(self, pestana, celdas):
        ws = self.libro[pestana]
        for f, c in celdas:
            ws.cell(f, c).comment = None

    def respaldar(self, pestana):
        nombre = f"_resp {pestana[:12]} " + datetime.now().strftime("%m%d %H%M%S")
        copia = self.libro.copy_worksheet(self.libro[pestana])
        copia.title = nombre[:31]
        copia.sheet_state = "hidden"
        return copia.title

    def refrescar(self):
        pass

    def guardar(self):
        buf = io.BytesIO()
        self.libro.save(buf)
        return buf.getvalue()


def a_excel(base, pestanas):
    """Arma un .xlsx con las pestañas indicadas (lo usan los módulos de cuadre que leen archivos)."""
    from openpyxl import Workbook
    wb = Workbook()
    wb.remove(wb.active)
    for p in pestanas:
        ws = wb.create_sheet(p[:31])
        valores = base._valores(p)
        enc = valores[0] if valores else []
        fechas = {i for i, h in enumerate(enc) if _es_fecha(h)}
        for fila in valores:
            ws.append([de_serial(v) if i in fechas and isinstance(v, (int, float)) else v
                       for i, v in enumerate(fila)])
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue()
