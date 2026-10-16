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
TAB_HISTORIAL = "Historial"
HISTORIAL_COLS = ["Fecha", "Origen", "Hoja", "Celda", "Cliente", "Campo", "Antes", "Después", "Motivo"]
RESPALDOS_A_GUARDAR = 3

_EPOCA = datetime(1899, 12, 30)        # día 0 de Google Sheets / Excel


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


def _texto_historial(v):
    if isinstance(v, (datetime, date)) or v is None or isinstance(v, (int, float)):
        return "" if v is None else v
    return str(v)


# =============================================================================
# Google Sheets
# =============================================================================
class BaseGoogle(_BaseComun):
    def __init__(self, info_cuenta_servicio, url):
        import gspread
        self._gs = gspread
        self.cliente = gspread.service_account_from_dict(dict(info_cuenta_servicio))
        self.libro = self.cliente.open_by_url(url)
        self.url = url
        self.titulo = self.libro.title
        self._cache = {}

    def pestanas(self):
        return [w.title for w in self.libro.worksheets()]

    def _hoja(self, pestana):
        return self.libro.worksheet(pestana)

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
        datos = [{"range": f"{letra(c)}{f}", "values": [["" if v is None else v]]} for f, c, v in celdas]
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
        tabla = [["" if v is None else v for v in f] + [""] * (ncols - len(f)) for f in tabla]
        ws.update(tabla, "A1", value_input_option=ValueInputOption.user_entered)
        if oculta:
            ws.hide()
        self._limpiar_cache(pestana)

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
        nombre = prefijo + datetime.now().strftime("%Y-%m-%d %H%M")
        copia = self._hoja(pestana).duplicate(new_sheet_name=nombre)
        copia.hide()
        viejas = sorted([w for w in self.libro.worksheets() if w.title.startswith(prefijo)], key=lambda w: w.title)
        for w in viejas[:-RESPALDOS_A_GUARDAR]:
            self.libro.del_worksheet(w)
        return nombre

    def refrescar(self):
        self._cache = {}


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
        if pestana in self.libro.sheetnames:
            idx = self.libro.sheetnames.index(pestana)
            del self.libro[pestana]
            ws = self.libro.create_sheet(pestana, idx)
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
