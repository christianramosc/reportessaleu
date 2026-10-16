# Herramientas MG Colima

App de Streamlit con una página por herramienta:

1. **KPIs de capacitación**: avance de capacitación de Ventas y Posventa.
2. **Base de seguimiento**: mantiene al día el Google Sheet *Seguimiento Rechazos MG Colima*. Subes las
   bitácoras mensuales de Cetelem, la app lee la *Bitácora de F&I*, te muestra los cambios propuestos
   (rechazos nuevos, mensualidades, folios, GAP, plazo, teléfono, Sale-U, "ya compró") y aplica los que
   marques. Antes de escribir guarda un respaldo oculto y todo queda en la pestaña **Historial**.
3. **Cuadres para reportes**: cuadra la **bitácora** (pestaña *Bitacora* del Sheet: capturas de Cetelem +
   Bitácora de F&I, **todas las financieras**) y el inventario. Con el Google Sheet conectado, el seguimiento
   y la bitácora se leen de la base; para Sale-U solo subes el reporte y la app escribe
   *Cliente con Registro en SalesU* en el Sheet:
   - **Seguimiento de rechazos vs bitácora** (solo pestañas de COLIMA). Cuadra solo las solicitudes
     **rechazadas** de cualquier financiera; cruza por folio, teléfono y nombre.
   - **Sale-U vs bitácora**: reporte de Sale-U (`reporteLeadsconCredito_*.csv`, tal cual se descarga). Toma
     las solicitudes de la bitácora dentro del periodo del reporte (sin ventas de contado), cruza por teléfono
     y nombre y compara estatus y financiera.
   - **BlueService vs Inventario** (`PDIs_MG_Colima.xlsx` vs Excel de inventario, hoja ALMACEN o la que
     elijas). Cruza por VIN, detecta VIN mal capturados con el dígito verificador y dice en qué otra hoja
     aparece cada unidad.

   Cada cuadre genera un **Excel** y un **PDF** descargables.

## Estructura

```
streamlit_app.py               # entrada: contraseña opcional y navegación
ui.py                          # estilos y componentes compartidos
kpis_core.py                   # lógica de KPIs de capacitación
paginas/kpis_capacitacion.py   # página de KPIs
paginas/cuadres_reportes.py    # página de cuadres
paginas/base_seguimiento.py    # página para actualizar la base con las bitácoras
conexion.py                    # conexión con el Google Sheet (o modo Excel)
cuadres/base.py                # leer/escribir el Sheet (Google o .xlsx), respaldo e Historial
cuadres/fuentes.py             # bitácoras de Cetelem y Bitácora de F&I → solicitudes
cuadres/actualizar.py          # cambios propuestos y cómo se aplican
cuadres/saleu_sheet.py         # registro en Sale-U → columna del Sheet
cuadres/pendientes.py          # limpiar pendientes resueltos de la pestaña Revision
cuadres/seguimiento.py         # cuadre vs Excel de seguimiento
cuadres/saleu.py               # cuadre vs Sale-U
cuadres/blueservice.py         # cuadre BlueService vs Inventario
.streamlit/config.toml         # tema
```

## Correr en local

```bash
pip install -r requirements.txt
streamlit run streamlit_app.py
```

## Contraseña (recomendado)

Los archivos traen nombres y teléfonos de clientes. Para pedir contraseña al entrar,
define `APP_PASSWORD` en `.streamlit/secrets.toml` (local) o en **Settings → Secrets**
de Streamlit Community Cloud. Ver `.streamlit/secrets.toml.example`.

## Conectar el Google Sheet (una sola vez, ~15 min)

Sin esto la app funciona en **modo Excel**: subes el .xlsx de seguimiento y descargas la versión actualizada.

1. Entra a <https://console.cloud.google.com/>, crea un proyecto (ej. `cuadres-mg`).
2. **APIs y servicios → Biblioteca**: habilita **Google Sheets API** y **Google Drive API**.
3. **IAM y administración → Cuentas de servicio → Crear cuenta de servicio** (ej. `cuadres-mg`). No necesita roles.
4. En la cuenta creada: **Claves → Agregar clave → JSON**. Se descarga un archivo `.json` (guárdalo bien,
   es la llave; nunca lo subas al repo).
5. En el Google Sheet *Seguimiento Rechazos MG Colima*: **Compartir** con el `client_email` del JSON como
   **Editor**. En la *Bitácora de F&I* (la original): compartir como **Lector**.
6. En Streamlit Cloud: **Settings → Secrets**, pega el contenido de `.streamlit/secrets.toml.example`
   con tus enlaces y los datos del JSON (cada campo del JSON va en la sección `[gcp_service_account]`).
7. Abre la página **Base de seguimiento**: debe decir *Conectado a …*.

### Cómo se usa cada mes

1. **Base de seguimiento** → sube las bitácoras de Cetelem del mes (el .xlsx tal cual) y, si no está conectada
   por Google, la Bitácora de F&I → *Revisar cambios* → desmarca lo que no quieras → *Aplicar*. Esto también
   actualiza la pestaña *Bitacora*, que es la que usan los cuadres. Lo desmarcado queda como "Ignorado" en el Historial y no se vuelve a
   proponer.
2. **Cuadres para reportes → Sale-U** → sube el reporte de Sale-U → *Hacer cuadre*. Además del Excel/PDF,
   marca *Registrado / No Registrado* en cada rechazo cuya fecha de ingreso cae en el periodo del reporte.
   Lo que ya dice *Registrado* o *Sí* (p. ej. marcado a mano) nunca se baja a *No Registrado*: si no aparece en
   el reporte se deja igual y se lista para que lo confirmes.
3. **Cuadres para reportes → Seguimiento** → *Hacer cuadre* (ya no pide archivos).
4. **Base de seguimiento → Pendientes** → *Revisar pendientes* → *Limpiar*. Los pendientes de la pestaña
   *Revision* que ya tienen el dato capturado pasan a *Corregido* y se les quita el morado y la nota en la
   hoja. Si un pendiente no aplica o lo resolviste de otra forma, cámbialo a mano a *Corregido* o *No aplica*
   en *Revision* y el botón le quita la marca.

Reglas que sigue la app al escribir en el Sheet:
- **Formalizado** = el crédito terminó en contrato con cualquier financiera (FINANCIADO en Cetelem,
  CONTRATO en la Bitácora de F&I). Así aparece en la *Bitacora* y en el *Resumen*.
- Lo de Cetelem manda en folio, estatus, mensualidad (máxima), plazo, teléfono, GAP y accesorios.
- Solo escribe en las columnas de datos; nunca en las de seguimiento (fechas realizadas, canal, contesta,
  comentarios) salvo agregar "Ya compró…" al final de COMENTARIOS.
- Un rechazo que solo aparece en F&I pero que en Cetelem está aprobado o en contrapropuesta se propone
  **desmarcado** para que lo confirmes.
- Antes de escribir guarda una copia oculta de la pestaña (`_resp …`, se conservan las 3 últimas).

## Notas

- Subir el CSV de Sale-U **sin abrirlo y guardarlo en Excel**: Excel convierte algunos
  teléfonos a notación científica y se pierden dígitos.
- Los módulos de `cuadres/` también funcionan solos en Google Colab (`python cuadres/saleu.py`).
- La sincronización con Apps Script (`SincronizarBitacora.gs`) quedó reemplazada por esta app. Si la
  instalaste: Extensiones → Apps Script → borra el código y, en *Activadores*, elimina el disparador; puedes
  borrar las pestañas *Configuracion* y *_sync_bitacora* del Sheet.
