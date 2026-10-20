# Herramientas MG Colima

App de Streamlit con una página por herramienta:

1. **Inicio**: los 6 indicadores del año (rechazos, seguimiento al día, contacto efectivo, registrados en Sale-U,
   seguimientos vencidos y pendientes en Revision), las **prioridades** de la semana, qué archivos faltan por subir,
   un **tablero** de gráficas interactivas sin nombres (vencidos por días de atraso e ID —clic en un rango para
   filtrar—, agenda de seguimientos por semana, seguimiento al día, embudo, resultado y canal del contacto, registro
   en Sale-U y rechazos por financiera) y el **resumen ejecutivo de una página** en PDF.
2. **Actualizar**: un solo paso para todo lo del día. Subes juntos los archivos que tengas —capturas de Cetelem,
   Bitácora de F&I, `reporteLeadsconCredito` y `ReporteFuentes` de Sale-U, tal cual se descargan— y la app
   **reconoce cada uno** (`cuadres/detectar.py`). *Revisar* muestra:
   1. **Bitácoras**: rechazos nuevos, correcciones, datos faltantes y comentarios para *MG Colima Rechazos* y
      *Finalizados* (desmarcas lo que no quieras; no se vuelve a proponer).
   2. **Sale-U**: el *Registro en Sale-U* (Registrado / Falta crédito / No Registrado) de Rechazos y de la
      *Bitacora*, los datos que se llenan y el reporte del cuadre Sale-U vs bitácora (Excel y PDF).

   *Aplicar todo* escribe primero las bitácoras (y rehace la *Bitacora*) y después el Registro en Sale-U sobre el
   Sheet ya actualizado. Arriba se ve **lo último que se subió** de cada archivo y avisa lo que falta («Falta la
   captura de octubre», «el reporte de Crédito tiene 9 días»); se guarda en la pestaña oculta *Cargas*. Antes de
   escribir guarda un respaldo oculto y todo queda en la pestaña **Historial**.
3. **Seguimiento**:
   - **Resumen**: indicadores del año elegido, gráficas, seguimientos vencidos (con **PDF de vencidos por
     asesor**: uno o todos, una página cada uno), por vencer y pendientes en Revision. Abajo, **Faltan en Sale-U**:
     **Dar de alta** («No Registrado») y **Registrar crédito** («Falta crédito»), con «Cómo está en Sale-U» y la
     *Fuente* del lead; *Generar PDF y Excel* arma el reporte para los asesores.
   - **Pendientes**: revisa los pendientes de *Revision* y limpia los resueltos.
4. **BlueService**: cuadre por VIN de `PDIs_MG_Colima.xlsx` contra el Excel de inventario (hoja ALMACEN o la que
   elijas); detecta VIN mal capturados y dice en qué otra hoja aparece cada unidad.
5. **Pólizas**: seguimiento de las pólizas de seguro de los créditos. Subes los PDF / ZIP de las pólizas (Crédito
   Click, columna Impresión) y/o el Excel `Polizas_extraidas.xlsx`; con ambos, los PDFs nuevos se juntan con los
   clientes del Excel. Muestra a quién llamar **hoy** (póliza externa hasta 45 días antes del aniversario, contado de
   30 días antes al día anterior, fin de crédito y clientes con el auto ya liberado), las oportunidades **por mes**
   (12 meses), la lista completa con filtros y una tabla de **seguimiento** (Contactado, Fecha, Resultado, Notas).
   Descarga el **Excel con el seguimiento** (hojas *Contacto* y *Base de datos*) y el **PDF para contactar**. Los
   datos de clientes solo viven en la sesión: guarda el Excel y súbelo la próxima vez. Leer PDFs necesita
   `pdftotext` (`packages.txt` → `poppler-utils`).
6. **Mantenimiento**: *Rehacer dashboard y formato* (pestaña Resumen del Sheet + formato y mantenimiento, ver
   abajo) e Historial de cambios.
7. **KPIs de capacitación**: avance de capacitación de Ventas y Posventa.

Las lecturas del Sheet se reusan entre clics durante 5 minutos (cuota de Google); se vuelven a leer al escribir o
con el botón **Recargar** junto al nombre del Sheet.

El cuadre de Sale-U compara contra **todos los créditos de todas las financieras y estatus**: la pestaña *Bitacora*
más *MG Colima Rechazos* y *Créditos Finalizados* (los Finalizados entran por su *FECHA DE ENTREGA*). Con el reporte
**Fuente** los que faltan se separan en *Dar de alta* y *Registrar crédito*; sus leads se guardan en el Sheet
(pestaña oculta *Fuente Sale-U*) y se suman a los ya guardados: sube una vez el de enero a la fecha y después basta
el del mes.

   Cada cuadre genera un **Excel** y un **PDF** descargables (`Cuadre_<tipo>_<periodo>_<fecha>`). El PDF abre con
   las cifras clave y solo trae las secciones con pendientes; el Excel usa el estilo del Sheet (Calibri 9),
   nombres de columna legibles y teléfonos como texto.

## Estructura

```
streamlit_app.py               # entrada: contraseña opcional y navegación
ui.py                          # estilos y componentes compartidos
kpis_core.py                   # lógica de KPIs de capacitación
paginas/kpis_capacitacion.py   # página de KPIs
paginas/inicio.py              # Inicio: indicadores, prioridades, archivos que faltan, tablero y PDF
tablero.py                     # gráficas interactivas (Altair) del Inicio
cuadres/columnas.py            # orden de las primeras columnas (fecha, nombre, teléfono) y sin MARCA
paginas/actualizar.py          # Actualizar: todos los archivos del día, revisar y aplicar
paginas/seguimiento.py         # Seguimiento: resumen, vencidos (PDF por asesor), faltan en Sale-U y pendientes
paginas/blueservice.py         # cuadre BlueService vs inventario
paginas/mantenimiento.py       # dashboard, formato y mantenimiento del Sheet; historial
paginas/polizas.py             # Pólizas: a quién llamar hoy, por mes, seguimiento, Excel y PDF
cuadres/polizas.py             # lee los PDF de pólizas (GNP, Quálitas, Chubb, HDI, Inbursa) y arma el Excel
                               #   (también corre suelto en Colab: botón para subir archivos)
cuadres/polizas_seguimiento.py # estatus de cada crédito por fecha, proyección por mes, seguimiento y PDF
packages.txt                   # poppler-utils (pdftotext) para Streamlit Cloud
comunes.py                     # utilidades de las páginas (tablas, descargas, cuadre de Sale-U)
cuadres/detectar.py            # reconoce cada archivo subido
cuadres/cargas.py              # lo último que se subió (pestaña oculta Cargas) y avisos
cuadres/reporte_ejecutivo.py   # PDF resumen ejecutivo de una página
cuadres/reporte_vencidos.py    # PDF de seguimientos vencidos por asesor
conexion.py                    # conexión con el Google Sheet (o modo Excel)
cuadres/base.py                # leer/escribir el Sheet (Google o .xlsx), respaldo e Historial
cuadres/fuentes.py             # bitácoras de Cetelem y Bitácora de F&I → solicitudes
cuadres/actualizar.py          # cambios propuestos y cómo se aplican
cuadres/saleu_sheet.py         # registro en Sale-U → columna del Sheet
cuadres/pendientes.py          # limpiar pendientes resueltos de la pestaña Revision y archivarlos
cuadres/ids.py                 # ID fijo y marca de reingreso en MG Colima Rechazos
cuadres/formato_sheet.py       # ajustes de estructura y formato del Sheet (botón «Rehacer dashboard y formato»)
cuadres/dashboard.py           # dashboard de la pestaña Resumen (KPI, lectura, gráficas)
cuadres/seguimiento.py         # cuadre vs Excel de seguimiento (ya no está en la app; sigue para Colab)
cuadres/resumen.py             # indicadores del Sheet y clientes que faltan en Sale-U
cuadres/reporte_saleu.py       # PDF «Clientes por registrar en Sale-U»
assets/fonts/                  # Archivo (SIL OFL) e iconos Font Awesome Free (licencias en la carpeta)
cuadres/saleu.py               # cuadre vs Sale-U
cuadres/saleu_fuente.py        # reporte Fuente de Sale-U (todos los leads): falta alta vs falta crédito
cuadres/blueservice.py         # cuadre BlueService vs Inventario
cuadres/salidas.py             # formato común de los Excel y PDF de los cuadres
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
7. Abre la página **Inicio**: debe decir *Conectado a …*.

### Cómo se usa

1. **Actualizar** → sube juntos los archivos que tengas (capturas de Cetelem, Bitácora de F&I si no está conectada
   por Google, reportes de Crédito y Fuente de Sale-U) → *Revisar* → desmarca lo que no quieras → *Aplicar todo*.
   Las bitácoras también actualizan la pestaña *Bitacora*, que es la que usan los cuadres. Lo desmarcado queda como "Ignorado" en el Historial y no se vuelve a
   proponer.
   El cuadre de Sale-U (en el mismo *Revisar*) compara contra la **bitácora completa** (todas las financieras y
   estatus; ver abajo) y se aplica después de las bitácoras.
   Un cliente está en Sale-U si coincide el
   teléfono (aunque el nombre esté capturado distinto; se avisa para verificar) o el nombre es muy parecido; es el
   mismo criterio que usa la columna *Registro en Sale-U*. La hoja *Faltan en Sale-U* del Excel trae
   nombre, vendedor, celular, financiera, estatus, plazo, unidad, mensualidad y correo de cada cliente para darlo de
   alta. Además del Excel/PDF, propone el *Registro en Sale-U* de cada rechazo (y de cada aprobado y rechazado de la *Bitacora*; los
   rechazados ligados a *MG Colima Rechazos* toman el mismo valor que su fila allá)
   cuya fecha cae en el periodo del reporte de Crédito: *Registrado* (está en Crédito), *Falta crédito* (no está en
   Crédito pero sí en Fuente) o *No Registrado* (no está en ninguno). Sin el reporte Fuente solo hay *Registrado* /
   *No Registrado*. Fuera del periodo de Crédito, un *No Registrado* que sí está en Fuente sube a *Falta crédito*.
   En los dos reportes el cliente se busca por teléfono **o** por nombre: basta con que coincida uno (si coincide
   el teléfono con otro nombre, o el nombre con otro teléfono, se avisa para verificar).
   Se escribe con *Aplicar todo*.
   Nunca se baja de nivel: lo que ya dice *Registrado* o *Sí* (p. ej. marcado a mano) se deja igual aunque no
   aparezca en los reportes, y se lista para que lo confirmes.
2. **Seguimiento → Resumen** → revisa indicadores y vencidos; *PDF de vencidos* por asesor y *Generar PDF y Excel*
   para mandar a los asesores la lista de quién falta en Sale-U.
3. **Seguimiento → Pendientes** → *Revisar pendientes* → *Limpiar*. Los pendientes de la pestaña
   *Revision* que ya tienen el dato capturado pasan a *Corregido* y se les quita el morado y la nota en la
   hoja. Si un pendiente no aplica o lo resolviste de otra forma, cámbialo a mano a *Corregido* o *No aplica*
   en *Revision* y el botón le quita la marca. Los resueltos pasan a la pestaña *Revision resueltos*; en
   *Revision* quedan solo los pendientes.

**Formato del Sheet (una vez):** *Mantenimiento → Rehacer dashboard y formato*. Guarda respaldo y, si ya está
aplicado, no duplica nada:
- Primeras columnas: *MG Colima Rechazos* **Fecha ingreso, Nombre de Cliente, Celular, Folio**; *Bitacora* **Fecha,
  Cliente, Teléfono, Folio**; *Créditos Finalizados* sin **MARCA** y con **NOMBRE, TELEFONO** al inicio. Se mueven
  columnas completas (datos, formato, listas, colores y fórmulas se ajustan) y se congelan fecha, nombre y teléfono.
- *MG Colima Rechazos*: al final, **Reingreso** («1 de 2», «2 de 2» cuando el celular está en otra solicitud, en
  amarillo) e **ID** (R-0001…, no cambia aunque se ordene la hoja; la *Bitacora* y *Revision* guardan el ID en lugar
  del número de fila). La app le pone ID a las filas nuevas o capturadas a mano cada vez que aplica bitácoras.
- *CONTESTA 1/2/3*: opción «No responde» y color naranja cuando hay fecha de contacto pero no se capturó si contestó.
- *Tipo de Subvención*, *Accesorios* y *Status 1* (siempre traen lo mismo) quedan agrupadas y contraídas (se abren
  con el «+» de arriba).
- *MG Colima Créditos Finalizados*: columna **FECHA DE ENTREGA**; con ella cada crédito entra al cuadre de Sale-U
  solo en su periodo.
- La columna de Sale-U se llama **Registro en Sale-U** en las dos pestañas (antes *Cliente con Registro en SalesU* y
  *CLIENTE CON REGISTRO EN SALE-U*). La app lee los dos nombres.
- Fechas de *Historial* y *Bitacora F&I* con formato de fecha. *Capturas Cetelem* y *Bitacora F&I* quedan ocultas
  (la app las sigue usando).
- **Mantenimiento**: un solo respaldo por pestaña (el último; se borran los anteriores y las copias ocultas
  «Hoja N»); *Tipo de Plan*, *Tipo de Subvención*, *Accesorios* y *Status 1* se llenan solos donde estén vacíos (también
  en los rechazos nuevos de *Actualizar*); *Créditos Finalizados* lleva **ID en 'MG Colima Rechazos'** (su rechazo
  previo, por teléfono o nombre); columnas calculadas de Rechazos (Fecha de Seguimiento, Efectividad, Reingreso, ID,
  Registro en Sale-U) y la *Bitacora* completa protegidas con **solo advertencia**; encabezado, fecha, nombre y teléfono congelados (en Finalizados, nombre y teléfono);
  una **vista filtrada por asesor** (*Datos → Vistas filtradas*); *Revision resueltos* oculta; pestañas en orden
  Resumen, Rechazos, Finalizados, Bitacora, Revision.
- El cuadre de Sale-U con el reporte **Fuente** también propone el **correo** que falte (del lead en Sale-U).

**La pestaña *Bitacora*** es el acumulado de los créditos de todas las financieras y **todos los estatus** (formalizados incluidos). Su columna **Registro en Sale-U** usa los mismos valores y colores que Rechazos y Finalizados: *Registrado* / *Falta crédito* / *No Registrado* para Aprobados y Rechazados (en los rechazos ligados por ID, lo que dice *MG Colima Rechazos*; si no, lo que dice F&I) y *No aplica* para los demás estatus. Se
rehace cada vez que aplicas en *Actualizar* con: las capturas de Cetelem (cada mes que subes reemplaza
al mismo mes), la Bitácora de F&I (se **acumula**: lo que trae la carga nueva reemplaza al mismo registro y se
conserva lo que ya no venga en el archivo) y los rechazos de *MG Colima Rechazos* que no salen de ninguna de ellas
(columna *Fuente*: «MG Colima Rechazos R-0005»; la columna *ID en 'MG Colima Rechazos'* liga cada solicitud con su rechazo). El vendedor se escribe con su nombre completo de *Catalogos*
(«CARLA» → «CARLA YOVET CARDENAS AHUMADA») cuando corresponde a un solo asesor. Toda la pestaña queda con el mismo
estilo (Calibri 9, encabezado gris), fechas dd/mm/aaaa y montos con $. El cuadre de *Sale-U* cuenta también los
formalizados.

**La pestaña *Resumen*** es un dashboard: indicadores del año (rechazos, seguimiento al día —de los que ya cumplieron
180 días, cuántos tienen primer seguimiento—, contacto
efectivo, % registrado en Sale-U, seguimientos vencidos, pendientes en Revision), una *Lectura del periodo* con frases
que interpretan la situación, 8 gráficas (rechazos por mes y financiera, seguimiento, por financiera, por asesor,
registro en Sale-U, pendientes, bitácora por estatus y % de rechazo) y, debajo, las tablas de donde salen. Todo son
fórmulas: se actualiza solo y el año se cambia en la celda B2. La app lo rehace cada vez que aplicas bitácoras (para
tomar financieras y asesores nuevos) o con *Mantenimiento → Rehacer dashboard y formato*.

Reglas que sigue la app al escribir en el Sheet:
- **Datos faltantes:** si un cliente de *MG Colima Rechazos* o *MG Colima Créditos Finalizados* cuadra por
  **teléfono** (mismo celular y nombre parecido, al menos la mitad de las palabras) o por **nombre** con la bitácora o con
  Sale-U, sus celdas **vacías** se llenan (teléfono, correo, folio, plazo, mensualidad, unidad, asesor, GAP,
  accesorios y, en Finalizados, el registro en Sale-U). Nunca se sobreescribe un dato. Si la bitácora y Sale-U
  traen el mismo dato, manda la bitácora (la mensualidad de Sale-U es aproximada). En *Actualizar*
  salen como "Dato faltante" para que los revises; en el cuadre de Sale-U se escriben al pulsar *Aplicar*.
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
