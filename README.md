# Herramientas MG Colima

App de Streamlit con una página por herramienta:

1. **KPIs de capacitación**: avance de capacitación de Ventas y Posventa.
2. **Cuadres para reportes**: cuadra las capturas mensuales de **Cetelem** (se piden al F&I) y el inventario:
   - **Seguimiento Rechazos Financiera MG vs Capturas CETELEM** (`Seguimiento Rechazos Financiera MG.xlsx`,
     solo pestañas de COLIMA). Cuadra solo las solicitudes **rechazadas**; cruza por folio, teléfono y nombre.
   - **Sale - U Aprobados y Rechazados CETELEM**: reporte de Sale-U (`reporteLeadsconCredito_*.csv`, tal cual
     se descarga). Cruza por teléfono y nombre; compara estatus y financiera.
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

## Notas

- Subir el CSV de Sale-U **sin abrirlo y guardarlo en Excel**: Excel convierte algunos
  teléfonos a notación científica y se pierden dígitos.
- Los módulos de `cuadres/` también funcionan solos en Google Colab (`python cuadres/saleu.py`).
