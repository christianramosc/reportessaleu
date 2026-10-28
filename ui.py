# -*- coding: utf-8 -*-
"""Estilos y componentes compartidos por todas las páginas (tema oscuro al estilo shadcn/ui, paleta zinc)."""
import html

import streamlit as st

# Tokens del tema: los mismos de .streamlit/config.toml
FONDO, CARTA, CARTA_2 = "#09090b", "#0f0f12", "#18181b"
BORDE, BORDE_2 = "#27272a", "#3f3f46"
TEXTO, TEXTO_2, TENUE = "#fafafa", "#d4d4d8", "#a1a1aa"
PRIMARIO = "#3b82f6"
ESTADO = dict(bien="#22c55e", aviso="#eab308", critico="#ef4444", info="#60a5fa")
FUENTE = "Geist, Inter, system-ui, sans-serif"


def aplicar_estilo():
    st.markdown(f"""
<style>
:root {{ --fondo:{FONDO}; --carta:{CARTA}; --carta-2:{CARTA_2}; --borde:{BORDE}; --borde-2:{BORDE_2};
        --texto:{TEXTO}; --texto-2:{TEXTO_2}; --tenue:{TENUE}; --primario:{PRIMARIO}; }}
.block-container {{ padding-top: 5rem; padding-bottom: 4rem; max-width: 1320px; }}
header[data-testid="stHeader"] {{ background: rgba(9,9,11,.82); backdrop-filter: blur(8px);
                                  border-bottom: 1px solid var(--borde); }}

/* encabezado de página */
.encabezado {{ display: flex; justify-content: space-between; align-items: flex-end; gap: 16px; flex-wrap: wrap;
              padding: 6px 0 16px; margin-bottom: 14px; border-bottom: 1px solid var(--borde); }}
.encabezado h1 {{ font-size: 1.85rem; font-weight: 700; letter-spacing: -.025em; margin: 0; padding: 0;
                 line-height: 1.15; color: var(--texto); }}
.encabezado p {{ margin: 6px 0 0; color: var(--tenue); font-size: .95rem; }}
.encabezado .eti {{ display: inline-flex; align-items: center; gap: 6px; font-size: .8rem; color: var(--texto-2);
                   border: 1px solid var(--borde); border-radius: 999px; padding: 4px 12px; background: var(--carta-2); }}
.encabezado .ico {{ display: inline-flex; width: 34px; height: 34px; border-radius: 9px; align-items: center;
                   justify-content: center; background: var(--carta-2); border: 1px solid var(--borde);
                   margin-right: 12px; vertical-align: middle; }}

/* títulos de sección */
.seccion {{ margin: 22px 0 10px; }}
.seccion h3 {{ font-size: 1.12rem; font-weight: 600; letter-spacing: -.01em; margin: 0; padding: 0; color: var(--texto); }}
.seccion p {{ margin: 3px 0 0; color: var(--tenue); font-size: .86rem; }}
.seccion .num {{ display: inline-flex; width: 22px; height: 22px; border-radius: 6px; background: var(--carta-2);
                border: 1px solid var(--borde); font-size: .75rem; align-items: center; justify-content: center;
                margin-right: 8px; color: var(--texto-2); vertical-align: 2px; }}

/* insignias y texto tenue */
.insignia {{ display: inline-flex; align-items: center; gap: 6px; font-size: .78rem; font-weight: 500;
            border: 1px solid var(--borde); border-radius: 999px; padding: 2px 10px; color: var(--texto-2); }}
.insignia .p {{ width: 7px; height: 7px; border-radius: 50%; background: var(--c, {ESTADO['bien']}); }}
.tenue {{ color: var(--tenue); font-size: .86rem; }}

/* tarjetas de indicadores */
.kv {{ font-size: 2rem !important; font-weight: 650 !important; letter-spacing: -.02em; line-height: 1.1 !important;
       margin: 2px 0 0 !important; font-variant-numeric: tabular-nums; }}
.ks {{ font-size: .8rem !important; color: var(--tenue); margin: 4px 0 2px; white-space: nowrap; overflow: hidden;
       text-overflow: ellipsis; }}
.kt {{ font-size: .86rem; color: var(--tenue); font-weight: 500; display: flex; align-items: center; gap: 6px; }}

div[data-testid="stMetric"] {{ background: var(--carta); border: 1px solid var(--borde); padding: .75rem 1rem;
                               border-radius: .6rem; }}
div[data-testid="stMetric"] label p {{ color: var(--tenue); font-size: .82rem; }}
div[data-testid="stMetricValue"] {{ font-variant-numeric: tabular-nums; letter-spacing: -.02em; }}

/* contenedores con borde = cards */
div[data-testid="stVerticalBlockBorderWrapper"] {{ background: var(--carta); }}
div[data-testid="stExpander"] details {{ background: var(--carta); border-color: var(--borde); }}
div[data-testid="stExpander"] summary:hover {{ background: var(--carta-2); }}

/* botones */
.stButton button, .stDownloadButton button, .stFormSubmitButton button {{ font-weight: 500; min-height: 2.4rem;
    transition: background .15s, border-color .15s; }}
.stButton button[kind="secondary"], .stDownloadButton button[kind="secondary"] {{ background: var(--carta-2); }}
.stButton button[kind="secondary"]:hover, .stDownloadButton button[kind="secondary"]:hover {{
    border-color: var(--borde-2); background: #202024; color: var(--texto); }}

/* pestañas como en shadcn: lista con fondo y la activa resaltada */
div[data-testid="stTabs"] [role="tablist"] {{ gap: 4px; background: var(--carta-2); padding: 4px; border-radius: .6rem;
    width: fit-content; max-width: 100%; overflow-x: auto; border: 1px solid var(--borde); margin-bottom: 8px; }}
div[data-testid="stTab"] {{ height: 2.1rem; padding: 0 14px; border-radius: .45rem; color: var(--tenue);
                            transition: background .15s, color .15s; }}
div[data-testid="stTab"]:hover {{ color: var(--texto); }}
div[data-testid="stTab"][aria-selected="true"] {{ background: var(--fondo); color: var(--texto);
                                                  box-shadow: 0 1px 3px rgba(0,0,0,.5); }}
div[data-testid="stTab"] > div:not([data-testid]) {{ display: none; }}

/* subida de archivos */
section[data-testid="stFileUploaderDropzone"] {{ background: var(--carta); border: 1px dashed var(--borde-2); }}

/* pasos de la semana */
.pasos {{ display: flex; flex-wrap: wrap; gap: 8px; align-items: center; font-size: .82rem; color: var(--tenue);
         margin: -2px 0 16px; }}
.pasos > span.paso {{ display: inline-flex; align-items: center; gap: 8px; border: 1px solid var(--borde);
                     border-radius: 999px; padding: 3px 12px 3px 4px; background: var(--carta); }}
.pasos b {{ color: var(--texto); font-weight: 500; }}
.pasos .n {{ display: inline-flex; width: 20px; height: 20px; border-radius: 50%; background: var(--primario);
            color: #fff; font-size: .72rem; font-weight: 600; align-items: center; justify-content: center; }}

/* KPIs de capacitación */
.kpis {{ display: grid; grid-template-columns: 1.25fr repeat(4, 1fr); gap: 12px; margin: 4px 0 18px; }}
@media (max-width: 900px) {{ .kpis {{ grid-template-columns: repeat(2, 1fr); }} }}
.kpi {{ background: var(--carta); border: 1px solid var(--borde); border-radius: .75rem; padding: 14px 16px; }}
.kpi .n {{ font-size: .85rem; color: var(--tenue); font-weight: 500; display: flex; align-items: center; gap: 8px; }}
.kpi .n::before {{ content: ""; width: 8px; height: 8px; border-radius: 50%; background: var(--c); }}
.kpi .v {{ font-size: 1.9rem; font-weight: 650; color: var(--texto); line-height: 1.15; letter-spacing: -.02em; }}
.kpi .s {{ font-size: .78rem; color: var(--tenue); }}
.kpi .bar {{ height: 6px; background: var(--carta-2); border-radius: 4px; margin: 8px 0 6px; overflow: hidden; }}
.kpi .bar > div {{ height: 100%; background: var(--c); border-radius: 4px; }}
.kpi.total {{ background: linear-gradient(180deg, #111827 0%, var(--carta) 100%); border-color: #1e3a8a; }}
.kpi.total .v {{ font-size: 2.3rem; }}

.alerta {{ border: 1px solid rgba(234,179,8,.35); background: rgba(234,179,8,.08); padding: 10px 14px;
          border-radius: .5rem; margin-bottom: 8px; color: var(--texto-2); font-size: .9rem; }}
.vacio {{ border: 1px dashed var(--borde-2); border-radius: .75rem; padding: 40px; text-align: center;
         color: var(--tenue); background: var(--carta); }}
.vacio b {{ color: var(--texto); font-size: 1.05rem; font-weight: 600; }}
</style>
""", unsafe_allow_html=True)


def encabezado(titulo, subtitulo, derecha="", icono=None):
    """Título de la página con su descripción y, a la derecha, una etiqueta (fecha de corte, archivo…)."""
    ico = f'<span class="ico">{html.escape(icono)}</span>' if icono else ""
    eti = f'<span class="eti">{html.escape(derecha)}</span>' if derecha else ""
    st.markdown(f"""<div class="encabezado"><div><h1>{ico}{html.escape(titulo)}</h1>
<p>{html.escape(subtitulo)}</p></div><div>{eti}</div></div>""", unsafe_allow_html=True)



def seccion(titulo, descripcion="", numero=None, donde=st):
    """Título de una sección con su descripción en gris."""
    num = f'<span class="num">{numero}</span>' if numero is not None else ""
    desc = f"<p>{html.escape(descripcion)}</p>" if descripcion else ""
    donde.markdown(f'<div class="seccion"><h3>{num}{html.escape(titulo)}</h3>{desc}</div>', unsafe_allow_html=True)


def insignia(texto, color=None):
    """Píldora pequeña con un punto de color (estado)."""
    return (f'<span class="insignia" style="--c:{color or ESTADO["bien"]}"><span class="p"></span>'
            f'{html.escape(texto)}</span>')


def archivo_persistente(etiqueta, key, tipos, donde=st, ayuda=None):
    """file_uploader que conserva el archivo al cambiar de página.

    Devuelve (nombre, bytes) o None. Streamlit borra los uploaders al cambiar de
    página; aquí el archivo se guarda en session_state y solo se quita cuando el
    usuario lo elimina del uploader o pulsa "Quitar".
    """
    slot = f"_archivo_{key}"

    def _al_cambiar():
        f = st.session_state.get(key)
        if f is None:
            st.session_state.pop(slot, None)
        else:
            st.session_state[slot] = (f.name, f.getvalue())

    guardado = st.session_state.get(slot)
    if guardado and st.session_state.get(key) is None:
        # viene de otra página: mostrar el archivo guardado en lugar del uploader vacío
        caja = donde.container(border=True)
        c1, c2 = caja.columns([4, 1], vertical_alignment="center")
        c1.markdown(f"**{etiqueta}**  \n:material/description: {html.escape(guardado[0])}")
        if c2.button("Cambiar", key=f"quitar_{key}", icon=":material/swap_horiz:", width="stretch"):
            st.session_state.pop(slot, None)
            st.rerun()
        return guardado
    donde.file_uploader(etiqueta, type=tipos, key=key, help=ayuda, on_change=_al_cambiar)
    return st.session_state.get(slot)


PASOS_MES = [("Actualizar", "subir archivos y aplicar"), ("Seguimiento", "Resumen"), ("Seguimiento", "Pendientes")]


def guia_mes():
    """Orden del cierre semanal, para no saltarse pasos."""
    pasos = "".join(f'<span class="paso"><span class="n">{i}</span>{html.escape(pagina)} · '
                    f'<b>{html.escape(pestana)}</b></span>' for i, (pagina, pestana) in enumerate(PASOS_MES, 1))
    st.markdown(f'<div class="pasos"><span>Cada semana</span>{pasos}</div>', unsafe_allow_html=True)


def tarjeta(donde, icono, titulo, valor, detalle="", color=None):
    """Indicador con icono de Material, número grande y detalle (el color se usa solo para estados)."""
    caja = donde.container(border=True)
    caja.markdown(f":gray[:material/{icono}: {titulo}]")
    caja.markdown(f'<div class="kv" style="color:{color or TEXTO}">{html.escape(str(valor))}</div>'
                  f'<div class="ks" title="{html.escape(detalle)}">{html.escape(detalle) or "&nbsp;"}</div>',
                  unsafe_allow_html=True)


def barra(donde=st, alinear="right"):
    """Fila de botones (acciones y descargas) juntos y alineados, a la derecha por defecto."""
    return donde.container(horizontal=True, horizontal_alignment=alinear, vertical_alignment="bottom", gap="small")
