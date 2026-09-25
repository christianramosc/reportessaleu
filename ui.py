# -*- coding: utf-8 -*-
"""Estilos y componentes compartidos por todas las páginas."""
import html

import streamlit as st

import kpis_core as kc


def aplicar_estilo():
    st.markdown(f"""
<style>
@import url('https://fonts.googleapis.com/css2?family=Archivo:wght@400;500;600;700;800&display=swap');
.hero, .kpi, .alerta, .vacio {{ font-family: 'Archivo', sans-serif; }}
.block-container {{ padding-top: 1.6rem; max-width: 1280px; }}

.hero {{ background: {kc.AZUL_OSCURO}; color: #fff; border-radius: 14px; padding: 22px 28px;
        display: flex; justify-content: space-between; align-items: flex-end; gap: 16px;
        flex-wrap: wrap; margin-bottom: 18px; }}
.hero h1 {{ font-size: 1.9rem; font-weight: 800; margin: 0; color: #fff; line-height: 1.1; }}
.hero p {{ margin: 6px 0 0; color: #C9D6EA; font-size: .95rem; }}
.hero .file {{ color: #C9D6EA; font-size: .85rem; text-align: right; }}

.kpis {{ display: grid; grid-template-columns: 1.25fr repeat(4, 1fr); gap: 12px; margin: 4px 0 18px; }}
@media (max-width: 900px) {{ .kpis {{ grid-template-columns: repeat(2, 1fr); }} }}
.kpi {{ background: #fff; border: 1px solid #E3E8EF; border-radius: 12px; padding: 14px 16px;
       border-top: 5px solid var(--c); }}
.kpi .n {{ font-size: .9rem; color: {kc.TEXTO_2}; font-weight: 600; }}
.kpi .v {{ font-size: 2rem; font-weight: 800; color: {kc.TEXTO}; line-height: 1.15; }}
.kpi .s {{ font-size: .8rem; color: {kc.TEXTO_2}; }}
.kpi .bar {{ height: 6px; background: #E7EBF1; border-radius: 4px; margin: 8px 0 6px; overflow: hidden; }}
.kpi .bar > div {{ height: 100%; background: var(--c); border-radius: 4px; }}
.kpi.total {{ background: {kc.AZUL_OSCURO}; border-color: {kc.AZUL_OSCURO}; }}
.kpi.total .n, .kpi.total .s {{ color: #C9D6EA; }}
.kpi.total .v {{ color: #fff; font-size: 2.4rem; }}
.kpi.total .bar {{ background: rgba(255,255,255,.18); }}

.alerta {{ border-left: 4px solid {kc.AMBAR}; background: #FFF9EC; padding: 10px 14px;
          border-radius: 6px; margin-bottom: 8px; color: {kc.TEXTO}; font-size: .93rem; }}
.vacio {{ border: 2px dashed #CBD4E1; border-radius: 14px; padding: 36px; text-align: center;
         color: {kc.TEXTO_2}; }}
.vacio b {{ color: {kc.TEXTO}; font-size: 1.1rem; }}
</style>
""", unsafe_allow_html=True)


def hero(titulo, subtitulo, derecha=""):
    st.markdown(f"""<div class="hero"><div><h1>{html.escape(titulo)}</h1>
<p>{html.escape(subtitulo)}</p></div><div class="file">{html.escape(derecha)}</div></div>""",
                unsafe_allow_html=True)


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
        caja.markdown(f"**{etiqueta}**  \n:material/description: {html.escape(guardado[0])}")
        if caja.button("Cambiar archivo", key=f"quitar_{key}", type="tertiary",
                       icon=":material/swap_horiz:"):
            st.session_state.pop(slot, None)
            st.rerun()
        return guardado
    donde.file_uploader(etiqueta, type=tipos, key=key, help=ayuda, on_change=_al_cambiar)
    return st.session_state.get(slot)
