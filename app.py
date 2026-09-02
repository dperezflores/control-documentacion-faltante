from __future__ import annotations

import streamlit as st

from app_state import AppState
from styles import get_custom_css
from views import CapturaFaltantesView, CortesView, GenerarOficioView
from services.excel_service import list_requirements


st.set_page_config(
    page_title="Control documental",
    layout="wide",
    initial_sidebar_state="expanded",
)

st.markdown(get_custom_css(), unsafe_allow_html=True)


def secret_dict():
    try:
        return dict(st.secrets)
    except Exception:
        return {}


state = AppState(secret_dict())

st.markdown(
    """
    <div class="institutional-header">
        <div class="institutional-eyebrow">Control documental</div>
        <div class="institutional-title">Documentación faltante</div>
        <div class="institutional-subtitle">
            Registro homologado de faltantes, administración de cortes y preparación de solicitudes.
        </div>
    </div>
    """,
    unsafe_allow_html=True,
)

loaded = state.load_files()
if loaded is None:
    if state.storage_mode == "Google Drive":
        st.error(
            "No fue posible localizar uno de los Excel configurados en Google Drive. "
            "Verifica los fileId y que la carpeta esté compartida con la cuenta de servicio."
        )
        st.stop()
    st.subheader("Inicializar archivos del prototipo")
    st.info("Carga una sola vez el Excel de documentación faltante y el Excel de codificación.")
    op_upload = st.file_uploader("Excel de documentación faltante", type=["xlsx"], key="op")
    cat_upload = st.file_uploader("Excel de codificación de documentos", type=["xlsx"], key="cat")
    if st.button("Inicializar prototipo", type="primary", disabled=not (op_upload and cat_upload)):
        try:
            state.initialize_files(op_upload.getvalue(), cat_upload.getvalue())
            st.success("Archivos inicializados correctamente.")
            st.rerun()
        except Exception as exc:
            st.error(f"No fue posible inicializar los archivos: {exc}")
    st.stop()

operational, catalog_bytes = loaded

st.sidebar.markdown("### Navegación")
section = st.sidebar.radio(
    "Sección",
    ["Capturar faltantes", "Cortes", "Generar oficio"],
    label_visibility="collapsed",
)
st.sidebar.divider()
st.sidebar.caption(f"Almacenamiento: {state.storage_mode}")
if st.sidebar.button("Actualizar datos"):
    with st.spinner("Actualizando información..."):
        state.load_files(force=True)
    st.rerun()

requirements = list_requirements(operational)
if not requirements:
    st.warning("No se encontraron hojas de requerimiento.")
    st.stop()

requirement = st.selectbox("Requerimiento", requirements)

views = {
    "Capturar faltantes": CapturaFaltantesView,
    "Cortes": CortesView,
    "Generar oficio": GenerarOficioView,
}
views[section](state, operational, catalog_bytes, requirement).render()
