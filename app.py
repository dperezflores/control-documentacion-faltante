from __future__ import annotations

from datetime import date

import streamlit as st

from services.excel_service import (
    add_faltantes,
    create_cut,
    document_history,
    documents_for_cuts,
    list_contracts,
    list_cuts,
    list_requirements,
    load_catalog,
    pending_summary,
)
from services.github_service import build_store
from services.word_service import build_request_docx

st.set_page_config(page_title="Control documental", layout="wide")

DATA_FILE = "data/Documentacion_faltante.xlsx"
CATALOG_FILE = "data/Codificacion_documentos.xlsx"

CUSTOM_CSS = """
<style>
    .stApp { background: #f7f8fa; }
    .block-container { max-width: 1180px; padding-top: 2.2rem; padding-bottom: 3rem; }
    h1 { color: #202124; font-size: 2rem !important; letter-spacing: -0.02em; margin-bottom: .25rem !important; }
    h2, h3 { color: #2f3337; letter-spacing: -0.01em; }
    [data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid #e7e9ed; }
    [data-testid="stSidebar"] .stRadio > label { font-weight: 700; }
    div[data-testid="stSelectbox"], div[data-testid="stMultiSelect"], div[data-testid="stTextInput"] {
        margin-bottom: .25rem;
    }
    div[data-testid="stMetric"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
        padding: 14px 16px;
    }
    div[data-testid="stExpander"] {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 10px;
    }
    .info-card {
        background: #ffffff;
        border: 1px solid #e5e7eb;
        border-radius: 12px;
        padding: 18px 20px;
        margin: 8px 0 18px 0;
    }
    .info-label { color: #6b7280; font-size: .78rem; font-weight: 700; text-transform: uppercase; letter-spacing: .04em; }
    .info-value { color: #25282c; font-size: .98rem; margin: 3px 0 10px 0; }
    .section-kicker { color: #6b7280; font-size: .86rem; margin-bottom: 1rem; }
    .stButton > button[kind="primary"] { border-radius: 8px; font-weight: 700; }
</style>
"""
st.markdown(CUSTOM_CSS, unsafe_allow_html=True)


def secret_dict():
    try:
        return dict(st.secrets)
    except Exception:
        return {}


store, storage_mode = build_store(secret_dict())


def load_session_files(force: bool = False) -> tuple[bytes, bytes] | None:
    """Carga desde GitHub sólo al iniciar sesión o cuando el usuario pide actualizar."""
    if force:
        st.session_state.pop("operational_bytes", None)
        st.session_state.pop("catalog_bytes", None)

    if "operational_bytes" in st.session_state and "catalog_bytes" in st.session_state:
        return st.session_state.operational_bytes, st.session_state.catalog_bytes

    try:
        operational = store.read(DATA_FILE).content
        catalog = store.read(CATALOG_FILE).content
        st.session_state.operational_bytes = operational
        st.session_state.catalog_bytes = catalog
        return operational, catalog
    except FileNotFoundError:
        return None
    except Exception as exc:
        st.error(f"No fue posible cargar los archivos de datos: {exc}")
        st.stop()


st.title("Control de documentación faltante")
st.markdown('<div class="section-kicker">Registro homologado de faltantes, control de cortes y preparación de solicitudes.</div>', unsafe_allow_html=True)

loaded = load_session_files()
if loaded is None:
    st.subheader("Inicializar archivos del prototipo")
    st.info("Carga una sola vez el Excel de documentación faltante y el Excel de codificación. A partir de ese momento la aplicación trabajará sobre los archivos guardados en GitHub.")
    op_upload = st.file_uploader("Excel de documentación faltante", type=["xlsx"], key="op")
    cat_upload = st.file_uploader("Excel de codificación de documentos", type=["xlsx"], key="cat")
    if st.button("Inicializar prototipo", type="primary", disabled=not (op_upload and cat_upload)):
        try:
            try:
                store.read(DATA_FILE)
            except FileNotFoundError:
                store.write_new(DATA_FILE, op_upload.getvalue(), "Inicializar Excel operativo")
            try:
                store.read(CATALOG_FILE)
            except FileNotFoundError:
                store.write_new(CATALOG_FILE, cat_upload.getvalue(), "Inicializar catálogo de codificación")
            st.session_state.operational_bytes = op_upload.getvalue()
            st.session_state.catalog_bytes = cat_upload.getvalue()
            st.success("Archivos inicializados correctamente.")
            st.rerun()
        except Exception as exc:
            st.error(f"No fue posible inicializar los archivos: {exc}")
    st.stop()

operational, catalog_bytes = loaded

st.sidebar.markdown("### Navegación")
section = st.sidebar.radio("Sección", ["Capturar faltantes", "Cortes", "Generar oficio"], label_visibility="collapsed")
st.sidebar.divider()
st.sidebar.caption(f"Almacenamiento: {storage_mode}")
if st.sidebar.button("Actualizar datos"):
    with st.spinner("Actualizando información..."):
        operational, catalog_bytes = load_session_files(force=True)
    st.rerun()

requirements = list_requirements(operational)
if not requirements:
    st.warning("No se encontraron hojas de requerimiento.")
    st.stop()

requirement = st.selectbox("Requerimiento", requirements)

if section == "Capturar faltantes":
    contracts = list_contracts(operational, requirement)
    if not contracts:
        st.info("No hay contratos disponibles en este requerimiento.")
        st.stop()

    auditors = sorted({str(c["auditor"]).strip() for c in contracts if str(c["auditor"]).strip()})
    f1, f2 = st.columns([1, 2.4])
    with f1:
        auditor_filter = st.selectbox("Auditor", ["Todos"] + auditors)
    filtered_contracts = contracts if auditor_filter == "Todos" else [c for c in contracts if str(c["auditor"]).strip() == auditor_filter]

    with f2:
        labels = {f"{c['contrato']} · {str(c['obra'])[:78]}": c for c in filtered_contracts}
        selected_label = st.selectbox("Contrato", list(labels.keys()))
    contract = labels[selected_label]

    st.markdown(
        f"""
        <div class="info-card">
          <div class="info-label">Obra</div>
          <div class="info-value">{contract['obra']}</div>
          <div class="info-label">Contratista</div>
          <div class="info-value">{contract['contratista']}</div>
          <div class="info-label">Auditor asignado</div>
          <div class="info-value">{contract['auditor'] or 'Sin asignar'}</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    procedure = st.selectbox(
        "Procedimiento de adjudicación",
        ["DIR", "LPU", "LSI"],
        format_func=lambda x: {
            "DIR": "DIR · Adjudicación Directa",
            "LPU": "LPU · Licitación Pública",
            "LSI": "LSI · Licitación Simplificada",
        }[x],
    )

    catalog = load_catalog(catalog_bytes, procedure)
    history = document_history(operational, requirement, contract["contrato"])

    # Los documentos que ya están pendientes no se duplican antes de crear un corte.
    # Los históricos sí vuelven a mostrarse para permitir reiterarlos en un corte posterior.
    available = [x for x in catalog if not history.get(x["codigo"], {}).get("pending", False)]
    options = {}
    for item in available:
        prior_cuts = history.get(item["codigo"], {}).get("cuts", [])
        label = item["documento"]
        if prior_cuts:
            label += f"  ·  Ya solicitado en corte(s): {', '.join(map(str, prior_cuts))}"
        options[label] = item

    st.subheader("Documentación faltante")
    st.caption("Selecciona uno o varios documentos del catálogo correspondiente al procedimiento.")
    selected_labels = st.multiselect("Documentos", list(options.keys()), placeholder="Seleccionar documentos")
    selected_items = [options[x] for x in selected_labels]

    repeated = []
    for item in selected_items:
        cuts = history.get(item["codigo"], {}).get("cuts", [])
        if cuts:
            repeated.append((item, cuts))

    repeat_confirmed = True
    if repeated:
        names = "; ".join(f"{item['documento']} (corte(s) {', '.join(map(str, cuts))})" for item, cuts in repeated)
        st.warning(f"La selección incluye documentación ya solicitada anteriormente: {names}. Si no fue entregada, puedes volver a solicitarla en el siguiente corte.")
        repeat_confirmed = st.checkbox("Confirmo que deseo volver a agregar esta documentación para una nueva solicitud.")

    auditor_for_save = str(contract["auditor"] or "").strip()
    if not auditor_for_save:
        auditor_for_save = st.text_input("Usuario / iniciales", help="Este contrato no tiene auditor asignado en el Excel. Indica quién registra los faltantes.").strip()

    if st.button("Guardar documentación faltante", type="primary", disabled=not selected_items or not repeat_confirmed):
        if not auditor_for_save:
            st.error("Este contrato no tiene auditor asignado. Indica el usuario o iniciales para continuar.")
        else:
            def mutation(latest: bytes) -> bytes:
                return add_faltantes(latest, requirement, contract["contrato"], procedure, auditor_for_save, selected_items)

            try:
                updated = store.mutate(DATA_FILE, mutation, f"Registrar faltantes {requirement} · {contract['contrato']}")
                st.session_state.operational_bytes = updated
                st.success(f"Se registraron {len(selected_items)} documento(s) para el siguiente corte.")
                st.rerun()
            except Exception as exc:
                st.error(f"No fue posible guardar: {exc}")

    existing_count = sum(1 for item in history.values() if item.get("cuts") or item.get("pending"))
    if existing_count:
        with st.expander(f"Historial del contrato ({existing_count} documentos)"):
            rows = []
            catalog_by_code = {x["codigo"]: x["documento"] for x in catalog}
            for code, item in history.items():
                status = "Pendiente de corte" if item.get("pending") else ""
                cuts = ", ".join(map(str, item.get("cuts", []))) or "—"
                rows.append({
                    "Documento": item.get("documento") or catalog_by_code.get(code, code),
                    "Cortes anteriores": cuts,
                    "Estado actual": status or "Sin solicitud pendiente",
                })
            st.dataframe(rows, use_container_width=True, hide_index=True)

elif section == "Cortes":
    pending = pending_summary(operational, requirement)
    cuts = list_cuts(operational, requirement)

    m1, m2 = st.columns(2)
    m1.metric("Pendientes de corte", len(pending))
    m2.metric("Cortes existentes", len(cuts))

    if cuts:
        st.subheader("Cortes registrados")
        st.dataframe(cuts, use_container_width=True, hide_index=True)

    st.subheader("Documentación pendiente")
    if pending:
        st.dataframe(pending, use_container_width=True, hide_index=True)
        c1, c2 = st.columns(2)
        with c1:
            cut_date = st.date_input("Fecha del nuevo corte", value=date.today())
        with c2:
            cut_user = st.text_input("Usuario / iniciales", help="Persona que realiza y registra el corte.").strip()

        if st.button(f"Crear corte con {len(pending)} documento(s)", type="primary"):
            if not cut_user:
                st.error("Indica el usuario o iniciales de la persona que realiza el corte.")
            else:
                def mutation(latest: bytes) -> bytes:
                    return create_cut(latest, requirement, cut_user, cut_date)
                try:
                    updated = store.mutate(DATA_FILE, mutation, f"Crear corte {requirement}")
                    st.session_state.operational_bytes = updated
                    st.success("Corte creado correctamente.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"No fue posible crear el corte: {exc}")
    else:
        st.info("No hay documentación pendiente de corte.")

else:
    cuts = list_cuts(operational, requirement)
    if not cuts:
        st.info("Este requerimiento todavía no tiene cortes.")
        st.stop()

    cut_map = {f"Corte {x['corte']} · {x['fecha']} · {x['documentos']} documentos": x["corte"] for x in cuts}
    selected_cut_labels = st.multiselect("Cortes a incluir", list(cut_map.keys()), placeholder="Seleccionar cortes")
    selected_cuts = [cut_map[x] for x in selected_cut_labels]

    if selected_cuts:
        records = documents_for_cuts(operational, requirement, selected_cuts)
        contracts_count = len({x["contrato"] for x in records})
        c1, c2 = st.columns(2)
        c1.metric("Contratos", contracts_count)
        c2.metric("Documentos", len(records))
        st.dataframe(records, use_container_width=True, hide_index=True)
        docx = build_request_docx(requirement, selected_cuts, records)
        filename = f"Solicitud_{requirement}_Cortes_{'-'.join(map(str, sorted(selected_cuts)))}.docx"
        st.download_button(
            "Generar y descargar Word",
            data=docx,
            file_name=filename,
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            type="primary",
        )
    else:
        st.caption("Selecciona uno o varios cortes. Puedes generar un oficio con un solo corte o consolidar varios.")
