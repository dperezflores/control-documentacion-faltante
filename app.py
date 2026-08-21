from __future__ import annotations

from datetime import date

import streamlit as st

from services.excel_service import (
    add_faltantes,
    create_cut,
    documents_for_cuts,
    get_faltantes,
    list_contracts,
    list_cuts,
    list_requirements,
    load_catalog,
    pending_summary,
)
from services.github_service import build_store
from services.word_service import build_request_docx

st.set_page_config(page_title="Control documental", page_icon="📋", layout="wide")

DATA_FILE = "data/Documentacion_faltante.xlsx"
CATALOG_FILE = "data/Codificacion_documentos.xlsx"


def secret_dict():
    try:
        return dict(st.secrets)
    except Exception:
        return {}


store, storage_mode = build_store(secret_dict())

st.title("Control de documentación faltante")
st.caption(f"Prototipo 0.1 · almacenamiento: {storage_mode}")

# Inicialización: permite cargar los dos Excel originales una sola vez desde la interfaz.
missing = []
for path in (DATA_FILE, CATALOG_FILE):
    try:
        store.read(path)
    except FileNotFoundError:
        missing.append(path)
    except Exception as exc:
        st.error(f"No fue posible consultar {path}: {exc}")
        st.stop()

if missing:
    st.subheader("Inicializar archivos del prototipo")
    st.info("Los archivos Excel todavía no están guardados en el almacenamiento del prototipo. Cárguelos una sola vez; la aplicación los conservará y trabajará sobre ellos.")
    op_upload = st.file_uploader("Excel de documentación faltante", type=["xlsx"], key="op")
    cat_upload = st.file_uploader("Excel de codificación de documentos", type=["xlsx"], key="cat")
    if st.button("Inicializar prototipo", type="primary", disabled=not (op_upload and cat_upload)):
        try:
            if DATA_FILE in missing:
                store.write_new(DATA_FILE, op_upload.getvalue(), "Inicializar Excel operativo")
            if CATALOG_FILE in missing:
                store.write_new(CATALOG_FILE, cat_upload.getvalue(), "Inicializar catálogo de codificación")
            st.success("Archivos inicializados correctamente.")
            st.rerun()
        except Exception as exc:
            st.error(f"No fue posible inicializar los archivos: {exc}")
    st.stop()

section = st.sidebar.radio("Menú", ["📋 Capturar faltantes", "✂️ Cortes", "📄 Generar oficio"])
user_name = st.sidebar.text_input("Usuario / iniciales", value="")

try:
    operational = store.read(DATA_FILE).content
    catalog_bytes = store.read(CATALOG_FILE).content
except Exception as exc:
    st.error(f"No fue posible cargar los archivos de datos: {exc}")
    st.stop()

requirements = list_requirements(operational)
if not requirements:
    st.warning("No se encontraron hojas de requerimiento.")
    st.stop()

requirement = st.selectbox("Requerimiento", requirements)

if section == "📋 Capturar faltantes":
    contracts = list_contracts(operational, requirement)
    if not contracts:
        st.info("No hay contratos disponibles en este requerimiento.")
        st.stop()

    labels = {f"{c['contrato']} · {str(c['obra'])[:70]}": c for c in contracts}
    selected_label = st.selectbox("Contrato", list(labels.keys()))
    contract = labels[selected_label]

    c1, c2 = st.columns(2)
    with c1:
        st.markdown("**Obra**")
        st.write(contract["obra"])
    with c2:
        st.markdown("**Contratista**")
        st.write(contract["contratista"])
        st.markdown("**Auditor registrado en Excel**")
        st.write(contract["auditor"] or "—")

    procedure = st.selectbox(
        "Procedimiento de adjudicación",
        ["DIR", "LPU", "LSI"],
        format_func=lambda x: {"DIR": "DIR · Adjudicación Directa", "LPU": "LPU · Licitación Pública", "LSI": "LSI · Licitación Simplificada"}[x],
    )
    catalog = load_catalog(catalog_bytes, procedure)
    query = st.text_input("Buscar documento", placeholder="Ej. contrato, fianza, bitácora, laboratorio...")
    if query.strip():
        q = query.casefold()
        filtered = [x for x in catalog if q in x["documento"].casefold() or q in x["codigo"].casefold()]
    else:
        filtered = catalog

    existing = get_faltantes(operational, requirement, contract["contrato"])
    existing_codes = {str(x.get("Codigo_documento")) for x in existing}
    available = [x for x in filtered if x["codigo"] not in existing_codes]

    st.subheader("Documentos disponibles")
    options = {f"{x['documento']}  ·  {x['codigo']}": x for x in available}
    selected_labels = st.multiselect("Seleccione uno o varios documentos faltantes", list(options.keys()))

    if existing:
        with st.expander(f"Ya registrados ({len(existing)})"):
            for item in existing:
                st.write(f"• {item.get('Documento')} · Corte: {item.get('Corte')}")

    if st.button("Guardar documentación faltante", type="primary", disabled=not selected_labels):
        if not user_name.strip() and not contract["auditor"]:
            st.error("Indique el usuario/iniciales o asegúrese de que el contrato tenga auditor en el Excel.")
        else:
            items = [options[x] for x in selected_labels]
            auditor = user_name.strip() or str(contract["auditor"])

            def mutation(latest: bytes) -> bytes:
                return add_faltantes(latest, requirement, contract["contrato"], procedure, auditor, items)

            try:
                store.mutate(DATA_FILE, mutation, f"Registrar faltantes {requirement} · {contract['contrato']}")
                st.success(f"Se registraron {len(items)} documento(s) faltante(s).")
                st.rerun()
            except Exception as exc:
                st.error(f"No fue posible guardar: {exc}")

elif section == "✂️ Cortes":
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
        cut_date = st.date_input("Fecha del nuevo corte", value=date.today())
        if st.button(f"Crear corte con {len(pending)} documento(s)", type="primary"):
            if not user_name.strip():
                st.error("Indique el usuario/iniciales en el menú lateral.")
            else:
                def mutation(latest: bytes) -> bytes:
                    return create_cut(latest, requirement, user_name.strip(), cut_date)
                try:
                    store.mutate(DATA_FILE, mutation, f"Crear corte {requirement}")
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

    cut_map = {f"Corte {x['corte']} · {x['fecha']} · {x['documentos']} docs": x["corte"] for x in cuts}
    selected_cut_labels = st.multiselect("Cortes a incluir", list(cut_map.keys()))
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
        st.caption("Seleccione uno o varios cortes. Puede generar un oficio con un solo corte o consolidar varios.")
