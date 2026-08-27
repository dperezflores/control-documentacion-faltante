from __future__ import annotations

from datetime import datetime
from zoneinfo import ZoneInfo

import streamlit as st

from services.excel_service import (
    add_faltantes,
    create_cut,
    cut_details,
    delete_cut,
    delete_pending_records,
    document_history,
    documents_for_cuts,
    list_contracts,
    list_cuts,
    list_requirements,
    load_catalog,
    manual_visible_documents,
    move_records_to_cut,
    pending_summary,
    remove_records_from_cut,
)
from services.github_service import build_store
from services.word_service import build_request_docx

st.set_page_config(page_title="Control documental", layout="wide")

DATA_FILE = "data/Documentacion_faltante.xlsx"
CATALOG_FILE = "data/Codificacion_documentos.xlsx"
LOCAL_TZ = ZoneInfo("America/Mexico_City")

CUSTOM_CSS = """
<style>
    .stApp { background: #f6f7f9; }
    .block-container { max-width: 1220px; padding-top: 2rem; padding-bottom: 3rem; }
    h1 { color: #1f2937; font-size: 2rem !important; letter-spacing: -0.025em; margin-bottom: .25rem !important; }
    h2, h3 { color: #27313d; letter-spacing: -0.015em; }
    [data-testid="stSidebar"] { background: #ffffff; border-right: 1px solid #e5e7eb; }
    [data-testid="stSidebar"] .stRadio > label { font-weight: 700; }
    div[data-testid="stMetric"] {
        background: #ffffff; border: 1px solid #e5e7eb; border-radius: 12px; padding: 14px 16px;
        box-shadow: 0 1px 2px rgba(0,0,0,.025);
    }
    div[data-testid="stExpander"] {
        background: #ffffff; border: 1px solid #e5e7eb; border-radius: 12px;
    }
    .info-card {
        background: #ffffff; border: 1px solid #e5e7eb; border-radius: 12px;
        padding: 18px 20px; margin: 8px 0 18px 0;
        box-shadow: 0 1px 2px rgba(0,0,0,.025);
    }
    .info-label { color: #6b7280; font-size: .76rem; font-weight: 700; text-transform: uppercase; letter-spacing: .05em; }
    .info-value { color: #25282c; font-size: .98rem; margin: 3px 0 10px 0; }
    .section-kicker { color: #6b7280; font-size: .88rem; margin-bottom: 1rem; }
    .subtle-box {
        background: #ffffff; border: 1px solid #e5e7eb; border-radius: 10px; padding: 14px 16px; margin-bottom: 12px;
    }
    .stButton > button { border-radius: 8px; font-weight: 650; white-space: nowrap; }
    .stButton > button[kind="primary"] { font-weight: 700; }
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


def persist(mutator, message: str) -> bytes:
    updated = store.mutate(DATA_FILE, mutator, message)
    st.session_state.operational_bytes = updated
    return updated


st.title("Control de documentación faltante")
st.markdown(
    '<div class="section-kicker">Registro homologado de faltantes, administración de cortes y preparación de solicitudes.</div>',
    unsafe_allow_html=True,
)

loaded = load_session_files()
if loaded is None:
    if storage_mode == "Google Drive":
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
        load_session_files(force=True)
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
    filtered_contracts = contracts if auditor_filter == "Todos" else [
        c for c in contracts if str(c["auditor"]).strip() == auditor_filter
    ]

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
    available = [x for x in catalog if not history.get(x["codigo"], {}).get("pending", False)]

    options = {}
    for item in available:
        prior_cuts = history.get(item["codigo"], {}).get("cuts", [])
        label = item["documento"]
        if prior_cuts:
            label += f"  ·  Ya solicitado en corte(s): {', '.join(map(str, prior_cuts))}"
        options[label] = item

    st.subheader("Documentación faltante")
    st.caption("Selecciona los documentos faltantes. Si necesitas pedir sólo una parte del documento catalogado, utiliza el campo de especificación.")
    selected_labels = st.multiselect("Documentos", list(options.keys()), placeholder="Seleccionar documentos")
    selected_items = [dict(options[x]) for x in selected_labels]

    if selected_items:
        st.markdown("#### Especificación de la solicitud")
        st.caption("Opcional. Si se deja vacío, se solicitará el nombre completo del documento del catálogo.")
        for item in selected_items:
            item["especificacion"] = st.text_input(
                item["documento"],
                key=f"spec_{requirement}_{contract['contrato']}_{item['codigo']}",
                placeholder="Ej. Tarjetas de precios unitarios",
                help="Escribe únicamente lo que deseas solicitar cuando el concepto del catálogo sea más amplio.",
            )

    repeated = []
    for item in selected_items:
        cuts = history.get(item["codigo"], {}).get("cuts", [])
        if cuts:
            repeated.append((item, cuts))

    repeat_confirmed = True
    if repeated:
        names = "; ".join(
            f"{item['documento']} (corte(s) {', '.join(map(str, cuts))})"
            for item, cuts in repeated
        )
        st.warning(
            f"La selección incluye documentación ya solicitada anteriormente: {names}. "
            "Si no fue entregada, puedes volver a solicitarla en el siguiente corte."
        )
        repeat_confirmed = st.checkbox("Confirmo que deseo volver a agregar esta documentación para una nueva solicitud.")

    auditor_for_save = str(contract["auditor"] or "").strip()
    if not auditor_for_save:
        auditor_for_save = st.text_input(
            "Usuario / iniciales",
            help="Este contrato no tiene auditor asignado en el Excel. Indica quién registra los faltantes.",
        ).strip()

    if st.button(
        "Guardar documentación faltante",
        type="primary",
        disabled=not selected_items or not repeat_confirmed,
    ):
        if not auditor_for_save:
            st.error("Este contrato no tiene auditor asignado. Indica el usuario o iniciales para continuar.")
        else:
            def mutation(latest: bytes) -> bytes:
                return add_faltantes(
                    latest, requirement, contract["contrato"], procedure,
                    auditor_for_save, selected_items
                )

            try:
                persist(mutation, f"Registrar faltantes {requirement} · {contract['contrato']}")
                st.success(f"Se registraron {len(selected_items)} documento(s) para el siguiente corte.")
                st.rerun()
            except Exception as exc:
                st.error(f"No fue posible guardar: {exc}")

    manual_docs = manual_visible_documents(
        operational, requirement, contract["contrato"]
    )
    existing_count = (
        sum(1 for item in history.values() if item.get("cuts") or item.get("pending"))
        + len(manual_docs)
    )
    if existing_count:
        with st.expander(f"Historial del contrato ({existing_count} documentos)"):
            rows = []
            catalog_by_code = {x["codigo"]: x["documento"] for x in catalog}
            for code, item in history.items():
                status = "Pendiente de corte" if item.get("pending") else "Sin solicitud pendiente"
                cuts = ", ".join(map(str, item.get("cuts", []))) or "—"
                rows.append({
                    "Documento": item.get("documento") or catalog_by_code.get(code, code),
                    "Origen": "Aplicación",
                    "Cortes anteriores": cuts,
                    "Estado actual": status,
                })

            for text in manual_docs:
                rows.append({
                    "Documento": text,
                    "Origen": "Registro previo / manual",
                    "Cortes anteriores": "—",
                    "Estado actual": "No registrado con la aplicación",
                })

            st.dataframe(rows, use_container_width=True, hide_index=True)
            if manual_docs:
                st.caption(
                    "Los registros marcados como 'Registro previo / manual' ya existían "
                    "en la columna Documentación faltante del Excel y no fueron creados "
                    "por la aplicación."
                )

elif section == "Cortes":
    pending_all = pending_summary(operational, requirement)
    cuts = list_cuts(operational, requirement)

    m1, m2 = st.columns(2)
    m1.metric("Pendientes de corte", len(pending_all))
    m2.metric("Cortes existentes", len(cuts))

    if cuts:
        st.subheader("Cortes existentes")
        st.caption("Consulta el detalle de cada corte, agrega documentos pendientes, retira documentos o elimina un corte completo.")

        for cut in cuts:
            with st.expander(
                f"Corte {cut['corte']} · {cut['fecha']} · {cut['documentos']} documento(s) · {cut['creado_por']}"
            ):
                details = cut_details(operational, requirement, cut["corte"])
                if details:
                    display_rows = [{
                        "Contrato": x["contrato"],
                        "Auditor": x["auditor"],
                        "Solicitud": x["solicitud"],
                        "Documento catálogo": x["documento"],
                        "Especificación": x["especificacion"],
                    } for x in details]
                    st.dataframe(display_rows, use_container_width=True, hide_index=True)

                    detail_options = {
                        f"{x['contrato']} · {x['solicitud']}": x["id"] for x in details
                    }
                    to_remove = st.multiselect(
                        "Retirar documentos de este corte",
                        list(detail_options.keys()),
                        key=f"remove_cut_{cut['corte']}",
                        placeholder="Seleccionar documentos",
                    )
                    if st.button(
                        "Retirar seleccionados del corte",
                        key=f"remove_cut_btn_{cut['corte']}",
                        disabled=not to_remove,
                    ):
                        ids = [detail_options[x] for x in to_remove]
                        try:
                            persist(
                                lambda latest: remove_records_from_cut(
                                    latest, requirement, ids, cut["corte"]
                                ),
                                f"Retirar documentos Corte {cut['corte']} · {requirement}",
                            )
                            st.success("Los documentos fueron retirados del corte y regresaron a documentación pendiente.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"No fue posible retirar los documentos: {exc}")
                else:
                    st.info("Este corte no contiene documentos.")

                if pending_all:
                    pending_options = {
                        f"{x['contrato']} · {x['solicitud']} · {x['auditor']}": x["id"]
                        for x in pending_all
                    }
                    to_add = st.multiselect(
                        "Agregar documentación pendiente a este corte",
                        list(pending_options.keys()),
                        key=f"add_cut_{cut['corte']}",
                        placeholder="Seleccionar documentos pendientes",
                    )
                    if st.button(
                        "Agregar seleccionados al corte",
                        key=f"add_cut_btn_{cut['corte']}",
                        disabled=not to_add,
                    ):
                        ids = [pending_options[x] for x in to_add]
                        try:
                            persist(
                                lambda latest: move_records_to_cut(
                                    latest, requirement, ids, cut["corte"]
                                ),
                                f"Agregar documentos Corte {cut['corte']} · {requirement}",
                            )
                            st.success("Los documentos fueron agregados al corte.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"No fue posible agregar los documentos: {exc}")

                st.divider()
                confirm_delete = st.checkbox(
                    f"Confirmo que deseo eliminar el Corte {cut['corte']}. Sus documentos regresarán a pendientes.",
                    key=f"confirm_delete_cut_{cut['corte']}",
                )
                if st.button(
                    f"Eliminar Corte {cut['corte']}",
                    key=f"delete_cut_{cut['corte']}",
                    disabled=not confirm_delete,
                ):
                    try:
                        persist(
                            lambda latest: delete_cut(latest, requirement, cut["corte"]),
                            f"Eliminar Corte {cut['corte']} · {requirement}",
                        )
                        st.success(f"El Corte {cut['corte']} fue eliminado. Sus documentos regresaron a pendientes.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"No fue posible eliminar el corte: {exc}")
    else:
        st.info("Todavía no existen cortes para este requerimiento.")

    st.subheader("Documentación pendiente")
    if pending_all:
        current_contract = st.session_state.get("pending_contract_filter", "Todos")
        current_auditor = st.session_state.get("pending_auditor_filter", "Todos")

        # Opciones de auditor limitadas por el contrato actualmente seleccionado.
        auditor_source = [
            x for x in pending_all
            if current_contract == "Todos" or str(x["contrato"]) == current_contract
        ]
        auditors_pending = sorted({
            str(x["auditor"]).strip()
            for x in auditor_source
            if str(x["auditor"]).strip()
        })
        auditor_options = ["Todos"] + auditors_pending
        if current_auditor not in auditor_options:
            current_auditor = "Todos"
            st.session_state["pending_auditor_filter"] = "Todos"

        # Opciones de contrato limitadas por el auditor actualmente seleccionado.
        contract_source = [
            x for x in pending_all
            if current_auditor == "Todos" or str(x["auditor"]).strip() == current_auditor
        ]
        contracts_pending = sorted({str(x["contrato"]) for x in contract_source})
        contract_options = ["Todos"] + contracts_pending
        if current_contract not in contract_options:
            current_contract = "Todos"
            st.session_state["pending_contract_filter"] = "Todos"

        p1, p2 = st.columns(2)
        with p1:
            pending_contract_filter = st.selectbox(
                "Filtrar por contrato",
                contract_options,
                key="pending_contract_filter",
            )
        with p2:
            pending_auditor_filter = st.selectbox(
                "Filtrar por auditor",
                auditor_options,
                key="pending_auditor_filter",
            )

        pending = [
            x for x in pending_all
            if (
                pending_contract_filter == "Todos"
                or str(x["contrato"]) == pending_contract_filter
            )
            and (
                pending_auditor_filter == "Todos"
                or str(x["auditor"]).strip() == pending_auditor_filter
            )
        ]

        header = st.columns([1.15, 0.85, 4.9, 1.8, 1.55, 1.35])
        for col, label in zip(
            header,
            ["Contrato", "Auditor", "Solicitud", "Origen", "Fecha", "Acción"],
        ):
            col.markdown(f"**{label}**")

        for row in pending:
            cols = st.columns([1.15, 0.85, 4.9, 1.8, 1.55, 1.35])
            cols[0].write(row["contrato"])
            cols[1].write(row["auditor"])
            cols[2].write(row["solicitud"])
            cols[3].write(row.get("origen", "Aplicación"))

            fecha = row.get("fecha")
            if hasattr(fecha, "strftime"):
                fecha_text = fecha.strftime("%d/%m/%Y %H:%M")
            else:
                fecha_text = "—"
            cols[4].write(fecha_text)

            if cols[5].button(
                "Eliminar",
                key=f"delete_pending_{row['id']}",
                use_container_width=True,
            ):
                st.session_state["pending_delete_confirm"] = row["id"]

            if st.session_state.get("pending_delete_confirm") == row["id"]:
                st.warning(
                    f"¿Seguro que deseas eliminar este registro pendiente? "
                    f"{row['contrato']} · {row['solicitud']}"
                )
                c_yes, c_no = st.columns([1, 1])
                if c_yes.button(
                    "Sí, eliminar",
                    key=f"confirm_delete_pending_{row['id']}",
                    type="primary",
                ):
                    try:
                        persist(
                            lambda latest: delete_pending_records(
                                latest, requirement, [row["id"]]
                            ),
                            f"Eliminar faltante pendiente · {requirement}",
                        )
                        st.session_state.pop("pending_delete_confirm", None)
                        st.success("El registro pendiente fue eliminado.")
                        st.rerun()
                    except Exception as exc:
                        st.error(f"No fue posible eliminar el registro: {exc}")

                if c_no.button(
                    "Cancelar",
                    key=f"cancel_delete_pending_{row['id']}",
                ):
                    st.session_state.pop("pending_delete_confirm", None)
                    st.rerun()

        st.divider()
        c1, c2 = st.columns(2)
        with c1:
            cut_date = st.date_input("Fecha del nuevo corte", value=datetime.now(LOCAL_TZ).date())
        with c2:
            cut_user = st.text_input(
                "Usuario / iniciales",
                help="Persona que realiza y registra el corte.",
            ).strip()

        if st.button(f"Crear corte con {len(pending_all)} documento(s)", type="primary"):
            if not cut_user:
                st.error("Indica el usuario o iniciales de la persona que realiza el corte.")
            else:
                try:
                    persist(
                        lambda latest: create_cut(latest, requirement, cut_user, cut_date),
                        f"Crear corte {requirement}",
                    )
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

    cut_map = {
        f"Corte {x['corte']} · {x['fecha']} · {x['documentos']} documentos": x["corte"]
        for x in cuts
    }
    selected_cut_labels = st.multiselect(
        "Cortes a incluir",
        list(cut_map.keys()),
        placeholder="Seleccionar cortes",
    )
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
