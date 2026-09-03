from __future__ import annotations

from datetime import datetime
from hashlib import sha1

import streamlit as st

from app_state import AppState
from services.excel_service import (
    add_faltantes,
    create_cut,
    create_general_cut,
    cut_details,
    delete_cut,
    delete_general_cut,
    delete_pending_records,
    document_history,
    documents_for_cuts,
    general_cut_details,
    list_contracts,
    list_cuts,
    list_general_cuts,
    load_catalog,
    manual_visible_documents,
    move_records_to_cut,
    move_records_to_general_cut,
    pending_summary,
    pending_summary_all,
    remove_records_from_cut,
    remove_records_from_general_cut,
)
from services.word_service import build_request_docx


APP_PASSWORD = "1234"


class BaseView:
    def __init__(
        self,
        state: AppState,
        operational: bytes,
        catalog_bytes: bytes,
        requirement: str,
    ) -> None:
        self.state = state
        self.operational = operational
        self.catalog_bytes = catalog_bytes
        self.requirement = requirement

    def persist(self, mutator, message: str) -> bytes:
        return self.state.persist(mutator, message)


class CapturaFaltantesView(BaseView):
    def render(self) -> None:
        requirement = self.requirement
        operational = self.operational
        catalog_bytes = self.catalog_bytes

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
        st.caption(
            "Selecciona documentos del catálogo. Si el documento que necesitas no aparece en la lista, "
            "puedes escribirlo directamente en el campo de solicitud libre."
        )
        selected_labels = st.multiselect(
            "Documentos",
            list(options.keys()),
            placeholder="Seleccionar documentos",
        )
        selected_items = [dict(options[x]) for x in selected_labels]

        st.markdown("#### Especificación de la solicitud")
        st.caption(
            "Para documentos del catálogo, la especificación es opcional y se mostrará entre paréntesis. "
            "También puedes capturar directamente un documento que no exista en el catálogo."
        )

        for item in selected_items:
            item["especificacion"] = st.text_input(
                item["documento"],
                key=f"spec_{requirement}_{contract['contrato']}_{item['codigo']}",
                placeholder="Ej. Tarjetas de precios unitarios",
                help="Escribe únicamente la parte específica que deseas solicitar.",
            )

        free_document = st.text_area(
            "Documento no listado / solicitud libre",
            key=f"free_doc_{requirement}_{contract['contrato']}",
            placeholder="Escribe aquí el documento faltante cuando no exista en la lista del catálogo.",
            help="Este texto se guardará como una solicitud independiente y aparecerá tal como lo escribas.",
        ).strip()

        custom_item = None
        custom_pending = False
        if free_document:
            normalized = " ".join(free_document.casefold().split())
            custom_code = "LIBRE_" + sha1(normalized.encode("utf-8")).hexdigest()[:16]
            custom_item = {
                "codigo": custom_code,
                "documento": free_document,
                "especificacion": "",
            }
            custom_history = history.get(custom_code, {})
            custom_pending = bool(custom_history.get("pending"))
            if custom_pending:
                st.info(
                    "Este documento escrito manualmente ya se encuentra pendiente de corte para el contrato seleccionado."
                )
            else:
                selected_items.append(custom_item)

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
                        auditor_for_save, selected_items,
                        contract_row=contract["row"],
                    )

                try:
                    self.persist(mutation, f"Registrar faltantes {requirement} · {contract['contrato']}")
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
                        "Origen": item.get("origen", "Aplicación"),
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


class CortesView(BaseView):
    def _detail_data_version(self):
        return (
            self.state.session.get("operational_version")
            or hash(self.operational)
        )

    def _get_cached_cut_details(self, cache_key: str, loader):
        version = self._detail_data_version()
        cached = self.state.session.get(cache_key)
        if not cached or cached.get("version") != version:
            cached = {
                "version": version,
                "details": loader(),
            }
            self.state.session[cache_key] = cached
        return cached["details"]

    def _show_create_cut_dialog(self, cut_user: str, cut_date) -> None:
        requirement = self.requirement

        @st.dialog("Contraseña")
        def confirm_create_cut():
            password_key = f"_create_cut_password_{requirement}"
            password = st.text_input("Contraseña", type="password", key=password_key)
            if st.button("Confirmar", key=f"_confirm_create_cut_{requirement}"):
                if password != APP_PASSWORD:
                    return
                try:
                    self.persist(
                        lambda latest: create_cut(latest, requirement, cut_user, cut_date),
                        f"Crear corte {requirement}",
                    )
                    st.session_state.pop(password_key, None)
                    st.success("Corte creado correctamente.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"No fue posible crear el corte: {exc}")

        password_key = f"_create_cut_password_{requirement}"
        st.session_state.pop(password_key, None)
        confirm_create_cut()

    def _show_create_general_cut_dialog(self, cut_user: str, cut_date) -> None:
        @st.dialog("Contraseña")
        def confirm_create_general_cut():
            password_key = "_create_general_cut_password"
            password = st.text_input("Contraseña", type="password", key=password_key)
            if st.button("Confirmar", key="_confirm_create_general_cut"):
                if password != APP_PASSWORD:
                    return
                try:
                    self.persist(
                        lambda latest: create_general_cut(latest, cut_user, cut_date),
                        "Crear corte general",
                    )
                    st.session_state.pop(password_key, None)
                    st.success("Corte creado correctamente.")
                    st.rerun()
                except Exception as exc:
                    st.error(f"No fue posible crear el corte: {exc}")

        password_key = "_create_general_cut_password"
        st.session_state.pop(password_key, None)
        confirm_create_general_cut()

    def _render_general(self) -> None:
        operational = self.operational
        pending_all = pending_summary_all(operational)
        cuts = list_general_cuts(operational)

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
                    detail_shown_key = f"detail_shown_general_{cut['corte']}"
                    detail_cache_key = f"detail_cache_general_{cut['corte']}"

                    if not self.state.session.get(detail_shown_key, False):
                        if st.button(
                            "Ver detalle de este corte",
                            key=f"load_detail_general_{cut['corte']}",
                        ):
                            self.state.session[detail_shown_key] = True

                    if self.state.session.get(detail_shown_key, False):
                        details = self._get_cached_cut_details(
                            detail_cache_key,
                            lambda: general_cut_details(operational, cut["corte"]),
                        )
                        if details:
                            display_rows = [{
                                "Requerimiento": x["requerimiento"],
                                "Contrato": x["contrato"],
                                "Auditor": x["auditor"],
                                "Solicitud": x["solicitud"],
                                "Origen": x.get("origen", "Aplicación"),
                                "Documento catálogo": x["documento"],
                                "Especificación": x["especificacion"],
                            } for x in details]
                            st.dataframe(display_rows, use_container_width=True, hide_index=True)

                            detail_options = {
                                f"{x['requerimiento']} · {x['contrato']} · {x['solicitud']}": x["id"]
                                for x in details
                            }
                            to_remove = st.multiselect(
                                "Retirar documentos de este corte",
                                list(detail_options.keys()),
                                key=f"remove_general_cut_{cut['corte']}",
                                placeholder="Seleccionar documentos",
                            )
                            if st.button(
                                "Retirar seleccionados del corte",
                                key=f"remove_general_cut_btn_{cut['corte']}",
                                disabled=not to_remove,
                            ):
                                ids = [detail_options[x] for x in to_remove]
                                try:
                                    self.persist(
                                        lambda latest: remove_records_from_general_cut(
                                            latest, ids, cut["corte"]
                                        ),
                                        f"Retirar documentos Corte general {cut['corte']}",
                                    )
                                    st.success("Los documentos fueron retirados del corte y regresaron a documentación pendiente.")
                                    st.rerun()
                                except Exception as exc:
                                    st.error(f"No fue posible retirar los documentos: {exc}")
                        else:
                            st.info("Este corte no contiene documentos.")

                    if pending_all:
                        pending_options = {
                            f"{x['requerimiento']} · {x['contrato']} · {x['solicitud']} · {x['auditor']}": x["id"]
                            for x in pending_all
                        }
                        to_add = st.multiselect(
                            "Agregar documentación pendiente a este corte",
                            list(pending_options.keys()),
                            key=f"add_general_cut_{cut['corte']}",
                            placeholder="Seleccionar documentos pendientes",
                        )
                        if st.button(
                            "Agregar seleccionados al corte",
                            key=f"add_general_cut_btn_{cut['corte']}",
                            disabled=not to_add,
                        ):
                            ids = [pending_options[x] for x in to_add]
                            try:
                                self.persist(
                                    lambda latest: move_records_to_general_cut(
                                        latest, ids, cut["corte"]
                                    ),
                                    f"Agregar documentos Corte general {cut['corte']}",
                                )
                                st.success("Los documentos fueron agregados al corte.")
                                st.rerun()
                            except Exception as exc:
                                st.error(f"No fue posible agregar los documentos: {exc}")

                    st.divider()
                    confirm_delete = st.checkbox(
                        f"Confirmo que deseo eliminar el Corte {cut['corte']}. Sus documentos regresarán a pendientes.",
                        key=f"confirm_delete_general_cut_{cut['corte']}",
                    )
                    if st.button(
                        f"Eliminar Corte {cut['corte']}",
                        key=f"delete_general_cut_{cut['corte']}",
                        disabled=not confirm_delete,
                    ):
                        try:
                            self.persist(
                                lambda latest: delete_general_cut(latest, cut["corte"]),
                                f"Eliminar Corte general {cut['corte']}",
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

            header = st.columns([1.35, 1.15, 0.85, 4.4, 1.65, 1.45, 1.25])
            for col, label in zip(
                header,
                ["Requerimiento", "Contrato", "Auditor", "Solicitud", "Origen", "Fecha", "Acción"],
            ):
                col.markdown(f"**{label}**")

            for idx, row in enumerate(pending):
                cols = st.columns([1.35, 1.15, 0.85, 4.4, 1.65, 1.45, 1.25])
                cols[0].write(row.get("requerimiento", ""))
                cols[1].write(row["contrato"])
                cols[2].write(row["auditor"])
                cols[3].write(row["solicitud"])
                cols[4].write(row.get("origen", "Aplicación"))

                fecha = row.get("fecha")
                if hasattr(fecha, "strftime"):
                    fecha_text = fecha.strftime("%d/%m/%Y %H:%M")
                else:
                    fecha_text = "—"
                cols[5].write(fecha_text)

                if cols[6].button(
                    "Eliminar",
                    key=f"delete_pending_all_{idx}_{row['id']}",
                    use_container_width=True,
                ):
                    st.session_state["pending_delete_confirm_all"] = row["id"]

                if st.session_state.get("pending_delete_confirm_all") == row["id"]:
                    st.warning(
                        f"¿Seguro que deseas eliminar este registro pendiente? "
                        f"{row['contrato']} · {row['solicitud']}"
                    )
                    c_yes, c_no = st.columns([1, 1])
                    if c_yes.button(
                        "Sí, eliminar",
                        key=f"confirm_delete_pending_all_{idx}_{row['id']}",
                        type="primary",
                    ):
                        try:
                            row_requirement = row.get("requerimiento")
                            self.persist(
                                lambda latest: delete_pending_records(
                                    latest, row_requirement, [row["id"]]
                                ),
                                f"Eliminar faltante pendiente · {row_requirement}",
                            )
                            st.session_state.pop("pending_delete_confirm_all", None)
                            st.success("El registro pendiente fue eliminado.")
                            st.rerun()
                        except Exception as exc:
                            st.error(f"No fue posible eliminar el registro: {exc}")

                    if c_no.button(
                        "Cancelar",
                        key=f"cancel_delete_pending_all_{idx}_{row['id']}",
                    ):
                        st.session_state.pop("pending_delete_confirm_all", None)
                        st.rerun()

            st.divider()
            c1, c2 = st.columns(2)
            with c1:
                cut_date = st.date_input(
                    "Fecha del nuevo corte",
                    value=datetime.now(self.state.timezone).date(),
                    key="general_cut_date",
                )
            with c2:
                cut_user = st.text_input(
                    "Usuario / iniciales",
                    help="Persona que realiza y registra el corte.",
                    key="general_cut_user",
                ).strip()

            if st.button(
                f"Crear corte con {len(pending_all)} documento(s)",
                type="primary",
                key="create_general_cut_button",
            ):
                if not cut_user:
                    st.error("Indica el usuario o iniciales de la persona que realiza el corte.")
                else:
                    self._show_create_general_cut_dialog(cut_user, cut_date)
        else:
            st.info("No hay documentación pendiente de corte.")

    def render(self) -> None:
        requirement = self.requirement
        operational = self.operational

        if requirement == "Todos":
            self._render_general()
            return

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
                    detail_shown_key = f"detail_shown_{requirement}_{cut['corte']}"
                    detail_cache_key = f"detail_cache_{requirement}_{cut['corte']}"

                    if not self.state.session.get(detail_shown_key, False):
                        if st.button(
                            "Ver detalle de este corte",
                            key=f"load_detail_{requirement}_{cut['corte']}",
                        ):
                            self.state.session[detail_shown_key] = True

                    if self.state.session.get(detail_shown_key, False):
                        details = self._get_cached_cut_details(
                            detail_cache_key,
                            lambda: cut_details(operational, requirement, cut["corte"]),
                        )
                        if details:
                            display_rows = [{
                                "Requerimiento": x["requerimiento"],
                                "Contrato": x["contrato"],
                                "Auditor": x["auditor"],
                                "Solicitud": x["solicitud"],
                                "Origen": x.get("origen", "Aplicación"),
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
                                    self.persist(
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
                                self.persist(
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
                            self.persist(
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

            header = st.columns([1.35, 1.15, 0.85, 4.4, 1.65, 1.45, 1.25])
            for col, label in zip(
                header,
                ["Requerimiento", "Contrato", "Auditor", "Solicitud", "Origen", "Fecha", "Acción"],
            ):
                col.markdown(f"**{label}**")

            for idx, row in enumerate(pending):
                cols = st.columns([1.35, 1.15, 0.85, 4.4, 1.65, 1.45, 1.25])
                cols[0].write(row.get("requerimiento", requirement))
                cols[1].write(row["contrato"])
                cols[2].write(row["auditor"])
                cols[3].write(row["solicitud"])
                cols[4].write(row.get("origen", "Aplicación"))

                fecha = row.get("fecha")
                if hasattr(fecha, "strftime"):
                    fecha_text = fecha.strftime("%d/%m/%Y %H:%M")
                else:
                    fecha_text = "—"
                cols[5].write(fecha_text)

                if cols[6].button(
                    "Eliminar",
                    key=f"delete_pending_{idx}_{row['id']}",
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
                        key=f"confirm_delete_pending_{idx}_{row['id']}",
                        type="primary",
                    ):
                        try:
                            self.persist(
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
                        key=f"cancel_delete_pending_{idx}_{row['id']}",
                    ):
                        st.session_state.pop("pending_delete_confirm", None)
                        st.rerun()

            st.divider()
            c1, c2 = st.columns(2)
            with c1:
                cut_date = st.date_input("Fecha del nuevo corte", value=datetime.now(self.state.timezone).date())
            with c2:
                cut_user = st.text_input(
                    "Usuario / iniciales",
                    help="Persona que realiza y registra el corte.",
                ).strip()

            if st.button(f"Crear corte con {len(pending_all)} documento(s)", type="primary"):
                if not cut_user:
                    st.error("Indica el usuario o iniciales de la persona que realiza el corte.")
                else:
                    self._show_create_cut_dialog(cut_user, cut_date)
        else:
            st.info("No hay documentación pendiente de corte.")


class GenerarOficioView(BaseView):
    @staticmethod
    def _authorize_once() -> None:
        password = st.session_state.get("_generar_oficio_password", "")
        authorized = password == APP_PASSWORD
        st.session_state["_generar_oficio_authorized"] = authorized
        st.session_state["_generar_oficio_password_error"] = not authorized
        st.session_state["_generar_oficio_password"] = ""

    def render(self) -> None:
        authorized = bool(
            st.session_state.get("_generar_oficio_authorized", False)
        )
        if not authorized:
            st.text_input(
                "Contraseña",
                type="password",
                key="_generar_oficio_password",
                on_change=self._authorize_once,
            )
            if st.session_state.get("_generar_oficio_password_error", False):
                st.error("Contraseña incorrecta.")
            return

        requirement = self.requirement
        operational = self.operational

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
            preview_records = [{
                "Requerimiento": x.get("requerimiento", requirement),
                "Contrato": x["contrato"],
                "Documento": x["documento"],
                "Corte": x["corte"],
            } for x in records]
            st.dataframe(preview_records, use_container_width=True, hide_index=True)
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
