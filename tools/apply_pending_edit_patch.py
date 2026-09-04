from pathlib import Path


def replace_once(path: str, old: str, new: str) -> None:
    p = Path(path)
    text = p.read_text(encoding="utf-8")
    if new in text:
        return
    if old not in text:
        raise RuntimeError(f"No se encontró el bloque esperado en {path}")
    p.write_text(text.replace(old, new, 1), encoding="utf-8")


# 1) Servicio: nueva función de actualización de pendientes.
service_path = Path("services/excel_service.py")
service = service_path.read_text(encoding="utf-8")
if "def update_pending_record(" not in service:
    anchor = "\n\ndef move_records_to_cut(data: bytes, requirement: str, record_ids: Iterable[str], cut_number: int) -> bytes:\n"
    function = r'''


def update_pending_record(
    data: bytes,
    requirement: str,
    record_id: str,
    documento: str,
    auditor: str,
) -> bytes:
    wb = _load(data)
    _ensure_technical_sheets(wb)
    tech = wb[FALTANTES_SHEET]
    h = _headers(tech)
    normalized_document = _normalize_single_line_text(documento)

    # Primero buscar un pendiente técnico existente.
    for r in range(2, tech.max_row + 1):
        if (
            tech.cell(r, h["Requerimiento"]).value == requirement
            and str(tech.cell(r, h["ID"]).value) == str(record_id)
            and str(tech.cell(r, h["Corte"]).value) == "PENDIENTE"
        ):
            contract = str(tech.cell(r, h["Contrato"]).value or "").strip()
            old_effective = _effective_document(
                tech.cell(r, h["Documento"]).value,
                tech.cell(r, h["Especificacion"]).value if "Especificacion" in h else "",
            )

            tech.cell(r, h["Documento"]).value = normalized_document
            if "Especificacion" in h:
                tech.cell(r, h["Especificacion"]).value = ""
            tech.cell(r, h["Auditor"]).value = auditor

            # Retirar de la celda visible el texto técnico anterior antes de
            # reconstruirla; de otro modo podría interpretarse como manual.
            visible_ws = wb[requirement]
            for visible_row in range(8, visible_ws.max_row + 1):
                if str(visible_ws.cell(visible_row, 2).value or "").strip() != contract:
                    continue
                visible_docs = _split_visible_documents(
                    visible_ws.cell(visible_row, 5).value or ""
                )
                visible_ws.cell(visible_row, 5).value = "\n".join(
                    text for text in visible_docs if text != old_effective
                )
                break

            _rebuild_visible_cell(wb, requirement, contract)
            return _save(wb)

    # Si no es técnico, puede ser un registro previo/manual que sólo existe
    # en la celda visible. Al editarlo se materializa como registro técnico.
    manual_rows = _manual_pending_rows_from_wb(wb, requirement)
    manual = next(
        (row for row in manual_rows if str(row["id"]) == str(record_id)),
        None,
    )
    if manual is None:
        raise ValueError(
            "No se encontró el registro a editar; es posible que ya haya cambiado."
        )

    contract = manual["contrato"]
    remaining_manual_docs = [
        row["solicitud"]
        for row in manual_rows
        if row["contrato"] == contract and str(row["id"]) != str(record_id)
    ]

    row = [None] * tech.max_column
    values = {
        "ID": uuid4().hex,
        "Requerimiento": requirement,
        "Contrato": contract,
        "Procedimiento": "",
        "Codigo_documento": "",
        "Documento": normalized_document,
        "Especificacion": "",
        "Fecha_deteccion": datetime.now(LOCAL_TZ).replace(
            tzinfo=None, microsecond=0
        ),
        "Auditor": auditor,
        "Corte": "PENDIENTE",
        "Origen": "Aplicación",
    }
    for name, value in values.items():
        row[h[name] - 1] = value
    tech.append(row)

    _write_visible_cell(
        wb,
        requirement,
        contract,
        remaining_manual_docs,
    )
    return _save(wb)
'''
    if anchor not in service:
        raise RuntimeError("No se encontró el punto de inserción de update_pending_record")
    service_path.write_text(service.replace(anchor, function + anchor, 1), encoding="utf-8")


# 2) Vista: importar la nueva función.
replace_once(
    "views.py",
    "    delete_pending_records,\n    document_history,\n",
    "    delete_pending_records,\n    document_history,\n    update_pending_record,\n",
)

# 3) Vista: diálogo reutilizable para ambos listados de pendientes.
replace_once(
    "views.py",
    '''        return cached["details"]\n\n    def _show_create_cut_dialog(self, cut_user: str, cut_date) -> None:\n''',
    '''        return cached["details"]\n\n    def _show_edit_pending_dialog(\n        self,\n        row: dict,\n        requirement: str,\n        key_prefix: str,\n    ) -> None:\n        record_id = str(row["id"])\n        document_key = f"_edit_pending_document_{key_prefix}_{record_id}"\n        auditor_key = f"_edit_pending_auditor_{key_prefix}_{record_id}"\n\n        @st.dialog("Editar registro pendiente")\n        def edit_pending():\n            new_document = st.text_area(\n                "Solicitud",\n                value=str(row.get("solicitud") or ""),\n                key=document_key,\n            ).strip()\n            new_auditor = st.text_input(\n                "Auditor",\n                value=str(row.get("auditor") or ""),\n                key=auditor_key,\n            ).strip()\n\n            c_save, c_cancel = st.columns(2)\n            if c_save.button(\n                "Guardar cambios",\n                key=f"_save_edit_pending_{key_prefix}_{record_id}",\n                type="primary",\n            ):\n                if not new_document or not new_auditor:\n                    st.error("La solicitud y el auditor son obligatorios.")\n                    return\n                try:\n                    self.persist(\n                        lambda latest: update_pending_record(\n                            latest,\n                            requirement,\n                            row["id"],\n                            new_document,\n                            new_auditor,\n                        ),\n                        f"Editar faltante pendiente · {requirement}",\n                    )\n                    st.session_state.pop(document_key, None)\n                    st.session_state.pop(auditor_key, None)\n                    st.success("El registro pendiente fue actualizado.")\n                    st.rerun()\n                except Exception as exc:\n                    st.error(f"No fue posible actualizar el registro: {exc}")\n\n            if c_cancel.button(\n                "Cancelar",\n                key=f"_cancel_edit_pending_{key_prefix}_{record_id}",\n            ):\n                st.session_state.pop(document_key, None)\n                st.session_state.pop(auditor_key, None)\n                st.rerun()\n\n        # Si se vuelve a abrir el mismo registro después de cerrar el diálogo,\n        # se precargan nuevamente los valores actuales de la fila.\n        st.session_state.pop(document_key, None)\n        st.session_state.pop(auditor_key, None)\n        edit_pending()\n\n    def _show_create_cut_dialog(self, cut_user: str, cut_date) -> None:\n''',
)

# 4) Vista general "Todos": botones Editar + Eliminar en la misma celda.
replace_once(
    "views.py",
    '''                if cols[6].button(\n                    "Eliminar",\n                    key=f"delete_pending_all_{idx}_{row['id']}",\n                    use_container_width=True,\n                ):\n                    st.session_state["pending_delete_confirm_all"] = row["id"]\n''',
    '''                with cols[6]:\n                    edit_col, delete_col = st.columns(2)\n                    row_requirement = row.get("requerimiento")\n                    if edit_col.button(\n                        "Editar",\n                        key=f"edit_pending_all_{idx}_{row['id']}",\n                        use_container_width=True,\n                    ):\n                        self._show_edit_pending_dialog(\n                            row,\n                            row_requirement,\n                            "all",\n                        )\n                    if delete_col.button(\n                        "Eliminar",\n                        key=f"delete_pending_all_{idx}_{row['id']}",\n                        use_container_width=True,\n                    ):\n                        st.session_state["pending_delete_confirm_all"] = row["id"]\n''',
)

# 5) Vista por requerimiento: botones Editar + Eliminar en la misma celda.
replace_once(
    "views.py",
    '''                if cols[6].button(\n                    "Eliminar",\n                    key=f"delete_pending_{idx}_{row['id']}",\n                    use_container_width=True,\n                ):\n                    st.session_state["pending_delete_confirm"] = row["id"]\n''',
    '''                with cols[6]:\n                    edit_col, delete_col = st.columns(2)\n                    if edit_col.button(\n                        "Editar",\n                        key=f"edit_pending_{idx}_{row['id']}",\n                        use_container_width=True,\n                    ):\n                        self._show_edit_pending_dialog(\n                            row,\n                            requirement,\n                            "individual",\n                        )\n                    if delete_col.button(\n                        "Eliminar",\n                        key=f"delete_pending_{idx}_{row['id']}",\n                        use_container_width=True,\n                    ):\n                        st.session_state["pending_delete_confirm"] = row["id"]\n''',
)

print("Parche aplicado correctamente.")
