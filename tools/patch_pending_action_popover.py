from pathlib import Path

path = Path("views.py")
text = path.read_text(encoding="utf-8")

replacements = [
    (
        '''                with cols[6]:\n                    edit_col, delete_col = st.columns(2)\n                    row_requirement = row.get("requerimiento")\n                    if edit_col.button(\n                        "✏️",\n                        key=f"edit_pending_all_{idx}_{row['id']}",\n                        use_container_width=True,\n                        help="Editar",\n                    ):\n                        self._show_edit_pending_dialog(\n                            row,\n                            row_requirement,\n                            "all",\n                        )\n                    if delete_col.button(\n                        "🗑️",\n                        key=f"delete_pending_all_{idx}_{row['id']}",\n                        use_container_width=True,\n                        help="Eliminar",\n                    ):\n                        st.session_state["pending_delete_confirm_all"] = row["id"]\n''',
        '''                with cols[6]:\n                    row_requirement = row.get("requerimiento")\n                    with st.popover("⋮", use_container_width=True):\n                        if st.button(\n                            "✏️ Editar",\n                            key=f"edit_pending_all_{idx}_{row['id']}",\n                            use_container_width=True,\n                        ):\n                            self._show_edit_pending_dialog(\n                                row,\n                                row_requirement,\n                                "all",\n                            )\n                        if st.button(\n                            "🗑️ Eliminar",\n                            key=f"delete_pending_all_{idx}_{row['id']}",\n                            use_container_width=True,\n                        ):\n                            st.session_state["pending_delete_confirm_all"] = row["id"]\n''',
    ),
    (
        '''                with cols[6]:\n                    edit_col, delete_col = st.columns(2)\n                    if edit_col.button(\n                        "✏️",\n                        key=f"edit_pending_{idx}_{row['id']}",\n                        use_container_width=True,\n                        help="Editar",\n                    ):\n                        self._show_edit_pending_dialog(\n                            row,\n                            requirement,\n                            "individual",\n                        )\n                    if delete_col.button(\n                        "🗑️",\n                        key=f"delete_pending_{idx}_{row['id']}",\n                        use_container_width=True,\n                        help="Eliminar",\n                    ):\n                        st.session_state["pending_delete_confirm"] = row["id"]\n''',
        '''                with cols[6]:\n                    with st.popover("⋮", use_container_width=True):\n                        if st.button(\n                            "✏️ Editar",\n                            key=f"edit_pending_{idx}_{row['id']}",\n                            use_container_width=True,\n                        ):\n                            self._show_edit_pending_dialog(\n                                row,\n                                requirement,\n                                "individual",\n                            )\n                        if st.button(\n                            "🗑️ Eliminar",\n                            key=f"delete_pending_{idx}_{row['id']}",\n                            use_container_width=True,\n                        ):\n                            st.session_state["pending_delete_confirm"] = row["id"]\n''',
    ),
]

for old, new in replacements:
    if old not in text:
        raise RuntimeError("No se encontró uno de los bloques de acciones esperado en views.py")
    text = text.replace(old, new, 1)

path.write_text(text, encoding="utf-8")
print("PATCH_OK")
