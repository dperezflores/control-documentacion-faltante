from pathlib import Path

path = Path("views.py")
text = path.read_text(encoding="utf-8")

replacements = [
    (
        '''                    if edit_col.button(\n                        "Editar",\n                        key=f"edit_pending_all_{idx}_{row['id']}",\n                        use_container_width=True,\n                    ):''',
        '''                    if edit_col.button(\n                        "✏️",\n                        key=f"edit_pending_all_{idx}_{row['id']}",\n                        use_container_width=True,\n                        help="Editar",\n                    ):''',
    ),
    (
        '''                    if delete_col.button(\n                        "Eliminar",\n                        key=f"delete_pending_all_{idx}_{row['id']}",\n                        use_container_width=True,\n                    ):''',
        '''                    if delete_col.button(\n                        "🗑️",\n                        key=f"delete_pending_all_{idx}_{row['id']}",\n                        use_container_width=True,\n                        help="Eliminar",\n                    ):''',
    ),
    (
        '''                    if edit_col.button(\n                        "Editar",\n                        key=f"edit_pending_{idx}_{row['id']}",\n                        use_container_width=True,\n                    ):''',
        '''                    if edit_col.button(\n                        "✏️",\n                        key=f"edit_pending_{idx}_{row['id']}",\n                        use_container_width=True,\n                        help="Editar",\n                    ):''',
    ),
    (
        '''                    if delete_col.button(\n                        "Eliminar",\n                        key=f"delete_pending_{idx}_{row['id']}",\n                        use_container_width=True,\n                    ):''',
        '''                    if delete_col.button(\n                        "🗑️",\n                        key=f"delete_pending_{idx}_{row['id']}",\n                        use_container_width=True,\n                        help="Eliminar",\n                    ):''',
    ),
]

for old, new in replacements:
    if old not in text:
        raise RuntimeError("No se encontró uno de los bloques esperados en views.py")
    text = text.replace(old, new, 1)

path.write_text(text, encoding="utf-8")
print("PATCH_OK")
