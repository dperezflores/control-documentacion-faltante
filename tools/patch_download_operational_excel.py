from pathlib import Path

path = Path("views.py")
text = path.read_text(encoding="utf-8")

old = '''        requirement = self.requirement\n        operational = self.operational\n\n        cuts = list_cuts(operational, requirement)\n'''
new = '''        requirement = self.requirement\n        operational = self.operational\n\n        st.download_button(\n            "Descargar Excel actualizado",\n            data=operational,\n            file_name="Documentacion_faltante.xlsx",\n            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",\n            use_container_width=True,\n            on_click="ignore",\n            help=(\n                "Descarga la copia operativa actualmente sincronizada por la aplicación. "\n                "Para forzar una sincronización inmediata con cambios externos, usa "\n                "'Actualizar datos' en la barra lateral antes de descargar."\n            ),\n        )\n\n        cuts = list_cuts(operational, requirement)\n'''

if old not in text:
    raise RuntimeError("No se encontró el punto esperado en GenerarOficioView.render()")

text = text.replace(old, new, 1)
path.write_text(text, encoding="utf-8")
print("PATCH_OK")
