# Control de documentación faltante

Prototipo en Python + Streamlit para capturar documentación faltante por requerimiento y contrato, homologarla con el catálogo de codificación, registrar cortes y preparar la generación de oficios.

## Estado

Versión piloto 0.1.

## Flujo inicial

1. Seleccionar requerimiento.
2. Seleccionar contrato.
3. Seleccionar procedimiento de adjudicación: DIR, LPU o LSI.
4. Elegir documentos faltantes desde el catálogo oficial.
5. Guardar los faltantes en el Excel operativo sin alterar el diseño visible.
6. Registrar cada faltante en una hoja técnica interna.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

## Datos

- `data/Documentacion_faltante.xlsx`: archivo operativo.
- `data/Codificacion_documentos.xlsx`: catálogo DIR/LPU/LSI.

## Persistencia piloto

La aplicación está preparada para trabajar con archivos Excel. La integración de escritura persistente con GitHub se incorpora como capa separada para poder sustituirla posteriormente por SharePoint sin rehacer la lógica principal.
