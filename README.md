# Control de documentación faltante

Prototipo en Python + Streamlit para capturar documentación faltante por requerimiento y contrato, homologarla con el catálogo de codificación, registrar cortes y generar borradores de oficio.

## Estado

Versión piloto 0.1.

## Funciones incluidas

1. Selección de requerimiento (`Req_003_ULOP`, `Req_004`, `Req_008`, etc.).
2. Selección de contrato.
3. Selección del procedimiento de adjudicación: DIR, LPU o LSI.
4. Búsqueda y selección múltiple de documentos desde el Excel de codificación.
5. Escritura automática en la columna `Documentación faltante` sin cambiar el formato visible de las hojas de requerimiento.
6. Registro técnico en hojas ocultas `_APP_FALTANTES`, `_APP_CORTES` y `_APP_OFICIOS`.
7. Creación de cortes con todos los faltantes pendientes al momento del corte.
8. Selección de uno o varios cortes para consolidarlos en un solo Word.
9. Persistencia opcional directamente en GitHub con control optimista de concurrencia.

## Ejecución local

```bash
pip install -r requirements.txt
streamlit run app.py
```

Si los Excel todavía no existen en `data/`, la propia aplicación permite cargarlos e inicializarlos.

## Despliegue en Streamlit Community Cloud

El archivo principal es `app.py`.

Configure estos secretos en la aplicación desplegada:

```toml
GITHUB_TOKEN = "github_pat_xxx"
GITHUB_OWNER = "dperezflores"
GITHUB_REPO = "control-documentacion-faltante"
GITHUB_BRANCH = "main"
```

El token debe tener permiso de lectura y escritura sobre el contenido del repositorio privado. No debe guardarse el token dentro del repositorio.

## Archivos de datos

La aplicación espera:

- `data/Documentacion_faltante.xlsx`
- `data/Codificacion_documentos.xlsx`

Como el conector utilizado para construir este prototipo no carga binarios directamente al repositorio, la primera ejecución de la aplicación muestra una pantalla de inicialización para subir ambos Excel una sola vez. Desde ese momento la aplicación trabaja sobre los archivos guardados en GitHub.

## Arquitectura

- `app.py`: interfaz Streamlit.
- `services/excel_service.py`: lectura/escritura de Excel, faltantes y cortes.
- `services/github_service.py`: almacenamiento local o persistencia GitHub.
- `services/word_service.py`: generación del Word del prototipo.

La capa de almacenamiento está separada para poder sustituir GitHub por SharePoint posteriormente sin rehacer la lógica de negocio.
