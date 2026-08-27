from __future__ import annotations

import threading
from dataclasses import dataclass
from typing import Callable

from google.auth.transport.requests import AuthorizedSession
from google.oauth2 import service_account

from services.github_service import build_store as build_fallback_store


class StorageConflict(RuntimeError):
    pass


@dataclass
class FileVersion:
    content: bytes
    version: str | None


_WRITE_LOCK = threading.Lock()


class GoogleDriveFileStore:
    """Almacenamiento de los dos Excel del proyecto en Google Drive.

    La aplicación sigue trabajando con nombres lógicos de archivo, pero
    internamente cada nombre se resuelve al fileId configurado en Streamlit.
    """

    DRIVE_SCOPE = "https://www.googleapis.com/auth/drive"

    def __init__(self, service_account_info: dict, file_ids: dict[str, str]) -> None:
        info = dict(service_account_info)
        credentials = service_account.Credentials.from_service_account_info(
            info,
            scopes=[self.DRIVE_SCOPE],
        )
        self.session = AuthorizedSession(credentials)
        self.file_ids = file_ids
        self.api_base = "https://www.googleapis.com/drive/v3/files"
        self.upload_base = "https://www.googleapis.com/upload/drive/v3/files"

    def _file_id(self, path: str) -> str:
        file_id = self.file_ids.get(path)
        if not file_id:
            raise FileNotFoundError(
                f"No existe un fileId de Google Drive configurado para {path}."
            )
        return file_id

    def _metadata(self, file_id: str) -> tuple[str | None, str | None]:
        response = self.session.get(
            f"{self.api_base}/{file_id}",
            params={"fields": "id,name,version,modifiedTime"},
            timeout=30,
        )
        if response.status_code == 404:
            raise FileNotFoundError(file_id)
        response.raise_for_status()
        payload = response.json()
        # Google devuelve ETag en la respuesta HTTP. La versión se conserva
        # también como respaldo informativo.
        etag = response.headers.get("ETag")
        version = payload.get("version")
        return etag, str(version) if version is not None else None

    def read(self, path: str) -> FileVersion:
        file_id = self._file_id(path)
        _, version = self._metadata(file_id)
        response = self.session.get(
            f"{self.api_base}/{file_id}",
            params={"alt": "media"},
            timeout=60,
        )
        if response.status_code == 404:
            raise FileNotFoundError(path)
        response.raise_for_status()
        return FileVersion(content=response.content, version=version)

    def write_new(self, path: str, content: bytes, message: str) -> bytes:
        raise RuntimeError(
            "En modo Google Drive los archivos deben existir previamente y "
            "estar configurados mediante sus fileId."
        )

    def mutate(
        self,
        path: str,
        mutator: Callable[[bytes], bytes],
        message: str,
    ) -> bytes:
        file_id = self._file_id(path)
        last_error: Exception | None = None

        # El lock serializa escrituras simultáneas dentro de la instancia
        # Streamlit. El ETag añade protección frente a cambios externos.
        with _WRITE_LOCK:
            for _ in range(4):
                metadata_response = self.session.get(
                    f"{self.api_base}/{file_id}",
                    params={"fields": "id,name,version,modifiedTime"},
                    timeout=30,
                )
                if metadata_response.status_code == 404:
                    raise FileNotFoundError(path)
                metadata_response.raise_for_status()
                etag = metadata_response.headers.get("ETag")

                download = self.session.get(
                    f"{self.api_base}/{file_id}",
                    params={"alt": "media"},
                    timeout=60,
                )
                download.raise_for_status()
                updated = mutator(download.content)

                headers = {
                    "Content-Type": (
                        "application/vnd.openxmlformats-officedocument."
                        "spreadsheetml.sheet"
                    )
                }
                if etag:
                    headers["If-Match"] = etag

                upload = self.session.patch(
                    f"{self.upload_base}/{file_id}",
                    params={"uploadType": "media"},
                    headers=headers,
                    data=updated,
                    timeout=120,
                )

                if upload.status_code in (200, 201):
                    return updated
                if upload.status_code in (409, 412):
                    last_error = StorageConflict(
                        "El Excel cambió mientras se guardaba. "
                        "La aplicación volvió a leer la versión más reciente."
                    )
                    continue

                upload.raise_for_status()

        raise last_error or StorageConflict(
            "No fue posible guardar el Excel después de varios intentos."
        )


def _plain_dict(value) -> dict:
    if value is None:
        return {}
    try:
        return {k: v for k, v in value.items()}
    except Exception:
        return dict(value)


def build_store(secrets: dict | None = None):
    secrets = secrets or {}

    data_id = secrets.get("GOOGLE_DRIVE_DATA_FILE_ID")
    catalog_id = secrets.get("GOOGLE_DRIVE_CATALOG_FILE_ID")
    service_info = secrets.get("google_service_account")

    if data_id and catalog_id and service_info:
        file_ids = {
            "data/Documentacion_faltante.xlsx": str(data_id),
            "data/Codificacion_documentos.xlsx": str(catalog_id),
        }
        return (
            GoogleDriveFileStore(
                service_account_info=_plain_dict(service_info),
                file_ids=file_ids,
            ),
            "Google Drive",
        )

    # Conserva el modo anterior como respaldo para desarrollo/local.
    return build_fallback_store(secrets)
