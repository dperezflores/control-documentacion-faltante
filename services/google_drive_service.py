from __future__ import annotations

import threading
from dataclasses import dataclass
from typing import Callable

from google.auth.transport.requests import AuthorizedSession
from google.oauth2 import service_account


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
            params={"fields": "id,name,version,modifiedTime,md5Checksum"},
            timeout=30,
        )
        if response.status_code == 404:
            raise FileNotFoundError(file_id)
        response.raise_for_status()
        payload = response.json()
        # Google devuelve ETag en la respuesta HTTP. La versión se conserva
        # también como respaldo informativo.
        etag = response.headers.get("ETag")
        version = "|".join(
            str(payload.get(key) or "")
            for key in ("modifiedTime", "md5Checksum", "version")
        )
        return etag, version

    def get_version(self, path: str) -> str | None:
        file_id = self._file_id(path)
        _, version = self._metadata(file_id)
        return version

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
        cached_content: bytes | None = None,
        cached_version: str | None = None,
    ) -> bytes:
        file_id = self._file_id(path)
        last_error: Exception | None = None
        cache_is_trusted = cached_content is not None and cached_version is not None
        self.last_mutation_version: str | None = None

        # El lock serializa escrituras simultáneas dentro de la instancia
        # Streamlit. El ETag añade protección frente a cambios externos.
        with _WRITE_LOCK:
            for _ in range(4):
                etag, remote_version = self._metadata(file_id)

                if cache_is_trusted and remote_version == cached_version:
                    current_content = cached_content
                else:
                    download = self.session.get(
                        f"{self.api_base}/{file_id}",
                        params={"alt": "media"},
                        timeout=60,
                    )
                    download.raise_for_status()
                    current_content = download.content

                updated = mutator(current_content)

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
                    params={
                        "uploadType": "media",
                        "fields": "id,version,modifiedTime,md5Checksum",
                    },
                    headers=headers,
                    data=updated,
                    timeout=120,
                )

                if upload.status_code in (200, 201):
                    payload = upload.json() if upload.content else {}
                    self.last_mutation_version = "|".join(
                        str(payload.get(key) or "")
                        for key in ("modifiedTime", "md5Checksum", "version")
                    )
                    return updated
                if upload.status_code in (409, 412):
                    last_error = StorageConflict(
                        "El Excel cambió mientras se guardaba. "
                        "La aplicación volvió a leer la versión más reciente."
                    )
                    # Después del primer conflicto, la copia local deja de ser
                    # confiable durante el resto de esta llamada.
                    cache_is_trusted = False
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


def build_google_drive_store(secrets: dict):
    data_id = secrets.get("GOOGLE_DRIVE_DATA_FILE_ID")
    catalog_id = secrets.get("GOOGLE_DRIVE_CATALOG_FILE_ID")
    service_info = secrets.get("google_service_account")
    if not (data_id and catalog_id and service_info):
        return None

    file_ids = {
        "data/Documentacion_faltante.xlsx": str(data_id),
        "data/Codificacion_documentos.xlsx": str(catalog_id),
    }
    return GoogleDriveFileStore(
        service_account_info=_plain_dict(service_info),
        file_ids=file_ids,
    )
