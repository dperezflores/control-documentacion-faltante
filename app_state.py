from __future__ import annotations

from typing import Callable
from zoneinfo import ZoneInfo

import streamlit as st

from services.github_service import build_store


DATA_FILE = "data/Documentacion_faltante.xlsx"
CATALOG_FILE = "data/Codificacion_documentos.xlsx"


class AppState:
    def __init__(self, secrets: dict | None = None) -> None:
        self.secrets = secrets or {}
        self.store, self.storage_mode = build_store(self.secrets)
        self.timezone = ZoneInfo(
            str(self.secrets.get("timezone", "America/Mexico_City"))
        )

    @property
    def session(self):
        return st.session_state

    def _remote_version(self, path: str) -> str | None:
        getter = getattr(self.store, "get_version", None)
        if getter is None:
            return None
        return getter(path)

    def _load_operational(self) -> bytes:
        remote_version = self._remote_version(DATA_FILE)
        cached_bytes = self.session.get("operational_bytes")
        cached_version = self.session.get("operational_version")

        if (
            cached_bytes is not None
            and remote_version is not None
            and cached_version == remote_version
        ):
            return cached_bytes

        current = self.store.read(DATA_FILE)
        self.session.operational_bytes = current.content
        self.session.operational_version = (
            remote_version if remote_version is not None else current.version
        )
        return current.content

    def _load_catalog(self, force: bool = False) -> bytes:
        remote_version = self._remote_version(CATALOG_FILE)
        if not force:
            cached_bytes = self.session.get("catalog_bytes")
            cached_version = self.session.get("catalog_version")
            if (
                cached_bytes is not None
                and remote_version is not None
                and cached_version == remote_version
            ):
                return cached_bytes

        current = self.store.read(CATALOG_FILE)
        self.session.catalog_bytes = current.content
        self.session.catalog_version = (
            remote_version if remote_version is not None else current.version
        )
        return current.content

    def load_files(self, force: bool = False) -> tuple[bytes, bytes] | None:
        try:
            operational = self._load_operational()
            catalog = self._load_catalog(force=force)
            return operational, catalog
        except FileNotFoundError:
            return None
        except Exception as exc:
            st.error(f"No fue posible cargar los archivos de datos: {exc}")
            st.stop()

    def persist(self, mutator: Callable[[bytes], bytes], message: str) -> bytes:
        updated = self.store.mutate(DATA_FILE, mutator, message)
        self.session.operational_bytes = updated
        try:
            self.session.operational_version = self._remote_version(DATA_FILE)
        except Exception:
            self.session.pop("operational_version", None)
        return updated

    def initialize_files(self, operational: bytes, catalog: bytes) -> None:
        try:
            self.store.read(DATA_FILE)
        except FileNotFoundError:
            self.store.write_new(DATA_FILE, operational, "Inicializar Excel operativo")
        try:
            self.store.read(CATALOG_FILE)
        except FileNotFoundError:
            self.store.write_new(
                CATALOG_FILE,
                catalog,
                "Inicializar catálogo de codificación",
            )

        self.session.operational_bytes = operational
        self.session.catalog_bytes = catalog
        try:
            self.session.operational_version = self._remote_version(DATA_FILE)
        except Exception:
            self.session.pop("operational_version", None)
        try:
            self.session.catalog_version = self._remote_version(CATALOG_FILE)
        except Exception:
            self.session.pop("catalog_version", None)
