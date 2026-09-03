from __future__ import annotations

import base64
import os
from dataclasses import dataclass
from pathlib import Path
from typing import Callable

import requests


class StorageConflict(RuntimeError):
    pass


@dataclass
class FileVersion:
    content: bytes
    version: str | None


class LocalFileStore:
    def get_version(self, path: str) -> str | None:
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(path)
        stat = p.stat()
        return f"{stat.st_mtime_ns}:{stat.st_size}"

    def read(self, path: str) -> FileVersion:
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(path)
        return FileVersion(p.read_bytes(), self.get_version(path))

    def write_new(self, path: str, content: bytes, message: str) -> bytes:
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(content)
        return content

    def mutate(
        self,
        path: str,
        mutator: Callable[[bytes], bytes],
        message: str,
        cached_content: bytes | None = None,
        cached_version: str | None = None,
    ) -> bytes:
        p = Path(path)
        current_version = self.get_version(path)
        if (
            cached_content is not None
            and cached_version is not None
            and current_version == cached_version
        ):
            current = cached_content
        else:
            current = p.read_bytes()
        updated = mutator(current)
        p.write_bytes(updated)
        return updated


class GitHubFileStore:
    def __init__(self, token: str, owner: str, repo: str, branch: str = "main") -> None:
        self.token = token
        self.owner = owner
        self.repo = repo
        self.branch = branch
        self.base_url = f"https://api.github.com/repos/{owner}/{repo}/contents"
        self.headers = {
            "Authorization": f"Bearer {token}",
            "Accept": "application/vnd.github+json",
            "X-GitHub-Api-Version": "2022-11-28",
        }

    def get_version(self, path: str) -> str | None:
        response = requests.get(
            f"https://api.github.com/repos/{self.owner}/{self.repo}/commits",
            headers=self.headers,
            params={"sha": self.branch, "path": path, "per_page": 1},
            timeout=30,
        )
        if response.status_code == 404:
            raise FileNotFoundError(path)
        response.raise_for_status()
        commits = response.json()
        if not commits:
            raise FileNotFoundError(path)
        return str(commits[0]["sha"])

    def read(self, path: str) -> FileVersion:
        response = requests.get(
            f"{self.base_url}/{path}",
            headers=self.headers,
            params={"ref": self.branch},
            timeout=30,
        )
        if response.status_code == 404:
            raise FileNotFoundError(path)
        response.raise_for_status()
        payload = response.json()
        content = base64.b64decode(payload["content"])
        return FileVersion(content=content, version=payload["sha"])

    def write_new(self, path: str, content: bytes, message: str) -> bytes:
        payload = {
            "message": message,
            "content": base64.b64encode(content).decode("ascii"),
            "branch": self.branch,
        }
        response = requests.put(
            f"{self.base_url}/{path}",
            headers=self.headers,
            json=payload,
            timeout=60,
        )
        if response.status_code == 422:
            raise StorageConflict(f"El archivo {path} ya existe en GitHub.")
        response.raise_for_status()
        return content

    def mutate(
        self,
        path: str,
        mutator: Callable[[bytes], bytes],
        message: str,
        cached_content: bytes | None = None,
        cached_version: str | None = None,
    ) -> bytes:
        last_error: Exception | None = None
        cache_is_trusted = cached_content is not None and cached_version is not None
        self.last_mutation_version: str | None = None

        for _ in range(3):
            # GitHub necesita el blob SHA actual para el PUT. La consulta al
            # endpoint contents proporciona ese SHA y sirve también para
            # validar que la versión remota sigue siendo la esperada.
            response = requests.get(
                f"{self.base_url}/{path}",
                headers=self.headers,
                params={"ref": self.branch},
                timeout=30,
            )
            if response.status_code == 404:
                raise FileNotFoundError(path)
            response.raise_for_status()
            payload_current = response.json()
            blob_sha = str(payload_current["sha"])

            if cache_is_trusted:
                remote_version = self.get_version(path)
            else:
                remote_version = None

            if cache_is_trusted and remote_version == cached_version:
                current_content = cached_content
            else:
                current_content = base64.b64decode(payload_current["content"])

            updated = mutator(current_content)
            payload = {
                "message": message,
                "content": base64.b64encode(updated).decode("ascii"),
                "sha": blob_sha,
                "branch": self.branch,
            }
            write_response = requests.put(
                f"{self.base_url}/{path}",
                headers=self.headers,
                json=payload,
                timeout=60,
            )
            if write_response.status_code in (200, 201):
                body = write_response.json()
                self.last_mutation_version = str(
                    body.get("commit", {}).get("sha") or ""
                ) or None
                return updated
            if write_response.status_code in (409, 422):
                last_error = StorageConflict("El archivo cambió mientras se guardaba; reintentando.")
                cache_is_trusted = False
                continue
            write_response.raise_for_status()
        raise last_error or StorageConflict("No fue posible guardar después de varios intentos.")


def build_store(secrets: dict | None = None):
    secrets = secrets or {}

    # Si existen los secretos de Google Drive, éste se convierte en el
    # almacenamiento operativo. GitHub queda sólo como repositorio de código.
    try:
        from services.google_drive_service import build_google_drive_store
        google_store = build_google_drive_store(secrets)
        if google_store is not None:
            return google_store, "Google Drive"
    except Exception:
        # Si la configuración de Google existe pero es inválida, dejamos que
        # el error aparezca al intentar leer para facilitar el diagnóstico.
        if (
            secrets.get("GOOGLE_DRIVE_DATA_FILE_ID")
            or secrets.get("GOOGLE_DRIVE_CATALOG_FILE_ID")
            or secrets.get("google_service_account")
        ):
            raise

    token = secrets.get("GITHUB_TOKEN") or os.getenv("GITHUB_TOKEN")
    owner = secrets.get("GITHUB_OWNER") or os.getenv("GITHUB_OWNER") or "dperezflores"
    repo = secrets.get("GITHUB_REPO") or os.getenv("GITHUB_REPO") or "control-documentacion-faltante"
    branch = secrets.get("GITHUB_BRANCH") or os.getenv("GITHUB_BRANCH") or "main"
    if token:
        return GitHubFileStore(token=token, owner=owner, repo=repo, branch=branch), "GitHub"
    return LocalFileStore(), "Local"
