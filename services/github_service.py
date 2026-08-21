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
    def read(self, path: str) -> FileVersion:
        p = Path(path)
        if not p.exists():
            raise FileNotFoundError(path)
        return FileVersion(p.read_bytes(), None)

    def write_new(self, path: str, content: bytes, message: str) -> bytes:
        p = Path(path)
        p.parent.mkdir(parents=True, exist_ok=True)
        p.write_bytes(content)
        return content

    def mutate(self, path: str, mutator: Callable[[bytes], bytes], message: str) -> bytes:
        p = Path(path)
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

    def mutate(self, path: str, mutator: Callable[[bytes], bytes], message: str) -> bytes:
        last_error: Exception | None = None
        for _ in range(3):
            current = self.read(path)
            updated = mutator(current.content)
            payload = {
                "message": message,
                "content": base64.b64encode(updated).decode("ascii"),
                "sha": current.version,
                "branch": self.branch,
            }
            response = requests.put(
                f"{self.base_url}/{path}",
                headers=self.headers,
                json=payload,
                timeout=60,
            )
            if response.status_code in (200, 201):
                return updated
            if response.status_code in (409, 422):
                last_error = StorageConflict("El archivo cambió mientras se guardaba; reintentando.")
                continue
            response.raise_for_status()
        raise last_error or StorageConflict("No fue posible guardar después de varios intentos.")


def build_store(secrets: dict | None = None):
    secrets = secrets or {}
    token = secrets.get("GITHUB_TOKEN") or os.getenv("GITHUB_TOKEN")
    owner = secrets.get("GITHUB_OWNER") or os.getenv("GITHUB_OWNER") or "dperezflores"
    repo = secrets.get("GITHUB_REPO") or os.getenv("GITHUB_REPO") or "control-documentacion-faltante"
    branch = secrets.get("GITHUB_BRANCH") or os.getenv("GITHUB_BRANCH") or "main"
    if token:
        return GitHubFileStore(token=token, owner=owner, repo=repo, branch=branch), "GitHub"
    return LocalFileStore(), "Local"
