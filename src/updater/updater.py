"""
Modulo di autoupdate: verifica in background (QThread) l'ultima release
pubblicata su GitHub e la confronta con la versione locale in version.txt.
Nessuna chiamata di rete deve mai avvenire nel thread principale della GUI.
"""

from __future__ import annotations

import json
import re
import urllib.error
import urllib.request
from dataclasses import dataclass
from pathlib import Path
from typing import Optional, Tuple

from PyQt6.QtCore import QThread, pyqtSignal

GITHUB_OWNER = "enkas79"
GITHUB_REPO = "exe-create"
GITHUB_API_LATEST_RELEASE = f"https://api.github.com/repos/{GITHUB_OWNER}/{GITHUB_REPO}/releases/latest"
REQUEST_TIMEOUT_SECONDS = 10


def read_local_version(version_file: Path) -> str:
    try:
        text = version_file.read_text(encoding="utf-8").strip()
        return text.splitlines()[0].strip() if text else "0.0.0"
    except Exception:
        return "0.0.0"


def parse_version(version: str) -> Tuple[int, ...]:
    numbers = re.findall(r"\d+", version)
    return tuple(int(n) for n in numbers) if numbers else (0,)


def is_newer(remote_version: str, local_version: str) -> bool:
    return parse_version(remote_version) > parse_version(local_version)


@dataclass
class UpdateInfo:
    version: str
    download_url: Optional[str]
    changelog: str
    html_url: str


class UpdateCheckThread(QThread):
    """Interroga l'API GitHub Releases senza bloccare la GUI."""

    update_available = pyqtSignal(object)  # UpdateInfo
    up_to_date = pyqtSignal(str)
    check_failed = pyqtSignal(str)

    def __init__(self, local_version: str):
        super().__init__()
        self.local_version = local_version

    def run(self) -> None:
        try:
            request = urllib.request.Request(
                GITHUB_API_LATEST_RELEASE,
                headers={"Accept": "application/vnd.github+json", "User-Agent": "ToExe-Universal-Simple"},
            )
            with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
                payload = json.loads(response.read().decode("utf-8"))
        except urllib.error.URLError as exc:
            self.check_failed.emit(f"Connessione non disponibile: verifica aggiornamenti non riuscita.\n{exc}")
            return
        except Exception as exc:
            self.check_failed.emit(f"Errore durante la verifica aggiornamenti: {exc}")
            return

        remote_version = str(payload.get("tag_name", "")).lstrip("vV").strip()
        if not remote_version:
            self.check_failed.emit("Risposta GitHub non valida: versione remota non trovata.")
            return

        if not is_newer(remote_version, self.local_version):
            self.up_to_date.emit(self.local_version)
            return

        download_url = None
        for asset in payload.get("assets", []):
            name = str(asset.get("name", "")).lower()
            if name.endswith((".exe", ".msi", ".zip")):
                download_url = asset.get("browser_download_url")
                break

        info = UpdateInfo(
            version=remote_version,
            download_url=download_url,
            changelog=str(payload.get("body", "") or "Nessuna nota di rilascio disponibile."),
            html_url=str(payload.get("html_url", "")),
        )
        self.update_available.emit(info)


class UpdateDownloadThread(QThread):
    """Scarica l'installer aggiornato mostrando l'avanzamento alla GUI."""

    progress = pyqtSignal(int)
    finished_signal = pyqtSignal(bool, str)

    def __init__(self, download_url: str, destination: Path):
        super().__init__()
        self.download_url = download_url
        self.destination = destination

    def run(self) -> None:
        try:
            request = urllib.request.Request(
                self.download_url,
                headers={"User-Agent": "ToExe-Universal-Simple"},
            )
            with urllib.request.urlopen(request, timeout=REQUEST_TIMEOUT_SECONDS) as response:
                total_size = int(response.headers.get("Content-Length", 0))
                downloaded = 0
                self.destination.parent.mkdir(parents=True, exist_ok=True)
                with open(self.destination, "wb") as output_file:
                    while True:
                        chunk = response.read(65536)
                        if not chunk:
                            break
                        output_file.write(chunk)
                        downloaded += len(chunk)
                        if total_size:
                            self.progress.emit(int(downloaded * 100 / total_size))
            self.finished_signal.emit(True, str(self.destination))
        except Exception as exc:
            self.finished_signal.emit(False, f"Errore durante il download dell aggiornamento: {exc}")
