"""
ToExe Universal Simple - Generatore EXE/Setup semplificato

Versione 0.2.0

Obiettivo:
- selezioni il file main.py del progetto;
- puoi scegliere manualmente un file immagine comune, che verra convertito automaticamente in .ico;
- se non scegli un immagine, il programma usa l icona .ico rilevata automaticamente, se presente;
- il programma rileva automaticamente nome programma, versione, cartella progetto, risorse e alcune dipendenze comuni;
- genera una build PyInstaller in modalita onedir;
- se Inno Setup è installato, genera anche il setup .exe.

Nota:
- va avviato dallo stesso ambiente Python dove funziona il progetto che vuoi trasformare in EXE;
- la modalita onedir e' la scelta piu stabile per progetti con GUI, risorse, librerie scientifiche o file esterni;
- e' pensato come base generica per progetti Python diversi.
"""

from __future__ import annotations

import importlib.util
import os
import re
import shutil
import subprocess
import sys
from dataclasses import dataclass, field
from pathlib import Path
from typing import Callable, Optional, Tuple

from PyQt6.QtCore import QThread, Qt, pyqtSignal, QTimer
from PyQt6.QtGui import QIcon, QTextCursor
from PyQt6.QtWidgets import (
    QApplication,
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QLineEdit,
    QMainWindow,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QTextEdit,
    QVBoxLayout,
    QWidget,
)

APP_VERSION = "0.2.0"
DEFAULT_COMPANY_NAME = "Enrico Martini"
APP_DISPLAY_NAME = "ToExe Universal Simple"
SUPPORTED_ICON_INPUT_SUFFIXES = {".ico", ".png", ".jpg", ".jpeg", ".bmp", ".gif", ".webp", ".tif", ".tiff"}
ICO_SIZES = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]


@dataclass
class SimpleBuildConfig:
    main_script: Path
    project_dir: Path
    app_name: str
    app_version: str
    icon_path: Optional[Path] = None
    icon_source_path: Optional[Path] = None
    output_dir: Optional[Path] = None
    hidden_imports: list[str] = field(default_factory=list)
    collect_all_modules: list[str] = field(default_factory=list)
    required_modules: list[str] = field(default_factory=list)
    add_datas: list[tuple[Path, str]] = field(default_factory=list)
    excluded_modules: list[str] = field(default_factory=list)


class BuildDetector:
    """Rilevamento generico del progetto.

    Questa classe non deve imporre dipendenze o nomi legati a un singolo programma.
    PyInstaller esegue gia una buona analisi statica: qui aggiungiamo solo aiuti
    pratici per nome, versione, risorse e librerie notoriamente problematiche.
    """

    SOURCE_EXCLUDE_DIRS = {
        ".git",
        ".hg",
        ".svn",
        ".mypy_cache",
        ".pytest_cache",
        ".ruff_cache",
        "__pycache__",
        "build",
        "dist",
        "venv",
        ".venv",
        "env",
        ".env",
        "node_modules",
    }

    COMMON_RESOURCE_DIRS = [
        "assets",
        "resources",
        "resource",
        "templates",
        "template",
        "static",
        "images",
        "img",
        "icons",
        "data",
        "config",
        "configs",
    ]

    COMMON_ROOT_DATA_SUFFIXES = {
        ".ini",
        ".json",
        ".yaml",
        ".yml",
        ".toml",
        ".cfg",
        ".txt",
        ".csv",
    }

    REQUIREMENT_TO_IMPORT = {
        "beautifulsoup4": "bs4",
        "opencv-python": "cv2",
        "opencv-contrib-python": "cv2",
        "pillow": "PIL",
        "pyqt6": "PyQt6",
        "pyqt5": "PyQt5",
        "pyside6": "PySide6",
        "pyside2": "PySide2",
        "python-docx": "docx",
        "pywin32": "win32com",
        "scikit-learn": "sklearn",
        "matplotlib": "matplotlib",
        "numpy": "numpy",
        "pandas": "pandas",
        "openpyxl": "openpyxl",
        "lxml": "lxml",
        "requests": "requests",
        "sqlalchemy": "sqlalchemy",
    }

    # Qt non deve essere gestito con --collect-all: i hook ufficiali di
    # PyInstaller includono automaticamente DLL e plugin necessari. La raccolta
    # indiscriminata trascina anche l'intero albero QML di sviluppo (file .obj,
    # .qmlir, sorgenti e directory con percorsi molto lunghi), causando build
    # enormi e possibili errori Inno Setup su Windows.
    COLLECT_ALL_CANDIDATES = {
        "pandas",
        "numpy",
        "openpyxl",
        "lxml",
        "matplotlib",
        "sklearn",
        "scipy",
        "cv2",
    }


    QT_FRAMEWORKS = ("PySide6", "PyQt6", "PyQt5", "PySide2")

    @classmethod
    def detect_qt_frameworks(cls, project_dir: Path) -> list[str]:
        declared = set(cls.requirement_modules(project_dir))
        scanned = cls.scan_source_imports(project_dir)
        found = [fw for fw in cls.QT_FRAMEWORKS if fw in declared or fw in scanned]
        return cls.unique_existing_order(found)

    @classmethod
    def detect_excluded_modules(cls, project_dir: Path) -> list[str]:
        used = set(cls.detect_qt_frameworks(project_dir))
        excluded = [fw for fw in cls.QT_FRAMEWORKS if fw not in used]
        # tkinter viene escluso solo quando il progetto usa chiaramente un framework Qt.
        if used:
            excluded.append("tkinter")
        return excluded

    @staticmethod
    def normalize_app_name(name: str) -> str:
        cleaned = re.sub(r'[<>:"/\\|?*\x00-\x1f]+', "", name, flags=re.UNICODE)
        cleaned = re.sub(r"\s+", " ", cleaned).strip(" .")
        if not cleaned:
            return "Applicazione"
        reserved = {"CON", "PRN", "AUX", "NUL"}
        reserved.update({f"COM{i}" for i in range(1, 10)})
        reserved.update({f"LPT{i}" for i in range(1, 10)})
        if cleaned.upper() in reserved:
            return "Applicazione"
        return cleaned[:80]

    @staticmethod
    def normalize_package_name(name: str) -> str:
        return re.sub(r"[-_.]+", "-", name.strip().lower())

    @staticmethod
    def read_text_safe(path: Path, limit: int = 50000) -> str:
        try:
            return path.read_text(encoding="utf-8", errors="ignore")[:limit]
        except Exception:
            return ""

    @classmethod
    def detect_name_from_metadata(cls, project_dir: Path) -> Optional[str]:
        pyproject = project_dir / "pyproject.toml"
        if pyproject.exists():
            text = cls.read_text_safe(pyproject)
            match = re.search(r"(?ms)^\s*\[project\].*?^\s*name\s*=\s*[\"']([^\"']+)[\"']", text)
            if match:
                return cls.normalize_app_name(match.group(1).replace("-", " ").replace("_", " "))

        setup_cfg = project_dir / "setup.cfg"
        if setup_cfg.exists():
            text = cls.read_text_safe(setup_cfg)
            match = re.search(r"(?ms)^\s*\[metadata\].*?^\s*name\s*=\s*([^\r\n#]+)", text)
            if match:
                return cls.normalize_app_name(match.group(1).strip().replace("-", " ").replace("_", " "))

        setup_py = project_dir / "setup.py"
        if setup_py.exists():
            text = cls.read_text_safe(setup_py)
            match = re.search(r"name\s*=\s*[\"']([^\"']+)[\"']", text)
            if match:
                return cls.normalize_app_name(match.group(1).replace("-", " ").replace("_", " "))
        return None

    @classmethod
    def detect_version(cls, main_script: Path, project_dir: Path) -> str:
        pyproject = project_dir / "pyproject.toml"
        if pyproject.exists():
            text = cls.read_text_safe(pyproject)
            match = re.search(r"(?ms)^\s*\[project\].*?^\s*version\s*=\s*[\"']([^\"']+)[\"']", text)
            if match:
                return match.group(1).strip()

        setup_cfg = project_dir / "setup.cfg"
        if setup_cfg.exists():
            text = cls.read_text_safe(setup_cfg)
            match = re.search(r"(?ms)^\s*\[metadata\].*?^\s*version\s*=\s*([^\r\n#]+)", text)
            if match:
                return match.group(1).strip()

        candidates = [main_script, project_dir / "version.py", project_dir / "__init__.py"]
        for child in project_dir.iterdir() if project_dir.exists() else []:
            if child.is_dir() and (child / "__init__.py").exists():
                candidates.extend([child / "version.py", child / "__init__.py"])

        patterns = [
            r"APP_VERSION\s*=\s*[\"']([^\"']+)[\"']",
            r"__version__\s*=\s*[\"']([^\"']+)[\"']",
            r"VERSION\s*=\s*[\"']([^\"']+)[\"']",
        ]
        for path in candidates:
            if not path.exists():
                continue
            text = cls.read_text_safe(path)
            for pattern in patterns:
                match = re.search(pattern, text)
                if match:
                    return match.group(1).strip()
        return "1.0.0"

    @classmethod
    def is_useful_app_name(cls, name: str) -> bool:
        cleaned = cls.normalize_app_name(name)
        if len(cleaned) < 2:
            return False
        generic = {
            "app",
            "application",
            "applicazione",
            "main",
            "main window",
            "mainwindow",
            "window",
            "programma",
            "python",
            "gui",
            "untitled",
            "senza titolo",
        }
        return cleaned.strip().lower() not in generic

    @classmethod
    def detect_name_from_source_file(cls, path: Path) -> Optional[str]:
        if not path.exists():
            return None

        text = cls.read_text_safe(path)

        # Priorita massima: costanti esplicite del programma.
        # Esempi consigliati: APP_NAME = "Mio Programma", APP_TITLE = "Mio Programma".
        constant_names = [
            "APP_NAME",
            "APP_DISPLAY_NAME",
            "APPLICATION_NAME",
            "PROGRAM_NAME",
            "PROGRAM_TITLE",
            "SOFTWARE_NAME",
            "SOFTWARE_TITLE",
            "WINDOW_TITLE",
            "APP_TITLE",
            "TITLE",
        ]
        for const_name in constant_names:
            pattern = rf"\b{const_name}\s*(?::\s*str\s*)?=\s*[\"']([^\"']+)[\"']"
            match = re.search(pattern, text)
            if match and cls.is_useful_app_name(match.group(1)):
                return cls.normalize_app_name(match.group(1))

        # Seconda priorita: titoli applicazione/finestra usati da GUI comuni.
        # Non e perfetto, ma e piu vicino al nome visibile del programma rispetto al nome cartella.
        call_patterns = [
            r"\bsetApplicationDisplayName\s*\(\s*[\"']([^\"']+)[\"']",
            r"\bsetApplicationName\s*\(\s*[\"']([^\"']+)[\"']",
            r"\bsetWindowTitle\s*\(\s*[\"']([^\"']+)[\"']",
            r"\.title\s*\(\s*[\"']([^\"']+)[\"']",
            r"title\s*=\s*[\"']([^\"']+)[\"']",
        ]
        for pattern in call_patterns:
            match = re.search(pattern, text)
            if match and cls.is_useful_app_name(match.group(1)):
                return cls.normalize_app_name(match.group(1))

        return None

    @classmethod
    def detect_name_from_package_files(cls, project_dir: Path) -> Optional[str]:
        candidates = [project_dir / "version.py", project_dir / "__init__.py"]
        for child in project_dir.iterdir() if project_dir.exists() else []:
            if child.is_dir() and (child / "__init__.py").exists():
                candidates.extend([child / "version.py", child / "__init__.py"])

        for path in candidates:
            found = cls.detect_name_from_source_file(path)
            if found:
                return found
        return None

    @classmethod
    def detect_name_from_project_sources(cls, main_script: Path, project_dir: Path) -> Optional[str]:
        if not project_dir.exists():
            return None

        preferred_stems = {
            "main",
            "app",
            "application",
            "gui",
            "ui",
            "window",
            "main_window",
            "mainwindow",
            "interface",
        }
        files = [p for p in cls.iter_python_files(project_dir, max_files=120) if p != main_script]
        files.sort(key=lambda p: (p.stem.lower() not in preferred_stems, len(p.parts), str(p).lower()))

        for path in files:
            found = cls.detect_name_from_source_file(path)
            if found:
                return found
        return None

    @classmethod
    def detect_app_name(cls, main_script: Path, project_dir: Path) -> str:
        # Ordine corretto per un generatore universale:
        # 1) nome esplicito o titolo visibile nel main.py;
        # 2) costanti nei file di package/versione;
        # 3) titolo/costanti in altri file sorgente del progetto;
        # 4) metadati Python packaging;
        # 5) nome file main, se non e generico;
        # 6) nome cartella solo come ultima risorsa.
        source_name = cls.detect_name_from_source_file(main_script)
        if source_name:
            return source_name

        package_name = cls.detect_name_from_package_files(project_dir)
        if package_name:
            return package_name

        project_source_name = cls.detect_name_from_project_sources(main_script, project_dir)
        if project_source_name:
            return project_source_name

        metadata_name = cls.detect_name_from_metadata(project_dir)
        if metadata_name:
            return metadata_name

        if main_script.stem.lower() not in {"main", "app", "run", "start", "launcher"}:
            return cls.normalize_app_name(main_script.stem.replace("_", " ").replace("-", " "))

        folder_name = project_dir.name.replace("_", " ").replace("-", " ").strip()
        if folder_name.lower() in {"src", "app", "source"}:
            folder_name = main_script.stem
        return cls.normalize_app_name(folder_name or main_script.stem)

    @staticmethod
    def detect_icon(project_dir: Path) -> Optional[Path]:
        preferred_names = [
            "app.ico",
            "icon.ico",
            "logo.ico",
            "main.ico",
            "application.ico",
            f"{project_dir.name}.ico",
            f"{project_dir.name.lower()}.ico",
        ]
        search_dirs = [
            project_dir,
            project_dir / "assets",
            project_dir / "resources",
            project_dir / "icons",
            project_dir / "images",
            project_dir / "static",
        ]
        for directory in search_dirs:
            for name in preferred_names:
                candidate = directory / name
                if candidate.exists() and candidate.is_file():
                    return candidate
        for directory in search_dirs:
            if directory.exists():
                icons = sorted(directory.glob("*.ico"))
                if icons:
                    return icons[0]
        return None

    @classmethod
    def parse_requirements(cls, project_dir: Path) -> set[str]:
        found: set[str] = set()
        req_files = [project_dir / "requirements.txt"]
        for req_path in req_files:
            if not req_path.exists():
                continue
            try:
                for raw_line in req_path.read_text(encoding="utf-8", errors="ignore").splitlines():
                    line = raw_line.strip()
                    if not line or line.startswith("#") or line.startswith("-"):
                        continue
                    pkg = re.split(r"[<>=!~;\[]", line, maxsplit=1)[0].strip()
                    if pkg:
                        found.add(cls.normalize_package_name(pkg))
            except Exception:
                pass

        pyproject = project_dir / "pyproject.toml"
        if pyproject.exists():
            text = cls.read_text_safe(pyproject, limit=120000)
            deps_blocks = re.findall(r"dependencies\s*=\s*\[(.*?)\]", text, flags=re.S)
            for block in deps_blocks:
                for item in re.findall(r"[\"']([^\"']+)[\"']", block):
                    pkg = re.split(r"[<>=!~;\[]", item, maxsplit=1)[0].strip()
                    if pkg:
                        found.add(cls.normalize_package_name(pkg))
        return found

    @classmethod
    def local_top_level_names(cls, project_dir: Path) -> set[str]:
        local: set[str] = set()
        if not project_dir.exists():
            return local
        for child in project_dir.iterdir():
            if child.name in cls.SOURCE_EXCLUDE_DIRS:
                continue
            if child.is_file() and child.suffix == ".py":
                local.add(child.stem)
            elif child.is_dir() and (child / "__init__.py").exists():
                local.add(child.name)
        return local

    @classmethod
    def iter_python_files(cls, project_dir: Path, max_files: int = 250) -> list[Path]:
        files: list[Path] = []
        for path in project_dir.rglob("*.py"):
            parts = set(path.relative_to(project_dir).parts)
            if parts.intersection(cls.SOURCE_EXCLUDE_DIRS):
                continue
            if any(part.endswith("_BuildOutput") for part in parts):
                continue
            files.append(path)
            if len(files) >= max_files:
                break
        return files

    @classmethod
    def scan_source_imports(cls, project_dir: Path) -> set[str]:
        import ast

        found: set[str] = set()
        local = cls.local_top_level_names(project_dir)
        stdlib = set(getattr(sys, "stdlib_module_names", set()))
        stdlib.update(sys.builtin_module_names)
        stdlib.update({"typing_extensions"})

        for py_file in cls.iter_python_files(project_dir):
            text = cls.read_text_safe(py_file, limit=300000)
            try:
                tree = ast.parse(text)
            except Exception:
                continue
            for node in ast.walk(tree):
                name = None
                if isinstance(node, ast.Import):
                    for alias in node.names:
                        name = alias.name.split(".")[0]
                        if name:
                            found.add(name)
                elif isinstance(node, ast.ImportFrom) and node.module and node.level == 0:
                    name = node.module.split(".")[0]
                    if name:
                        found.add(name)

        return {m for m in found if m not in local and m not in stdlib and not m.startswith("_")}

    @classmethod
    def requirement_modules(cls, project_dir: Path) -> list[str]:
        reqs = cls.parse_requirements(project_dir)
        modules: list[str] = []
        for req in sorted(reqs):
            modules.append(cls.REQUIREMENT_TO_IMPORT.get(req, req.replace("-", "_")))
        return cls.unique_existing_order(modules)

    @staticmethod
    def unique_existing_order(items: list[str]) -> list[str]:
        unique: list[str] = []
        for item in items:
            if item and item not in unique:
                unique.append(item)
        return unique

    @classmethod
    def iter_nested_resource_dirs(cls, project_dir: Path) -> list[Path]:
        """Find resource folders also when nested inside Python packages.

        A project may store runtime files in paths such as
        ``my_package/checklist/data``. Previous versions only inspected resource
        folders at project root, so DOCX/XLSX files used by the application were
        omitted from the EXE.
        """
        found: list[Path] = []
        resource_names = {name.casefold() for name in cls.COMMON_RESOURCE_DIRS}

        for current_root, dir_names, _file_names in os.walk(project_dir):
            current = Path(current_root)
            dir_names[:] = [
                name
                for name in dir_names
                if name not in cls.SOURCE_EXCLUDE_DIRS
                and not name.endswith("_BuildOutput")
            ]

            for name in list(dir_names):
                if name.casefold() not in resource_names:
                    continue
                resource_dir = current / name
                found.append(resource_dir)
                # The complete folder will be copied, so nested traversal would
                # only create duplicate --add-data entries.
                dir_names.remove(name)

        return sorted(found, key=lambda path: str(path).casefold())

    @classmethod
    def detect_local_resource_packages(cls, project_dir: Path) -> list[str]:
        """Return dotted package names for resource dirs with __init__.py."""
        packages: list[str] = []
        for resource_dir in cls.iter_nested_resource_dirs(project_dir):
            if not (resource_dir / "__init__.py").exists():
                continue

            relative = resource_dir.relative_to(project_dir)
            cursor = project_dir
            valid_package = True
            for part in relative.parts:
                cursor = cursor / part
                if not (cursor / "__init__.py").exists():
                    valid_package = False
                    break
            if valid_package:
                packages.append(".".join(relative.parts))
        return cls.unique_existing_order(packages)

    @classmethod
    def detect_hidden_imports(cls, project_dir: Path) -> list[str]:
        modules = cls.requirement_modules(project_dir)
        scanned = sorted(cls.scan_source_imports(project_dir))
        modules.extend(scanned)

        # win32com.client viene spesso caricato dinamicamente nei progetti Windows.
        if "win32com" in modules and "win32com.client" not in modules:
            modules.append("win32com.client")

        # Non aggiunge hidden-import inesistenti: evitata una falsa rottura per progetti generici.
        installed_or_known: list[str] = []
        for module in modules:
            try:
                if importlib.util.find_spec(module) is not None:
                    installed_or_known.append(module)
            except Exception:
                continue
        # Resource subpackages can be referenced dynamically through
        # importlib.resources and therefore may not appear in the AST import
        # scan. Include the local package explicitly when it exists.
        installed_or_known.extend(cls.detect_local_resource_packages(project_dir))
        return cls.unique_existing_order(installed_or_known)

    @classmethod
    def detect_collect_all_modules(cls, project_dir: Path) -> list[str]:
        req_modules = set(cls.requirement_modules(project_dir))
        scanned = cls.scan_source_imports(project_dir)
        candidates = req_modules.union(scanned).intersection(cls.COLLECT_ALL_CANDIDATES)
        # Non aggiungere mai i framework Qt a --collect-all. PyInstaller dispone
        # di hook dedicati per PySide/PyQt e raccoglie il runtime in modo selettivo.
        candidates.difference_update(cls.QT_FRAMEWORKS)
        modules: list[str] = []
        for module in sorted(candidates):
            try:
                if importlib.util.find_spec(module) is not None:
                    modules.append(module)
            except Exception:
                pass
        return cls.unique_existing_order(modules)

    @classmethod
    def detect_required_modules(cls, project_dir: Path) -> list[str]:
        # Per un uso universale controlliamo solo le dipendenze dichiarate.
        # Le importazioni facoltative o dinamiche nel codice non devono bloccare a priori la build.
        return cls.requirement_modules(project_dir)

    @classmethod
    def detect_add_datas(cls, project_dir: Path) -> list[tuple[Path, str]]:
        datas: list[tuple[Path, str]] = []

        # Include resource directories at every project depth, preserving their
        # path relative to the project root. Example:
        # collaudo_suite/checklist/data -> collaudo_suite/checklist/data
        for resource_dir in cls.iter_nested_resource_dirs(project_dir):
            destination = str(resource_dir.relative_to(project_dir))
            datas.append((resource_dir, destination))

        for file_path in project_dir.iterdir() if project_dir.exists() else []:
            if file_path.is_file() and file_path.suffix.lower() in cls.COMMON_ROOT_DATA_SUFFIXES:
                # requirements e metadata servono alla build, non quasi mai all'app finale.
                if file_path.name.lower() in {"requirements.txt", "requirements-dev.txt", "pyproject.toml", "setup.cfg"}:
                    continue
                datas.append((file_path, file_path.name))

        unique: list[tuple[Path, str]] = []
        seen: set[tuple[str, str]] = set()
        for src, dst in datas:
            key = (str(src.resolve()), dst)
            if key not in seen:
                unique.append((src, dst))
                seen.add(key)
        return unique

    @classmethod
    def build_config(cls, main_script: Path) -> SimpleBuildConfig:
        project_dir = main_script.parent
        app_name = cls.detect_app_name(main_script, project_dir)
        version = cls.detect_version(main_script, project_dir)
        icon = cls.detect_icon(project_dir)
        output_dir = project_dir / f"{app_name}_BuildOutput"
        return SimpleBuildConfig(
            main_script=main_script,
            project_dir=project_dir,
            app_name=app_name,
            app_version=version,
            icon_path=icon,
            icon_source_path=icon,
            output_dir=output_dir,
            hidden_imports=cls.detect_hidden_imports(project_dir),
            collect_all_modules=cls.detect_collect_all_modules(project_dir),
            required_modules=cls.detect_required_modules(project_dir),
            add_datas=cls.detect_add_datas(project_dir),
            excluded_modules=cls.detect_excluded_modules(project_dir),
        )


class IconConversionService:
    @staticmethod
    def is_supported(path: Path) -> bool:
        return path.suffix.lower() in SUPPORTED_ICON_INPUT_SUFFIXES

    @staticmethod
    def pillow_install_message() -> str:
        return (
            "Per convertire immagini in .ico serve Pillow.\n\n"
            "Installa nel tuo ambiente Python con:\n"
            "python -m pip install pillow"
        )

    @staticmethod
    def make_safe_icon_name(app_name: str) -> str:
        safe = BuildDetector.normalize_app_name(app_name).replace(" ", "_")
        safe = re.sub(r"[^A-Za-z0-9_.-]+", "_", safe).strip("._")
        return safe or "app"

    @classmethod
    def convert_to_ico(cls, source_path: Path, output_dir: Path, app_name: str) -> Tuple[bool, str, Optional[Path]]:
        if not source_path.exists() or not source_path.is_file():
            return False, "Il file immagine selezionato non esiste.", None

        if not cls.is_supported(source_path):
            return (
                False,
                "Formato non supportato. Usa PNG, JPG, JPEG, BMP, GIF, WEBP, TIFF oppure ICO.",
                None,
            )

        if source_path.suffix.lower() == ".ico":
            return True, "File .ico gia pronto: non serve conversione.", source_path

        try:
            from PIL import Image, ImageOps
        except Exception:
            return False, cls.pillow_install_message(), None

        try:
            icon_dir = output_dir / "icons"
            icon_dir.mkdir(parents=True, exist_ok=True)
            ico_path = icon_dir / f"{cls.make_safe_icon_name(app_name)}_icon.ico"

            with Image.open(source_path) as img:
                img = ImageOps.exif_transpose(img)
                img = img.convert("RGBA")

                side = max(img.width, img.height)
                square = Image.new("RGBA", (side, side), (0, 0, 0, 0))
                square.paste(img, ((side - img.width) // 2, (side - img.height) // 2), img)
                square.save(ico_path, format="ICO", sizes=ICO_SIZES)

            if not ico_path.exists():
                return False, "Conversione terminata, ma il file .ico non e stato creato.", None

            return True, f"Icona convertita in .ico: {ico_path}", ico_path
        except Exception as exc:
            return False, f"Errore durante la conversione dell immagine in .ico: {exc}", None


class BuilderService:
    @staticmethod
    def get_pyinstaller_command() -> Tuple[bool, list[str], str]:
        current = [sys.executable, "-m", "PyInstaller"]
        try:
            check = subprocess.run(
                current + ["--version"],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                timeout=20,
            )
            if check.returncode == 0:
                version = (check.stdout or "").strip().splitlines()[0] if check.stdout else "versione sconosciuta"
                return True, current, f"PyInstaller ambiente corrente: {version}"
        except Exception:
            pass

        found = shutil.which("pyinstaller")
        if found:
            return True, [found], f"PyInstaller trovato nel PATH: {found}"
        return False, [], "PyInstaller non trovato. Installa con: pip install pyinstaller"

    @staticmethod
    def find_inno_compiler() -> Optional[str]:
        found = shutil.which("iscc")
        if found:
            return found
        for path in [
            r"C:\Program Files (x86)\Inno Setup 6\iscc.exe",
            r"C:\Program Files\Inno Setup 6\iscc.exe",
            r"C:\Program Files (x86)\Inno Setup 5\iscc.exe",
        ]:
            if os.path.exists(path):
                return path
        return None

    @staticmethod
    def add_data_argument(source: Path, dest: str) -> str:
        sep = ";" if os.name == "nt" else ":"
        return f"{source}{sep}{dest}"

    @staticmethod
    def version_tuple(version: str) -> str:
        parts = [p for p in version.replace("-", ".").split(".") if p.isdigit()]
        parts += ["0"] * (4 - len(parts))
        return ", ".join(parts[:4])

    @classmethod
    def create_version_info_file(cls, cfg: SimpleBuildConfig) -> Optional[Path]:
        path = cfg.output_dir / "version_info.txt"
        try:
            v = cls.version_tuple(cfg.app_version)
            safe_name = repr(cfg.app_name)
            safe_version = repr(cfg.app_version)
            safe_file = repr(f"{cfg.app_name}.exe")
            content = f"""VSVersionInfo(
  ffi=FixedFileInfo(filevers=({v}), prodvers=({v}), mask=0x3f, flags=0x0, OS=0x40004, fileType=0x1, subtype=0x0, date=(0, 0)),
  kids=[
    StringFileInfo([
      StringTable('040904B0', [
        StringStruct('CompanyName', {repr(DEFAULT_COMPANY_NAME)}),
        StringStruct('FileDescription', {safe_name}),
        StringStruct('FileVersion', {safe_version}),
        StringStruct('InternalName', {safe_name}),
        StringStruct('OriginalFilename', {safe_file}),
        StringStruct('ProductName', {safe_name}),
        StringStruct('ProductVersion', {safe_version})])
      ]),
    VarFileInfo([VarStruct('Translation', [1033, 1200])])
  ]
)"""
            path.write_text(content, encoding="utf-8")
            return path
        except Exception:
            return None

    @staticmethod
    def validate_build_environment(cfg: SimpleBuildConfig) -> Tuple[bool, str]:
        missing: list[str] = []
        for module in cfg.required_modules:
            try:
                if importlib.util.find_spec(module) is None:
                    missing.append(module)
            except Exception:
                missing.append(module)

        if not missing:
            return True, "Ambiente Python valido: dipendenze principali trovate."

        requirements = cfg.project_dir / "requirements.txt"
        install_cmd = f"python -m pip install -r \"{requirements}\"" if requirements.exists() else "python -m pip install <nome-pacchetto-mancante>"
        return (
            False,
            "Mancano moduli nell ambiente Python usato per avviare ToExe: "
            + ", ".join(missing)
            + "\n\nInstalla le dipendenze e poi ricrea la build pulita:\n"
            + install_cmd,
        )

    @classmethod
    def run_pyinstaller(cls, cfg: SimpleBuildConfig, log: Callable[[str], None]) -> Tuple[bool, str]:
        assert cfg.output_dir is not None
        cfg.output_dir.mkdir(parents=True, exist_ok=True)

        env_ok, env_msg = cls.validate_build_environment(cfg)
        log(env_msg)
        if not env_ok:
            return False, env_msg

        dist_path = cfg.output_dir / "dist"
        build_path = cfg.output_dir / "build"
        spec_path = cfg.output_dir / "spec"

        # Elimina completamente la build precedente. PyInstaller --clean pulisce
        # la cache di lavoro, ma non garantisce che una vecchia cartella onedir
        # non lasci file extra. Questo evita che restino alberi QML generati da
        # versioni precedenti del generatore.
        old_app_dir = dist_path / cfg.app_name
        for stale_path in (old_app_dir, build_path):
            if stale_path.exists():
                log(f"Pulizia build precedente: {stale_path}")
                try:
                    shutil.rmtree(stale_path)
                except Exception as exc:
                    return False, f"Impossibile eliminare la vecchia build: {stale_path}\n{exc}"

        spec_path.mkdir(parents=True, exist_ok=True)

        ok, command_base, msg = cls.get_pyinstaller_command()
        log(msg)
        if not ok:
            return False, msg

        version_info = cls.create_version_info_file(cfg)

        command = [
            *command_base,
            "--noconfirm",
            "--noconsole",
            "--onedir",
            "--name",
            cfg.app_name,
            "--distpath",
            str(dist_path),
            "--workpath",
            str(build_path),
            "--specpath",
            str(spec_path),
            "--clean",
        ]

        for module in cfg.excluded_modules:
            command.extend(["--exclude-module", module])

        for module in cfg.hidden_imports:
            command.extend(["--hidden-import", module])

        for module in cfg.collect_all_modules:
            command.extend(["--collect-all", module])

        for source, dest in cfg.add_datas:
            if source.exists():
                command.extend(["--add-data", cls.add_data_argument(source, dest)])

        if cfg.icon_path and cfg.icon_path.exists():
            command.extend(["--icon", str(cfg.icon_path)])

        if version_info and version_info.exists():
            command.extend(["--version-file", str(version_info)])

        command.append(str(cfg.main_script))

        log("")
        log("Comando PyInstaller:")
        log(" ".join(f'\"{c}\"' if " " in c else c for c in command))
        log("")

        kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if os.name == "nt" else {}
        try:
            proc = subprocess.Popen(
                command,
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=str(cfg.project_dir),
                **kwargs,
            )
            assert proc.stdout is not None
            for line in iter(proc.stdout.readline, ""):
                if line:
                    log(line.rstrip())
            proc.wait()
            if proc.returncode != 0:
                return False, f"Errore PyInstaller. Codice: {proc.returncode}"

            exe_path = dist_path / cfg.app_name / f"{cfg.app_name}.exe"
            return True, str(exe_path)
        except Exception as exc:
            return False, f"Errore imprevisto durante PyInstaller: {exc}"

    @staticmethod
    def _iss_escape(value: object) -> str:
        """Escapa i doppi apici per i valori testuali di Inno Setup."""
        return str(value).replace('"', '""')

    @classmethod
    def create_installer(cls, cfg: SimpleBuildConfig, exe_path: str, log: Callable[[str], None]) -> Tuple[bool, str]:
        assert cfg.output_dir is not None

        app_folder = cfg.output_dir / "dist" / cfg.app_name
        exe_file = app_folder / f"{cfg.app_name}.exe"
        if not app_folder.exists() or not app_folder.is_dir():
            return False, f"Cartella della build non trovata: {app_folder}"
        if not exe_file.exists():
            return False, f"EXE della build non trovato: {exe_file}"

        # Nome dello script senza caratteri problematici. Lo script viene collocato
        # nella cartella BuildOutput e usa {#SourcePath}, evitando percorsi assoluti
        # fragili o non correttamente interpretati da Inno Setup.
        safe_stem = re.sub(r"[^A-Za-z0-9_.-]+", "_", cfg.app_name).strip("._") or "application"
        iss_path = cfg.output_dir / f"{safe_stem}_installer.iss"
        setup_name = f"{cfg.app_name}_Setup_v{cfg.app_version}"

        app_name = cls._iss_escape(cfg.app_name)
        app_version = cls._iss_escape(cfg.app_version)
        setup_filename = cls._iss_escape(setup_name)

        icon_line = ""
        if cfg.icon_path and cfg.icon_path.exists():
            # Inno Setup accetta il percorso assoluto tra virgolette. I doppi apici
            # interni vengono escapati per sicurezza.
            icon_path = cls._iss_escape(cfg.icon_path.resolve())
            icon_line = f'SetupIconFile="{icon_path}"'

        # Usa i percorsi assoluti gia verificati da Python. Inno Setup gestisce
        # correttamente gli spazi quando il valore Source e racchiuso tra virgolette.
        # Non usare {#SourcePath}: non e una variabile ISPP predefinita e verrebbe
        # interpretata letteralmente nel percorso.
        app_folder_iss = cls._iss_escape(app_folder.resolve())
        setup_output_dir = cfg.output_dir / "setup"
        setup_output_dir.mkdir(parents=True, exist_ok=True)
        setup_output_iss = cls._iss_escape(setup_output_dir.resolve())

        iss = rf"""
[Setup]
AppName={app_name}
AppVersion={app_version}
AppPublisher={cls._iss_escape(DEFAULT_COMPANY_NAME)}
DefaultDirName={{localappdata}}\Programs\{app_name}
DefaultGroupName={app_name}
OutputDir="{setup_output_iss}"
OutputBaseFilename={setup_filename}
Compression=lzma
SolidCompression=yes
WizardStyle=modern
PrivilegesRequired=lowest
ArchitecturesInstallIn64BitMode=x64compatible
{icon_line}
UninstallDisplayIcon={{app}}\{app_name}.exe

[Tasks]
Name: "desktopicon"; Description: "Crea un'icona sul desktop"; GroupDescription: "Icone aggiuntive:"; Flags: unchecked

[Files]
; Esclude artefatti di sviluppo che non servono a runtime e che possono avere
; percorsi molto lunghi. Le DLL, i plugin Qt, i file Python e le risorse normali
; rimangono inclusi.
Source: "{app_folder_iss}\*"; DestDir: "{{app}}"; Excludes: "*.obj,*.cpp.obj,*.pdb,*.lib,*.exp,*.qmlir,__pycache__\*,*.pyc"; Flags: ignoreversion recursesubdirs createallsubdirs

[Icons]
Name: "{{group}}\{app_name}"; Filename: "{{app}}\{app_name}.exe"
Name: "{{autodesktop}}\{app_name}"; Filename: "{{app}}\{app_name}.exe"; Tasks: desktopicon

[Run]
Filename: "{{app}}\{app_name}.exe"; Description: "Avvia {app_name}"; Flags: nowait postinstall skipifsilent
""".strip()

        try:
            iss_path.write_text(iss, encoding="utf-8-sig")
            log(f"Script Inno Setup creato: {iss_path}")
            log(f"Sorgente installer verificata: {app_folder}")
        except Exception as exc:
            return False, f"Errore scrittura file ISS: {exc}"

        iscc = cls.find_inno_compiler()
        if not iscc:
            return True, "Inno Setup non trovato: EXE creato, script .iss pronto ma setup non compilato."

        log(f"Avvio Inno Setup Compiler: {iscc}")
        kwargs = {"creationflags": subprocess.CREATE_NO_WINDOW} if os.name == "nt" else {}
        try:
            proc = subprocess.Popen(
                [iscc, str(iss_path.resolve())],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                encoding="utf-8",
                errors="replace",
                cwd=str(cfg.output_dir),
                **kwargs,
            )
            assert proc.stdout is not None
            output_lines: list[str] = []
            for line in iter(proc.stdout.readline, ""):
                if line:
                    clean = line.rstrip()
                    output_lines.append(clean)
                    log(clean)
            proc.wait()
            if proc.returncode != 0:
                details = "\n".join(output_lines[-8:])
                message = f"Errore Inno Setup. Codice: {proc.returncode}"
                if details:
                    message += f"\n\nDettagli finali:\n{details}"
                return False, message

            setup_path = cfg.output_dir / "setup" / f"{setup_name}.exe"
            return True, f"Setup creato: {setup_path}"
        except Exception as exc:
            return False, f"Errore Inno Setup: {exc}"



class BuildThread(QThread):
    log_signal = pyqtSignal(str)
    finished_signal = pyqtSignal(bool, str, str)

    def __init__(self, cfg: SimpleBuildConfig):
        super().__init__()
        self.cfg = cfg

    def log(self, text: str) -> None:
        self.log_signal.emit(text)

    def run(self) -> None:
        assert self.cfg.output_dir is not None
        self.log(f"Build {self.cfg.app_name} v{self.cfg.app_version}")
        self.log(f"Main: {self.cfg.main_script}")
        self.log(f"Progetto: {self.cfg.project_dir}")
        self.log(f"Output: {self.cfg.output_dir}")
        if self.cfg.icon_path:
            if self.cfg.icon_source_path and self.cfg.icon_source_path != self.cfg.icon_path:
                self.log(f"Icona sorgente: {self.cfg.icon_source_path}")
                self.log(f"Icona convertita: {self.cfg.icon_path}")
            else:
                self.log(f"Icona: {self.cfg.icon_path}")
        else:
            self.log("Icona: nessuna")
        self.log(f"Hidden imports: {', '.join(self.cfg.hidden_imports) if self.cfg.hidden_imports else 'nessuno'}")
        self.log(f"Collect-all (Qt escluso intenzionalmente): {', '.join(self.cfg.collect_all_modules) if self.cfg.collect_all_modules else 'nessuno'}")
        self.log(f"Moduli esclusi: {', '.join(self.cfg.excluded_modules) if self.cfg.excluded_modules else 'nessuno'}")
        self.log(f"Moduli richiesti: {', '.join(self.cfg.required_modules) if self.cfg.required_modules else 'nessuno'}")
        self.log(f"Add-data: {len(self.cfg.add_datas)} elemento/i")
        self.log("")

        self.log("=" * 60)
        self.log("1. CREAZIONE EXE CON PYINSTALLER")
        self.log("=" * 60)
        ok, result = BuilderService.run_pyinstaller(self.cfg, self.log)
        if not ok:
            self.finished_signal.emit(False, result, str(self.cfg.output_dir))
            return

        exe_path = result
        self.log("")
        self.log(f"EXE creato: {exe_path}")
        self.log("")
        self.log("=" * 60)
        self.log("2. CREAZIONE SETUP CON INNO SETUP")
        self.log("=" * 60)

        ok, installer_msg = BuilderService.create_installer(self.cfg, exe_path, self.log)
        self.log(installer_msg)
        if not ok:
            self.finished_signal.emit(False, installer_msg, str(self.cfg.output_dir))
            return

        self.finished_signal.emit(True, "Build completata.", str(self.cfg.output_dir))


class SimpleExeWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker: Optional[BuildThread] = None
        self.config: Optional[SimpleBuildConfig] = None
        self.manual_icon_source_path: Optional[Path] = None
        self.setWindowTitle(f"{APP_DISPLAY_NAME} v{APP_VERSION}")
        self.setMinimumSize(920, 620)
        self._build_ui()
        QTimer.singleShot(0, self._center)

    def _center(self) -> None:
        screen = self.screen() or QApplication.primaryScreen()
        if not screen:
            return
        area = screen.availableGeometry()
        w = min(1050, area.width() - 60)
        h = min(760, area.height() - 60)
        x = area.x() + max(20, (area.width() - w) // 2)
        y = area.y() + max(30, (area.height() - h) // 2)
        self.setGeometry(x, y, w, h)

    def _build_ui(self) -> None:
        self.setStyleSheet(
            """
            QMainWindow { background: #eef2f6; }
            QFrame#card { background: white; border: 1px solid #d7dee7; border-radius: 12px; }
            QLabel { color: #1f2a44; font-size: 13px; }
            QLabel#title { font-size: 25px; font-weight: 700; }
            QLabel#subtitle { color: #60708a; }
            QLineEdit { background: #fbfcfd; border: 1px solid #cfd7e3; border-radius: 8px; padding: 8px; }
            QPushButton { border-radius: 8px; padding: 10px 14px; font-weight: 600; border: 1px solid #c7d2df; background: white; }
            QPushButton:hover { background: #f3f7fb; }
            QPushButton#primary { background: #1565d8; color: white; border: none; }
            QPushButton#primary:hover { background: #0f57bd; }
            QPushButton#success { background: #198754; color: white; border: none; font-size: 14px; padding: 12px 18px; }
            QPushButton#success:hover { background: #157347; }
            QPushButton#danger { background: #dc3545; color: white; border: none; }
            QTextEdit#log { background: #121a26; color: #d8e1ea; border: 1px solid #253246; border-radius: 8px; font-family: Consolas, 'Courier New', monospace; }
            QProgressBar { border: 1px solid #cfd7e3; border-radius: 8px; text-align: center; background: white; }
            QProgressBar::chunk { border-radius: 7px; background: #1565d8; }
            """
        )
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(16, 16, 16, 16)
        root.setSpacing(12)

        header = QFrame()
        header.setObjectName("card")
        h = QVBoxLayout(header)
        h.setContentsMargins(18, 16, 18, 16)
        title = QLabel("Crea EXE / Setup universale")
        title.setObjectName("title")
        subtitle = QLabel("Seleziona il main.py di qualsiasi progetto Python. Puoi anche scegliere un immagine: verra convertita automaticamente in .ico.")
        subtitle.setObjectName("subtitle")
        subtitle.setWordWrap(True)
        h.addWidget(title)
        h.addWidget(subtitle)
        root.addWidget(header)

        card = QFrame()
        card.setObjectName("card")
        c = QVBoxLayout(card)
        c.setContentsMargins(18, 18, 18, 18)
        c.setSpacing(12)

        row = QHBoxLayout()
        self.main_entry = QLineEdit()
        self.main_entry.setReadOnly(True)
        select_btn = QPushButton("Seleziona main.py")
        select_btn.setObjectName("primary")
        select_btn.clicked.connect(self.select_main)
        row.addWidget(self.main_entry, 1)
        row.addWidget(select_btn)
        c.addLayout(row)

        meta_row = QHBoxLayout()
        self.app_name_entry = QLineEdit()
        self.app_name_entry.setPlaceholderText("Nome applicazione rilevato automaticamente")
        self.app_name_entry.editingFinished.connect(self.apply_manual_metadata)
        self.version_entry = QLineEdit()
        self.version_entry.setPlaceholderText("Versione rilevata automaticamente")
        self.version_entry.editingFinished.connect(self.apply_manual_metadata)
        meta_row.addWidget(QLabel("Nome app:"))
        meta_row.addWidget(self.app_name_entry, 2)
        meta_row.addWidget(QLabel("Versione:"))
        meta_row.addWidget(self.version_entry, 1)
        c.addLayout(meta_row)

        icon_row = QHBoxLayout()
        self.icon_entry = QLineEdit()
        self.icon_entry.setReadOnly(True)
        self.icon_entry.setPlaceholderText("Icona non selezionata: verra usata quella rilevata automaticamente, se presente")
        icon_btn = QPushButton("Seleziona immagine/icona")
        icon_btn.clicked.connect(self.select_icon)
        clear_icon_btn = QPushButton("Rimuovi icona")
        clear_icon_btn.clicked.connect(self.clear_icon)
        icon_row.addWidget(self.icon_entry, 1)
        icon_row.addWidget(icon_btn)
        icon_row.addWidget(clear_icon_btn)
        c.addLayout(icon_row)

        self.detected_label = QLabel("Nessun file selezionato.")
        self.detected_label.setWordWrap(True)
        c.addWidget(self.detected_label)
        root.addWidget(card)

        log_card = QFrame()
        log_card.setObjectName("card")
        l = QVBoxLayout(log_card)
        l.setContentsMargins(14, 14, 14, 14)
        l.setSpacing(10)
        log_title = QLabel("Log")
        log_title.setStyleSheet("font-size: 16px; font-weight: 700;")
        l.addWidget(log_title)
        self.log_output = QTextEdit()
        self.log_output.setObjectName("log")
        self.log_output.setReadOnly(True)
        l.addWidget(self.log_output, 1)
        self.progress = QProgressBar()
        self.progress.setRange(0, 0)
        self.progress.hide()
        l.addWidget(self.progress)
        root.addWidget(log_card, 1)

        footer = QHBoxLayout()
        self.status_label = QLabel("Pronto.")
        footer.addWidget(self.status_label, 1)
        exit_btn = QPushButton("Esci")
        exit_btn.setObjectName("danger")
        exit_btn.clicked.connect(self.close)
        self.build_btn = QPushButton("CREA EXE / SETUP")
        self.build_btn.setObjectName("success")
        self.build_btn.setEnabled(False)
        self.build_btn.clicked.connect(self.start_build)
        footer.addWidget(exit_btn)
        footer.addWidget(self.build_btn)
        root.addLayout(footer)

    def log(self, text: str) -> None:
        self.log_output.moveCursor(QTextCursor.MoveOperation.End)
        self.log_output.insertPlainText(text + "\n")
        self.log_output.moveCursor(QTextCursor.MoveOperation.End)

    def apply_manual_metadata(self) -> None:
        if not self.config:
            return

        name = BuildDetector.normalize_app_name(self.app_name_entry.text().strip() or self.config.app_name)
        version = (self.version_entry.text().strip() or self.config.app_version or "1.0.0").strip()
        old_name = self.config.app_name
        old_output = self.config.output_dir

        self.config.app_name = name
        self.config.app_version = version
        self.config.output_dir = self.config.project_dir / f"{name}_BuildOutput"

        # Se l'utente cambia nome app dopo avere convertito un'immagine, riconvertiamo
        # nel nuovo output per mantenere cartelle e nomi coerenti.
        if (
            self.manual_icon_source_path
            and self.manual_icon_source_path.exists()
            and self.config.output_dir != old_output
        ):
            ok, message, ico_path = IconConversionService.convert_to_ico(
                self.manual_icon_source_path,
                self.config.output_dir,
                self.config.app_name,
            )
            if ok and ico_path:
                self.config.icon_path = ico_path
                self.config.icon_source_path = self.manual_icon_source_path
                self.status_label.setText(message)

        if self.config.app_name != old_name or self.config.output_dir != old_output:
            self.status_label.setText("Nome/versione aggiornati. Output ricalcolato.")
        self._refresh_detected_summary()
        self._apply_window_icon()

    def _refresh_detected_summary(self) -> None:
        if not self.config:
            self.detected_label.setText("Nessun file selezionato.")
            self.build_btn.setEnabled(False)
            return

        if self.app_name_entry.text() != self.config.app_name:
            self.app_name_entry.setText(self.config.app_name)
        if self.version_entry.text() != self.config.app_version:
            self.version_entry.setText(self.config.app_version)

        datas = "<br>".join(f"{src} => {dst}" for src, dst in self.config.add_datas) or "nessuno"
        if self.config.icon_path:
            icon_text = str(self.config.icon_path)
            if self.config.icon_source_path and self.config.icon_source_path != self.config.icon_path:
                icon_text += f"<br><span style='color:#60708a;'>Origine: {self.config.icon_source_path}</span>"
        else:
            icon_text = "nessuna"

        entry_path = self.config.icon_source_path or self.config.icon_path
        self.icon_entry.setText("" if not entry_path else str(entry_path))
        self.detected_label.setText(
            f"<b>App:</b> {self.config.app_name}<br>"
            f"<b>Versione:</b> {self.config.app_version}<br>"
            f"<b>Cartella progetto:</b> {self.config.project_dir}<br>"
            f"<b>Icona .ico usata:</b> {icon_text}<br>"
            f"<b>Risorse incluse:</b><br>{datas}<br>"
            f"<b>Moduli richiesti:</b> {', '.join(self.config.required_modules)}<br>"
            f"<b>Collect-all:</b> {', '.join(self.config.collect_all_modules)}<br>"
            f"<b>Output:</b> {self.config.output_dir}"
        )
        self.build_btn.setEnabled(True)

    def _apply_window_icon(self) -> None:
        if self.config and self.config.icon_path and self.config.icon_path.exists():
            self.setWindowIcon(QIcon(str(self.config.icon_path)))
        else:
            self.setWindowIcon(QIcon())

    def _set_icon_from_image(self, image_path: Path, show_success: bool = True) -> bool:
        if not IconConversionService.is_supported(image_path):
            QMessageBox.warning(
                self,
                "Formato immagine non supportato",
                "Usa PNG, JPG, JPEG, BMP, GIF, WEBP, TIFF oppure ICO.",
            )
            return False

        if not self.config:
            self.manual_icon_source_path = image_path
            self.icon_entry.setText(str(image_path))
            if show_success:
                QMessageBox.information(
                    self,
                    "Immagine selezionata",
                    "Immagine selezionata. Dopo aver scelto main.py verra convertita automaticamente in .ico.",
                )
            return True

        assert self.config.output_dir is not None
        ok, message, ico_path = IconConversionService.convert_to_ico(
            image_path,
            self.config.output_dir,
            self.config.app_name,
        )
        if not ok or not ico_path:
            QMessageBox.warning(self, "Conversione icona non riuscita", message)
            return False

        self.manual_icon_source_path = image_path
        self.config.icon_path = ico_path
        self.config.icon_source_path = image_path
        self._refresh_detected_summary()
        self._apply_window_icon()
        self.status_label.setText(message)
        if show_success and image_path.suffix.lower() != ".ico":
            QMessageBox.information(self, "Icona convertita", message)
        return True

    def select_main(self) -> None:
        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleziona main.py principale",
            "",
            "Python (*.py);;Tutti i file (*.*)",
        )
        if not path:
            return
        main = Path(path)
        if main.suffix.lower() != ".py":
            QMessageBox.warning(self, "File non valido", "Seleziona un file Python .py.")
            return

        manual_icon = self.manual_icon_source_path
        self.config = BuildDetector.build_config(main)
        self.main_entry.setText(str(main))
        self.app_name_entry.setText(self.config.app_name)
        self.version_entry.setText(self.config.app_version)

        if manual_icon and manual_icon.exists():
            self._set_icon_from_image(manual_icon, show_success=False)
        else:
            self._refresh_detected_summary()
            self._apply_window_icon()

        self.status_label.setText("File selezionato. Puoi avviare la build o cambiare icona.")

    def select_icon(self) -> None:
        start_dir = ""
        if self.config:
            start_dir = str(self.config.project_dir)

        path, _ = QFileDialog.getOpenFileName(
            self,
            "Seleziona immagine o icona applicazione",
            start_dir,
            "Immagini supportate (*.ico *.png *.jpg *.jpeg *.bmp *.gif *.webp *.tif *.tiff);;"
            "Icone Windows (*.ico);;"
            "PNG (*.png);;"
            "JPEG (*.jpg *.jpeg);;"
            "Tutti i file (*.*)",
        )
        if not path:
            return

        image_path = Path(path)
        if not image_path.exists() or not image_path.is_file():
            QMessageBox.warning(self, "File non trovato", "Il file immagine selezionato non esiste.")
            return

        self._set_icon_from_image(image_path, show_success=True)

    def clear_icon(self) -> None:
        if self.config:
            self.config.icon_path = None
            self.config.icon_source_path = None
        self.manual_icon_source_path = None
        self.icon_entry.clear()
        self._refresh_detected_summary()
        self._apply_window_icon()
        self.status_label.setText("Icona rimossa. La build verra creata senza icona personalizzata.")

    def start_build(self) -> None:
        if not self.config:
            QMessageBox.warning(self, "Attenzione", "Seleziona prima il file main.py.")
            return
        if self.worker and self.worker.isRunning():
            QMessageBox.information(self, "Build in corso", "Attendi la fine della build corrente.")
            return
        if not self.config.main_script.exists():
            QMessageBox.warning(self, "Errore", "Il file main.py selezionato non esiste piu.")
            return

        self.apply_manual_metadata()
        if self.config.icon_path and not self.config.icon_path.exists():
            QMessageBox.warning(self, "Errore", "Il file icona selezionato non esiste piu.")
            return
        if self.config.icon_path and self.config.icon_path.suffix.lower() != ".ico":
            QMessageBox.warning(self, "Errore", "L icona finale deve essere un file .ico.")
            return

        self.log_output.clear()
        self.build_btn.setEnabled(False)
        self.progress.show()
        self.status_label.setText("Build in corso...")
        self.worker = BuildThread(self.config)
        self.worker.log_signal.connect(self.log)
        self.worker.finished_signal.connect(self.on_finished)
        self.worker.start()

    def on_finished(self, success: bool, message: str, output_dir: str) -> None:
        self.build_btn.setEnabled(True)
        self.progress.hide()
        self.log("")
        self.log("=" * 60)
        self.log(message)
        self.status_label.setText(message)
        if success:
            QMessageBox.information(self, "Completato", message)
            try:
                if os.name == "nt":
                    os.startfile(output_dir)  # type: ignore[attr-defined]
                elif sys.platform == "darwin":
                    subprocess.Popen(["open", output_dir])
                else:
                    subprocess.Popen(["xdg-open", output_dir])
            except Exception:
                pass
        else:
            QMessageBox.critical(self, "Errore", message)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setApplicationName(APP_DISPLAY_NAME)
    window = SimpleExeWindow()
    window.show()
    sys.exit(app.exec())
