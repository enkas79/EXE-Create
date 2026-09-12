"""
Finestra principale (solo GUI): layout, widget, segnali/slot. La logica di
build vive in src/core/builder_core.py, l'autoupdate in src/updater/updater.py.
"""

from __future__ import annotations

import os
import subprocess
import sys
import webbrowser
from pathlib import Path
from typing import Optional

from PyQt6.QtCore import QTimer
from PyQt6.QtGui import QAction, QIcon, QTextCursor
from PyQt6.QtWidgets import (
    QApplication,
    QDialog,
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

from core.builder_core import BuildDetector, BuildThread, IconConversionService, SimpleBuildConfig
from updater.updater import UpdateCheckThread, UpdateInfo, read_local_version

APP_DISPLAY_NAME = "ToExe Universal Simple"
DEFAULT_COMPANY_NAME = "Enrico Martini"
VERSION_FILE_NAME = "version.txt"

GUIDE_TEXT = (
    "Guida rapida\n\n"
    "1. Seleziona il file main.py del progetto da trasformare in EXE.\n"
    "2. Facoltativo: seleziona un'immagine per l'icona, verra convertita automaticamente in .ico.\n"
    "3. Controlla nome applicazione e versione rilevati (puoi modificarli manualmente).\n"
    "4. Premi 'CREA EXE / SETUP' per avviare la build PyInstaller (onedir).\n"
    "5. Se Inno Setup e installato, verra generato anche il setup .exe.\n\n"
    "Nota: avvia ToExe dallo stesso ambiente Python del progetto da convertire."
)


def _project_root() -> Path:
    return Path(__file__).resolve().parents[2]


def read_app_version() -> str:
    return read_local_version(_project_root() / VERSION_FILE_NAME)


class SimpleExeWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.worker: Optional[BuildThread] = None
        self.update_check_worker: Optional[UpdateCheckThread] = None
        self.config: Optional[SimpleBuildConfig] = None
        self.manual_icon_source_path: Optional[Path] = None
        self.app_version = read_app_version()
        self.setWindowTitle(f"{APP_DISPLAY_NAME} v{self.app_version}")
        self.setMinimumSize(920, 620)
        self._build_menu_bar()
        self._build_ui()
        QTimer.singleShot(0, self._center)
        QTimer.singleShot(1500, lambda: self.check_for_updates(silent=True))

    # ------------------------------------------------------------------
    # Menu bar (obbligatoria: Aiuto -> Informazioni, Controlla aggiornamenti, Guida)
    # ------------------------------------------------------------------
    def _build_menu_bar(self) -> None:
        menu_bar = self.menuBar()

        file_menu = menu_bar.addMenu("&File")
        exit_action = QAction("Esci", self)
        exit_action.triggered.connect(self.close)
        file_menu.addAction(exit_action)

        help_menu = menu_bar.addMenu("&Aiuto")

        about_action = QAction("Informazioni", self)
        about_action.triggered.connect(self.show_about_dialog)
        help_menu.addAction(about_action)

        check_updates_action = QAction("Controlla aggiornamenti", self)
        check_updates_action.triggered.connect(lambda: self.check_for_updates(silent=False))
        help_menu.addAction(check_updates_action)

        guide_action = QAction("Guida", self)
        guide_action.triggered.connect(self.show_guide_dialog)
        help_menu.addAction(guide_action)

    def show_about_dialog(self) -> None:
        QMessageBox.about(
            self,
            f"Informazioni su {APP_DISPLAY_NAME}",
            f"<b>{APP_DISPLAY_NAME}</b><br>"
            f"Versione: {self.app_version}<br>"
            f"Autore: {DEFAULT_COMPANY_NAME}",
        )

    def show_guide_dialog(self) -> None:
        dialog = QDialog(self)
        dialog.setWindowTitle("Guida")
        dialog.setMinimumSize(480, 360)
        layout = QVBoxLayout(dialog)
        text = QTextEdit()
        text.setReadOnly(True)
        text.setPlainText(GUIDE_TEXT)
        layout.addWidget(text)
        close_btn = QPushButton("Chiudi")
        close_btn.clicked.connect(dialog.close)
        layout.addWidget(close_btn)
        dialog.exec()

    def check_for_updates(self, silent: bool) -> None:
        if self.update_check_worker and self.update_check_worker.isRunning():
            return
        self._update_check_silent = silent
        self.update_check_worker = UpdateCheckThread(self.app_version)
        self.update_check_worker.update_available.connect(self._on_update_available)
        self.update_check_worker.up_to_date.connect(self._on_up_to_date)
        self.update_check_worker.check_failed.connect(self._on_update_check_failed)
        self.update_check_worker.start()

    def _on_update_available(self, info: UpdateInfo) -> None:
        answer = QMessageBox.question(
            self,
            "Aggiornamento disponibile",
            f"E' disponibile la versione {info.version} (attuale: {self.app_version}).\n\n"
            f"Note di rilascio:\n{info.changelog}\n\n"
            "Vuoi aprire la pagina di download?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if answer == QMessageBox.StandardButton.Yes:
            url = info.download_url or info.html_url
            if url:
                webbrowser.open(url)

    def _on_up_to_date(self, current_version: str) -> None:
        if not self._update_check_silent:
            QMessageBox.information(self, "Nessun aggiornamento", f"Stai gia usando l'ultima versione ({current_version}).")

    def _on_update_check_failed(self, message: str) -> None:
        if not self._update_check_silent:
            QMessageBox.warning(self, "Verifica aggiornamenti non riuscita", message)

    # ------------------------------------------------------------------
    # Layout principale
    # ------------------------------------------------------------------
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
