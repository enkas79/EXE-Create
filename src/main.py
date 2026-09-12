"""
ToExe Universal Simple - Generatore EXE/Setup semplificato.

Punto di ingresso: avvia solo la QApplication e la finestra principale.
Logica di build in core/builder_core.py, autoupdate in updater/updater.py,
GUI in gui/main_window.py (separazione richiesta da CLAUDE.md).

Avvio: python src/main.py
"""

from __future__ import annotations

import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent))

from PyQt6.QtWidgets import QApplication

from gui.main_window import APP_DISPLAY_NAME, SimpleExeWindow


def main() -> int:
    app = QApplication(sys.argv)
    app.setApplicationName(APP_DISPLAY_NAME)
    window = SimpleExeWindow()
    window.show()
    return app.exec()


if __name__ == "__main__":
    sys.exit(main())
