"""Pantalla de inicio."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel, QPushButton, QVBoxLayout, QWidget

from remasep.ui.components.widgets import Card
from remasep.ui.open_location import open_containing_folder

_OUTPUTS_DIR = Path("outputs")


def _latest_report() -> Path | None:
    if not _OUTPUTS_DIR.is_dir():
        return None
    files = sorted(
        (p for p in _OUTPUTS_DIR.glob("REMASEP_*.xlsm") if p.is_file()),
        key=lambda p: p.stat().st_mtime,
        reverse=True,
    )
    return files[0] if files else None


class HomeScreen(QWidget):
    name = "home"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        title = QLabel("REMASEP")
        title.setProperty("role", "h1")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        subtitle = QLabel("Generación mensual de informe estadístico")
        subtitle.setProperty("role", "subtitle")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.start_button = QPushButton("Nuevo informe")
        self.start_button.setProperty("variant", "primary")
        self.start_button.setMinimumWidth(240)
        self.start_button.clicked.connect(lambda: self._app.navigate("new_report"))

        self._last_card = Card(muted=True)
        self._last_title = QLabel("Último informe generado")
        self._last_title.setProperty("role", "h3")
        self._last_name = QLabel()
        self._last_name.setProperty("role", "muted")
        self._last_open = QPushButton("Abrir carpeta")
        self._last_open.setProperty("variant", "ghost")
        self._last_open.clicked.connect(self._open_last_folder)
        self._last_card.body.addWidget(self._last_title)
        self._last_card.body.addWidget(self._last_name)
        self._last_card.body.addWidget(self._last_open, alignment=Qt.AlignmentFlag.AlignLeft)
        self._last_card.setMaximumWidth(420)
        self._last_card.setVisible(False)

        version = QLabel("Fundación Gantz")
        version.setProperty("role", "faint")
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(56, 56, 56, 28)
        layout.addStretch()
        layout.addWidget(title)
        layout.addSpacing(10)
        layout.addWidget(subtitle)
        layout.addSpacing(34)
        layout.addWidget(self.start_button, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addSpacing(26)
        layout.addWidget(self._last_card, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()
        layout.addWidget(version)

    def on_enter(self) -> None:
        latest = _latest_report()
        self._last_report_path = latest
        if latest is None:
            self._last_card.setVisible(False)
            return
        self._last_name.setText(latest.name)
        self._last_card.setVisible(True)

    def _open_last_folder(self) -> None:
        if getattr(self, "_last_report_path", None):
            open_containing_folder(self._last_report_path)
