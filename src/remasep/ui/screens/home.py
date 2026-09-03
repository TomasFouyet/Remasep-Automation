"""Pantalla de inicio."""

from __future__ import annotations

from PySide6.QtCore import Qt
from PySide6.QtWidgets import QLabel, QPushButton, QVBoxLayout, QWidget

from remasep.services.mock_remasep import APP_VERSION


class HomeScreen(QWidget):
    name = "home"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        title = QLabel("REMASEP Automático")
        title.setProperty("role", "h1")
        title.setAlignment(Qt.AlignmentFlag.AlignCenter)

        subtitle = QLabel("Fundación Gantz")
        subtitle.setProperty("role", "subtitle")
        subtitle.setAlignment(Qt.AlignmentFlag.AlignCenter)

        blurb = QLabel("Generación mensual de informes REMASEP")
        blurb.setProperty("role", "muted")
        blurb.setAlignment(Qt.AlignmentFlag.AlignCenter)

        self.start_button = QPushButton("Nuevo reporte mensual")
        self.start_button.setProperty("variant", "primary")
        self.start_button.setMinimumWidth(260)
        self.start_button.clicked.connect(lambda: self._app.navigate("new_report"))

        version = QLabel(f"Versión {APP_VERSION}")
        version.setProperty("role", "muted")
        version.setAlignment(Qt.AlignmentFlag.AlignCenter)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(48, 48, 48, 24)
        layout.addStretch()
        layout.addWidget(title)
        layout.addSpacing(8)
        layout.addWidget(subtitle)
        layout.addSpacing(20)
        layout.addWidget(blurb)
        layout.addSpacing(28)
        layout.addWidget(self.start_button, alignment=Qt.AlignmentFlag.AlignCenter)
        layout.addStretch()
        layout.addWidget(version)

    def on_enter(self) -> None:
        """Hook de navegación: nada que refrescar en Home."""
