"""Selector de archivo reutilizable.

Abre un ``QFileDialog`` y muestra nombre + tamaño + estado. NO lee ni procesa el
contenido del archivo: solo registra la ruta elegida.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from remasep.ui.styles import refresh_style

_UNITS = ["B", "KB", "MB", "GB"]


def human_size(num_bytes: int) -> str:
    size = float(num_bytes)
    for unit in _UNITS:
        if size < 1024 or unit == _UNITS[-1]:
            if unit == "B":
                return f"{int(size)} {unit}"
            return f"{size:.1f} {unit}".replace(".", ",")
        size /= 1024
    return f"{num_bytes} B"


class FileSelector(QFrame):
    fileSelected = Signal(object)  # Path | None

    def __init__(
        self,
        title: str,
        *,
        extensions: list[str] | None = None,
        badge: str | None = None,
        hint: str | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("card")
        self._extensions = [e.lower().lstrip(".") for e in (extensions or ["xlsx", "xlsm", "csv"])]
        self._path: Path | None = None

        outer = QVBoxLayout(self)
        outer.setContentsMargins(16, 14, 16, 14)
        outer.setSpacing(6)

        header = QHBoxLayout()
        title_label = QLabel(title)
        title_label.setProperty("role", "h2")
        header.addWidget(title_label)
        header.addStretch()
        if badge:
            pill = QLabel(badge)
            pill.setProperty("pill", "pending")
            header.addWidget(pill)
        outer.addLayout(header)

        if hint:
            hint_label = QLabel(hint)
            hint_label.setProperty("role", "muted")
            hint_label.setWordWrap(True)
            outer.addWidget(hint_label)

        row = QHBoxLayout()
        self._button = QPushButton("Seleccionar archivo")
        self._button.clicked.connect(self._open_dialog)
        self._name_label = QLabel("Sin archivo seleccionado")
        self._name_label.setProperty("role", "muted")
        row.addWidget(self._button)
        row.addWidget(self._name_label, stretch=1)
        outer.addLayout(row)

        self._status_label = QLabel("")
        self._status_label.setProperty("role", "muted")
        outer.addWidget(self._status_label)

    # --- API pública -----------------------------------------------------

    @property
    def path(self) -> Path | None:
        return self._path

    def set_selection(self, path: Path | None) -> None:
        """Registra una selección (usado tras el diálogo y desde los tests)."""
        self._path = Path(path) if path else None
        if self._path is None:
            self._name_label.setText("Sin archivo seleccionado")
            self._status_label.setText("")
        else:
            size = self._path.stat().st_size if self._path.exists() else 0
            self._name_label.setText(self._path.name)
            self._status_label.setText(f"{human_size(size)} · seleccionado")
            self._button.setText("Reemplazar archivo")
        for widget in (self._name_label, self._status_label):
            refresh_style(widget)
        self.fileSelected.emit(self._path)

    def clear(self) -> None:
        self.set_selection(None)
        self._button.setText("Seleccionar archivo")

    # --- interno -------------------------------------------------------

    def _filter(self) -> str:
        patterns = " ".join(f"*.{ext}" for ext in self._extensions)
        return f"Fuentes ({patterns})"

    def _open_dialog(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo", "", self._filter())
        if selected:
            self.set_selection(Path(selected))
