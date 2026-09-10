"""Selector de archivo reutilizable.

Abre un ``QFileDialog`` y muestra el **nombre** del archivo y un estado ✓/aviso.
No expone la ruta completa salvo tooltip. No lee ni procesa el contenido.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Signal
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QWidget,
)

from remasep.ui.components.widgets import Card
from remasep.ui.styles import refresh_style


class FileSelector(Card):
    fileSelected = Signal(object)  # Path | None

    def __init__(
        self,
        title: str,
        *,
        extensions: list[str] | None = None,
        hint: str | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent=parent)
        self.body.setSpacing(8)
        self._extensions = [e.lower().lstrip(".") for e in (extensions or ["xlsx", "xlsm", "csv"])]
        self._path: Path | None = None

        title_label = QLabel(title)
        title_label.setProperty("role", "h3")
        self.body.addWidget(title_label)

        if hint:
            hint_label = QLabel(hint)
            hint_label.setProperty("role", "muted")
            hint_label.setWordWrap(True)
            self.body.addWidget(hint_label)

        row = QHBoxLayout()
        row.setSpacing(10)
        self._button = QPushButton("Seleccionar archivo")
        self._button.clicked.connect(self._open_dialog)
        self._name_label = QLabel("Ningún archivo seleccionado")
        self._name_label.setProperty("role", "muted")
        row.addWidget(self._button)
        row.addWidget(self._name_label, stretch=1)
        self.body.addLayout(row)

        self._status_label = QLabel("")
        self._status_label.setProperty("role", "muted")
        self._status_label.setWordWrap(True)
        self._status_label.setVisible(False)
        self.body.addWidget(self._status_label)

    # --- API pública -----------------------------------------------------

    @property
    def path(self) -> Path | None:
        return self._path

    @property
    def has_valid_extension(self) -> bool:
        return self._path is not None and self._path.suffix.lower().lstrip(".") in self._extensions

    def set_selection(self, path: Path | None) -> None:
        self._path = Path(path) if path else None
        if self._path is None:
            self._name_label.setText("Ningún archivo seleccionado")
            self._name_label.setToolTip("")
            self._status_label.setVisible(False)
            self._button.setText("Seleccionar archivo")
        else:
            self._name_label.setText(self._path.name)
            self._name_label.setToolTip(str(self._path))
            self._button.setText("Cambiar archivo")
            if self.has_valid_extension:
                self.set_status("✓  Archivo seleccionado", status="ok")
            else:
                self.set_status(
                    "El archivo no tiene una extensión válida "
                    f"({', '.join('.' + e for e in self._extensions)}).",
                    status="warning",
                )
        refresh_style(self._name_label)
        self.fileSelected.emit(self._path)

    def set_status(self, text: str, *, status: str = "neutral") -> None:
        self._status_label.setText(text)
        self._status_label.setProperty("status", status)
        self._status_label.setVisible(bool(text))
        refresh_style(self._status_label)

    def clear(self) -> None:
        self.set_selection(None)

    # --- interno -------------------------------------------------------

    def _filter(self) -> str:
        patterns = " ".join(f"*.{ext}" for ext in self._extensions)
        return f"Archivos ({patterns})"

    def _open_dialog(self) -> None:
        selected, _ = QFileDialog.getOpenFileName(self, "Seleccionar archivo", "", self._filter())
        if selected:
            self.set_selection(Path(selected))
