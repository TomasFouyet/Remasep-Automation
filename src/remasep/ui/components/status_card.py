"""Fila/tarjeta de estado reutilizable (validaciones, módulos, datos del resumen)."""

from __future__ import annotations

from PySide6.QtWidgets import QFrame, QHBoxLayout, QLabel, QVBoxLayout, QWidget

from remasep.ui.styles import refresh_style


def make_card(*, title: str | None = None) -> tuple[QFrame, QVBoxLayout]:
    """Devuelve una tarjeta discreta y su layout interno para ir agregando filas."""
    card = QFrame()
    card.setObjectName("card")
    layout = QVBoxLayout(card)
    layout.setContentsMargins(18, 16, 18, 16)
    layout.setSpacing(6)
    if title:
        heading = QLabel(title)
        heading.setProperty("role", "section")
        layout.addWidget(heading)
    return card, layout

_GLYPHS = {
    "ok": "✓",
    "warning": "!",
    "error": "✕",
    "pending": "○",
    "neutral": "•",
}


class StatusRow(QWidget):
    """Un ítem: glifo de estado + título + valor opcional a la derecha."""

    def __init__(
        self,
        title: str,
        *,
        status: str = "neutral",
        value: str | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self._icon = QLabel()
        self._title = QLabel(title)
        self._value = QLabel(value or "")
        self._value.setProperty("role", "muted" if value is None else None)

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 4, 0, 4)
        layout.setSpacing(10)
        layout.addWidget(self._icon)
        layout.addWidget(self._title, stretch=1)
        layout.addWidget(self._value)

        self.set_status(status)

    def set_status(self, status: str) -> None:
        self._status = status
        self._icon.setText(_GLYPHS.get(status, "•"))
        self._icon.setProperty("status", status)
        self._title.setProperty("status", status if status in {"error"} else None)
        refresh_style(self._icon)
        refresh_style(self._title)

    def set_value(self, text: str) -> None:
        self._value.setText(text)

    @property
    def status(self) -> str:
        return self._status

    @property
    def title(self) -> str:
        return self._title.text()

    @property
    def value(self) -> str:
        return self._value.text()
