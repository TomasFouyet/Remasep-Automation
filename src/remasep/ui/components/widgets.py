"""Componentes reutilizables de la aplicación (Sprint 3.10).

Pocos, consistentes: `MetricCard`, `StatusBadge`, `SectionHeader`, `EmptyState`,
`ErrorBanner`, `Card`. Toman su estilo del QSS central (`styles.py`); no definen
QSS propio duplicado.
"""

from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSizePolicy,
    QVBoxLayout,
    QWidget,
)

from remasep.ui.styles import format_int, refresh_style


class Card(QFrame):
    """Contenedor blanco discreto con padding y radio consistentes."""

    def __init__(self, *, muted: bool = False, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("cardMuted" if muted else "card")
        self.body = QVBoxLayout(self)
        self.body.setContentsMargins(20, 18, 20, 18)
        self.body.setSpacing(10)


class SectionHeader(QWidget):
    def __init__(self, text: str, *, hint: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(2)
        label = QLabel(text.upper())
        label.setProperty("role", "section")
        layout.addWidget(label)
        if hint:
            sub = QLabel(hint)
            sub.setProperty("role", "faint")
            sub.setWordWrap(True)
            layout.addWidget(sub)


class StatusBadge(QLabel):
    """Etiqueta compacta de estado: ok / warning / pending / info."""

    def __init__(self, text: str, *, kind: str = "info", parent: QWidget | None = None) -> None:
        super().__init__(text, parent)
        self.setProperty("badge", kind)
        self.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.setSizePolicy(QSizePolicy.Policy.Maximum, QSizePolicy.Policy.Maximum)

    def set_state(self, text: str, kind: str) -> None:
        self.setText(text)
        self.setProperty("badge", kind)
        refresh_style(self)


class MetricCard(Card):
    """KPI: valor grande + etiqueta + pista opcional."""

    def __init__(
        self,
        label: str,
        *,
        hint: str = "",
        accent: bool = False,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent=parent)
        self.body.setSpacing(4)
        self._value = QLabel("—")
        self._value.setProperty("role", "kpi-value-accent" if accent else "kpi-value")
        self._label = QLabel(label)
        self._label.setProperty("role", "kpi-label")
        self._label.setWordWrap(True)
        self._hint = QLabel(hint)
        self._hint.setProperty("role", "kpi-hint")
        self._hint.setWordWrap(True)
        self._hint.setVisible(bool(hint))
        self.body.addWidget(self._value)
        self.body.addWidget(self._label)
        self.body.addWidget(self._hint)

    def set_value(self, value: int | str) -> None:
        self._value.setText(format_int(value) if isinstance(value, int) else str(value))

    def set_hint(self, text: str) -> None:
        self._hint.setText(text)
        self._hint.setVisible(bool(text))


class EmptyState(QWidget):
    def __init__(self, title: str, *, detail: str = "", parent: QWidget | None = None) -> None:
        super().__init__(parent)
        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 32, 0, 32)
        layout.setSpacing(6)
        layout.setAlignment(Qt.AlignmentFlag.AlignHCenter)
        t = QLabel(title)
        t.setProperty("role", "h3")
        t.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(t)
        if detail:
            d = QLabel(detail)
            d.setProperty("role", "muted")
            d.setAlignment(Qt.AlignmentFlag.AlignCenter)
            d.setWordWrap(True)
            layout.addWidget(d)


class ErrorBanner(QFrame):
    """Aviso de error humano con acción de reintento opcional."""

    actionClicked = Signal()

    def __init__(self, parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self.setObjectName("bannerError")
        layout = QHBoxLayout(self)
        layout.setContentsMargins(16, 12, 12, 12)
        layout.setSpacing(12)

        text_box = QVBoxLayout()
        text_box.setSpacing(2)
        self._title = QLabel()
        self._title.setProperty("role", "h3")
        self._title.setWordWrap(True)
        self._detail = QLabel()
        self._detail.setProperty("role", "muted")
        self._detail.setWordWrap(True)
        text_box.addWidget(self._title)
        text_box.addWidget(self._detail)
        layout.addLayout(text_box, stretch=1)

        self._action = QPushButton()
        self._action.clicked.connect(self.actionClicked.emit)
        layout.addWidget(self._action, alignment=Qt.AlignmentFlag.AlignTop)
        self.setVisible(False)

    def show_error(self, title: str, detail: str = "", action_label: str = "") -> None:
        self._title.setText(title)
        self._detail.setText(detail)
        self._detail.setVisible(bool(detail))
        self._action.setText(action_label or "Volver")
        self._action.setVisible(bool(action_label))
        self.setVisible(True)

    def clear(self) -> None:
        self.setVisible(False)
