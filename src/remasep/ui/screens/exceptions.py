"""Pantalla 'Excepciones': reclasificación visual (no persiste nada)."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from remasep.services.mock_remasep import ExceptionItem
from remasep.ui.components.status_card import make_card
from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.styles import format_int, refresh_style

_STEPS = ["Archivos", "Análisis", "Revisión", "Resultado"]
_PLACEHOLDER = "— Seleccionar —"


class ExceptionRow(QFrame):
    def __init__(self, item: ExceptionItem, on_change) -> None:
        super().__init__()
        self.setObjectName("card")
        self.item = item
        self._on_change = on_change

        layout = QVBoxLayout(self)
        layout.setContentsMargins(18, 14, 18, 14)
        layout.setSpacing(6)

        title = QLabel(item.value)
        title.setProperty("role", "h2")
        count = QLabel(f"{format_int(item.count)} registros")
        count.setProperty("role", "muted")
        reason = QLabel(f"Motivo: {item.reason}")
        reason.setProperty("role", "muted")
        reason.setWordWrap(True)

        self.combo = QComboBox()
        self.combo.addItem(_PLACEHOLDER, "")
        for category in item.possible_categories:
            self.combo.addItem(category, category)
        self.combo.currentIndexChanged.connect(lambda _i: self._on_change())

        self.remember = QCheckBox("Recordar esta clasificación para próximos meses")

        layout.addWidget(title)
        layout.addWidget(count)
        layout.addWidget(reason)
        layout.addSpacing(4)
        layout.addWidget(QLabel("Clasificar como:"))
        layout.addWidget(self.combo)
        layout.addWidget(self.remember)

    @property
    def decision(self) -> str:
        return self.combo.currentData() or ""

    @property
    def resolved(self) -> bool:
        return bool(self.decision)


class ExceptionsScreen(QWidget):
    name = "exceptions"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app
        self.rows: list[ExceptionRow] = []

        self.steps = StepIndicator(_STEPS)
        heading = QLabel("Revisión de excepciones")
        heading.setProperty("role", "h2")

        self.banner = QFrame()
        self.banner.setObjectName("banner")
        banner_layout = QVBoxLayout(self.banner)
        banner_layout.setContentsMargins(14, 10, 14, 10)
        self.banner_label = QLabel()
        self.banner_label.setWordWrap(True)
        banner_layout.addWidget(self.banner_label)

        self._list_card, self._list_layout = make_card(title="CLASIFICACIONES PENDIENTES")

        self.back_button = QPushButton("Atrás")
        self.back_button.clicked.connect(lambda: self._app.navigate("analysis"))
        self.continue_button = QPushButton("Continuar")
        self.continue_button.setProperty("variant", "primary")
        self.continue_button.clicked.connect(self._continue)

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addStretch()
        actions.addWidget(self.continue_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 28, 40, 24)
        layout.setSpacing(14)
        layout.addWidget(self.steps)
        layout.addWidget(heading)
        layout.addWidget(self.banner)
        layout.addWidget(self._list_card)
        layout.addStretch()
        layout.addLayout(actions)

    def on_enter(self) -> None:
        self.steps.set_current(2)
        result = self._app.state.analysis
        while self._list_layout.count() > 1:
            item = self._list_layout.takeAt(1)
            if item.widget():
                item.widget().deleteLater()
        self.rows = []
        if result is None:
            return
        for exception in result.exceptions:
            row = ExceptionRow(exception, self._refresh)
            saved = self._app.state.exception_decisions.get(exception.value)
            if saved:
                index = row.combo.findData(saved)
                if index >= 0:
                    row.combo.setCurrentIndex(index)
            self.rows.append(row)
            self._list_layout.addWidget(row)
        self._refresh()

    def _refresh(self) -> None:
        pending = [row for row in self.rows if not row.resolved]
        if pending:
            noun = "excepción" if len(pending) == 1 else "excepciones"
            self.banner.setObjectName("banner")
            self.banner_label.setText(
                f"{len(pending)} {noun} sin resolver. "
                "Una clasificación pendiente bloquea la generación del REMASEP."
            )
        else:
            self.banner.setObjectName("bannerOk")
            self.banner_label.setText("Todas las excepciones tienen una clasificación asignada.")
        refresh_style(self.banner)
        self.continue_button.setEnabled(not pending)

    def _continue(self) -> None:
        decisions = {row.item.value: row.decision for row in self.rows if row.resolved}
        self._app.state.exception_decisions = decisions
        self._app.navigate("summary")
