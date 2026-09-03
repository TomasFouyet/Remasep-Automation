"""Pantalla 'Análisis': resultados del MockRemasepService."""

from __future__ import annotations

from PySide6.QtWidgets import QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from remasep.ui.components.status_card import StatusRow, make_card
from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.styles import format_int

_STEPS = ["Archivos", "Análisis", "Revisión", "Resultado"]


class AnalysisScreen(QWidget):
    name = "analysis"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        self.steps = StepIndicator(_STEPS)
        self.period_label = QLabel()
        self.period_label.setProperty("role", "h2")
        self.count_label = QLabel()
        self.count_label.setProperty("role", "metric")

        self._validations_card, self._validations_layout = make_card(title="VALIDACIONES")
        self._result_card, self._result_layout = make_card(title="RESULTADO")
        self._modules_card, self._modules_layout = make_card(title="MÓDULOS DETECTADOS")

        self.back_button = QPushButton("Atrás")
        self.back_button.clicked.connect(lambda: self._app.navigate("new_report"))
        self.review_button = QPushButton("Revisar problemas")
        self.review_button.setProperty("variant", "primary")
        self.review_button.clicked.connect(lambda: self._app.navigate("exceptions"))

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addStretch()
        actions.addWidget(self.review_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 28, 40, 24)
        layout.setSpacing(14)
        layout.addWidget(self.steps)
        layout.addWidget(self.period_label)
        layout.addWidget(self.count_label)
        layout.addWidget(self._validations_card)
        layout.addWidget(self._result_card)
        layout.addWidget(self._modules_card)
        layout.addStretch()
        layout.addLayout(actions)

    def on_enter(self) -> None:
        self.steps.set_current(1)
        result = self._app.state.analysis
        if result is None:
            return

        self.period_label.setText(result.period_label)
        self.count_label.setText(f"{format_int(result.total_records)} atenciones encontradas")

        _clear(self._validations_layout, keep=1)
        for validation in result.validations:
            self._validations_layout.addWidget(
                StatusRow(validation.name, status=validation.status)
            )

        _clear(self._result_layout, keep=1)
        self._result_layout.addWidget(
            StatusRow("Clasificadas", status="ok", value=format_int(result.classified_records))
        )
        self._result_layout.addWidget(
            StatusRow(
                "Ignoradas justificadamente",
                status="neutral",
                value=format_int(result.ignored_records),
            )
        )
        self._result_layout.addWidget(
            StatusRow(
                "Registros sin clasificar",
                status="warning" if result.unclassified_records else "ok",
                value=format_int(result.unclassified_records),
            )
        )
        self._result_layout.addWidget(
            StatusRow(
                "Tipos que requieren revisión",
                status="warning" if result.review_groups else "ok",
                value=format_int(result.review_groups),
            )
        )

        _clear(self._modules_layout, keep=1)
        for module in result.modules:
            self._modules_layout.addWidget(
                StatusRow(module.name, status="ok" if module.detected else "pending")
            )
        other_title = QLabel("Otras fuentes")
        other_title.setProperty("role", "muted")
        self._modules_layout.addWidget(other_title)
        for source in result.other_sources:
            self._modules_layout.addWidget(
                StatusRow(source.name, status="pending", value=source.note)
            )


def _clear(layout: QVBoxLayout, *, keep: int) -> None:
    """Elimina los widgets del layout salvo los primeros ``keep`` (p.ej. el título)."""
    while layout.count() > keep:
        item = layout.takeAt(keep)
        widget = item.widget()
        if widget is not None:
            widget.deleteLater()
