"""Pantalla 'Resumen': estado final de la maqueta (sin generación real)."""

from __future__ import annotations

from PySide6.QtWidgets import QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from remasep.services.mock_remasep import TEMPLATE_VERSION
from remasep.ui.components.status_card import StatusRow, make_card
from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.styles import format_int

_STEPS = ["Archivos", "Análisis", "Revisión", "Resultado"]
_GENERATE_TOOLTIP = (
    "La generación oficial estará disponible cuando se integre el motor de "
    "procesamiento y Microsoft Excel."
)


class SummaryScreen(QWidget):
    name = "summary"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        self.steps = StepIndicator(_STEPS)
        self.title_label = QLabel()
        self.title_label.setProperty("role", "h1")

        self._analysis_card, self._analysis_layout = make_card(title="RESULTADO DEL ANÁLISIS")
        self._review_card, self._review_layout = make_card(title="ESTADO TRAS LA REVISIÓN")
        self._sources_card, self._sources_layout = make_card(title="FUENTES")
        self._modules_card, self._modules_layout = make_card(title="MÓDULOS")
        self._template_card, self._template_layout = make_card(title="PLANTILLA OFICIAL")

        self.generate_button = QPushButton("Generar REMASEP")
        self.generate_button.setProperty("variant", "primary")
        self.generate_button.setEnabled(False)
        self.generate_button.setToolTip(_GENERATE_TOOLTIP)

        self.back_button = QPushButton("Volver")
        self.back_button.clicked.connect(lambda: self._app.navigate("exceptions"))
        self.home_button = QPushButton("Volver al inicio")
        self.home_button.clicked.connect(self._app.reset_flow)

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addWidget(self.home_button)
        actions.addStretch()
        actions.addWidget(self.generate_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 28, 40, 24)
        layout.setSpacing(14)
        layout.addWidget(self.steps)
        layout.addWidget(self.title_label)
        layout.addWidget(self._analysis_card)
        layout.addWidget(self._review_card)
        layout.addWidget(self._sources_card)
        layout.addWidget(self._modules_card)
        layout.addWidget(self._template_card)
        layout.addStretch()
        layout.addLayout(actions)

    def on_enter(self) -> None:
        self.steps.set_current(3)
        state = self._app.state
        result = state.analysis
        outcome = state.review_outcome()
        if result is None or outcome is None:
            return

        self.title_label.setText(f"REMASEP — {result.period_label.upper()}")

        # 1) Resultado original del análisis (no cambia con las decisiones).
        _clear(self._analysis_layout, keep=1)
        self._analysis_layout.addWidget(
            StatusRow("Atenciones", status="neutral", value=format_int(result.total_records))
        )
        self._analysis_layout.addWidget(
            StatusRow("Clasificadas", status="ok", value=format_int(result.classified_records))
        )
        self._analysis_layout.addWidget(
            StatusRow("Ignoradas", status="neutral", value=format_int(result.ignored_records))
        )
        self._analysis_layout.addWidget(
            StatusRow(
                "Sin clasificar",
                status="warning" if result.unclassified_records else "ok",
                value=format_int(result.unclassified_records),
            )
        )

        # 2) Estado tras aplicar las decisiones de la sesión.
        _clear(self._review_layout, keep=1)
        self._review_layout.addWidget(
            StatusRow("Clasificadas", status="ok", value=format_int(outcome.final_classified))
        )
        self._review_layout.addWidget(
            StatusRow("Ignoradas", status="neutral", value=format_int(outcome.final_ignored))
        )
        self._review_layout.addWidget(
            StatusRow(
                "Sin clasificar",
                status="warning" if outcome.final_unclassified else "ok",
                value=format_int(outcome.final_unclassified),
            )
        )
        self._review_layout.addWidget(
            StatusRow(
                "Grupos pendientes de revisión",
                status="warning" if outcome.pending_groups else "ok",
                value=format_int(outcome.pending_groups),
            )
        )

        _clear(self._sources_layout, keep=1)
        self._sources_layout.addWidget(StatusRow("Medinet", status="ok"))
        for source in result.other_sources:
            self._sources_layout.addWidget(
                StatusRow(source.name, status="pending", value=f"— {source.note}")
            )

        _clear(self._modules_layout, keep=1)
        for module in result.modules:
            self._modules_layout.addWidget(
                StatusRow(module.name, status="ok" if module.detected else "pending")
            )

        _clear(self._template_layout, keep=1)
        self._template_layout.addWidget(
            StatusRow("Versión", status="neutral", value=result.template_version or TEMPLATE_VERSION)
        )
        self._template_layout.addWidget(
            StatusRow(
                "CONTROL", status="pending", value="Pendiente integración con Microsoft Excel"
            )
        )


def _clear(layout: QVBoxLayout, *, keep: int) -> None:
    while layout.count() > keep:
        item = layout.takeAt(keep)
        if item.widget() is not None:
            item.widget().deleteLater()
