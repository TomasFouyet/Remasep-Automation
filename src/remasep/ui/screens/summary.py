"""Pantalla 'Resumen'.

- **DEMO**: resultado del análisis + estado tras la revisión (``ReviewOutcome``).
- **REAL**: resumen diagnóstico del análisis Medinet. NO usa ``ReviewOutcome``.

En ambos modos "Generar REMASEP" está deshabilitado.
"""

from __future__ import annotations

from PySide6.QtWidgets import (
    QFrame,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from remasep.services.mock_remasep import TEMPLATE_VERSION
from remasep.ui.components.status_card import StatusRow, make_card
from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.styles import format_int, refresh_style

_STEPS = ["Archivos", "Análisis", "Revisión", "Resultado"]
_GENERATE_TOOLTIP = (
    "La generación oficial estará disponible cuando se integre el motor de "
    "procesamiento y Microsoft Excel."
)
_LEGACY_UI_LABELS = {
    "AG": "Consultas médicas",
    "AH": "Controles odontológicos",
    "AI": "Evaluaciones odontológicas",
    "AJ": "Controles ortodoncia",
    "AK": "Controles ortopedia",
    "AL": "Instalaciones ortopedia",
}


def _clear(layout: QVBoxLayout, *, keep: int) -> None:
    while layout.count() > keep:
        item = layout.takeAt(keep)
        if item.widget() is not None:
            item.widget().deleteLater()


class SummaryScreen(QWidget):
    name = "summary"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        self.steps = StepIndicator(_STEPS)
        self.title_label = QLabel()
        self.title_label.setProperty("role", "h1")

        self.real_banner = QFrame()
        self.real_banner.setObjectName("bannerOk")
        banner_layout = QVBoxLayout(self.real_banner)
        banner_layout.setContentsMargins(14, 10, 14, 10)
        self._banner_line1 = QLabel("Análisis Medinet completado")
        self._banner_line2 = QLabel("Generación REMASEP todavía no disponible.")
        self._banner_line2.setProperty("role", "muted")
        banner_layout.addWidget(self._banner_line1)
        banner_layout.addWidget(self._banner_line2)

        # DEMO
        self._analysis_card, self._analysis_layout = make_card(title="RESULTADO DEL ANÁLISIS")
        self._review_card, self._review_layout = make_card(title="ESTADO TRAS LA REVISIÓN")
        self._sources_card, self._sources_layout = make_card(title="FUENTES")
        self._modules_card, self._modules_layout = make_card(title="MÓDULOS")
        self._template_card, self._template_layout = make_card(title="PLANTILLA OFICIAL")

        # REAL
        self._real_registros_card, self._real_registros_layout = make_card(title="REGISTROS")
        self._real_legacy_card, self._real_legacy_layout = make_card(title="CLASIFICACIÓN LEGACY")
        self._real_columns_card, self._real_columns_layout = make_card(title="COLUMNAS DETECTADAS")
        self._real_diag_card, self._real_diag_layout = make_card(title="DIAGNÓSTICOS FUNCIONALES")

        self.generate_button = QPushButton("Generar REMASEP")
        self.generate_button.setProperty("variant", "primary")
        self.generate_button.setEnabled(False)
        self.generate_button.setToolTip(_GENERATE_TOOLTIP)

        self.back_button = QPushButton("Volver")
        self.back_button.clicked.connect(self._go_back)
        self.home_button = QPushButton("Volver al inicio")
        self.home_button.clicked.connect(self._app.reset_flow)

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addWidget(self.home_button)
        actions.addStretch()
        actions.addWidget(self.generate_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 24, 40, 24)
        layout.setSpacing(12)
        layout.addWidget(self.steps)
        layout.addWidget(self.title_label)
        layout.addWidget(self.real_banner)
        for widget in (
            self._analysis_card,
            self._review_card,
            self._sources_card,
            self._modules_card,
            self._template_card,
            self._real_registros_card,
            self._real_legacy_card,
            self._real_columns_card,
            self._real_diag_card,
        ):
            layout.addWidget(widget)
        layout.addStretch()
        layout.addLayout(actions)

    def _go_back(self) -> None:
        self._app.navigate("analysis" if self._app.state.is_real else "exceptions")

    def on_enter(self) -> None:
        self.steps.set_current(3)
        state = self._app.state
        is_real = state.is_real

        demo_widgets = (
            self._analysis_card,
            self._review_card,
            self._sources_card,
            self._modules_card,
            self._template_card,
        )
        real_widgets = (
            self.real_banner,
            self._real_registros_card,
            self._real_legacy_card,
            self._real_columns_card,
            self._real_diag_card,
        )
        for widget in demo_widgets:
            widget.setVisible(not is_real)
        for widget in real_widgets:
            widget.setVisible(is_real)

        if is_real:
            self._render_real(state.medinet_analysis)
        else:
            self._render_demo(state)

    # -- DEMO -----------------------------------------------------------

    def _render_demo(self, state) -> None:
        result = state.analysis
        outcome = state.review_outcome()
        if result is None or outcome is None:
            return

        self.title_label.setText(f"REMASEP — {result.period_label.upper()}")

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

    # -- REAL -----------------------------------------------------------

    def _render_real(self, result) -> None:
        if result is None:
            self.title_label.setText("REMASEP — ANÁLISIS MEDINET")
            return
        self.title_label.setText(f"REMASEP — {result.period_label.upper()}")
        self._banner_line1.setText(f"Análisis Medinet completado · {result.source_name}")
        refresh_style(self.real_banner)

        _clear(self._real_registros_layout, keep=1)
        rows = (
            ("Total", "neutral", result.total_records),
            ("Válidos", "ok", result.valid_records),
            ("Inválidos", "warning" if result.invalid_records else "ok", result.invalid_records),
            ("Dentro del período", "ok", result.records_in_period),
            (
                "Fuera del período",
                "warning" if result.records_outside_period else "neutral",
                result.records_outside_period,
            ),
        )
        for name, status, value in rows:
            self._real_registros_layout.addWidget(
                StatusRow(name, status=status, value=format_int(value))
            )
        if result.min_service_date and result.max_service_date:
            self._real_registros_layout.addWidget(
                StatusRow(
                    "Período observado",
                    status="neutral",
                    value=f"{result.min_service_date} → {result.max_service_date}",
                )
            )
        if result.structural_empty_rows:
            empty_note = QLabel(
                f"{format_int(result.structural_empty_rows)} filas estructurales vacías "
                "ignoradas."
            )
            empty_note.setProperty("role", "muted")
            empty_note.setWordWrap(True)
            self._real_registros_layout.addWidget(empty_note)

        _clear(self._real_legacy_layout, keep=1)
        for code in ("AG", "AH", "AI", "AJ", "AK", "AL"):
            self._real_legacy_layout.addWidget(
                StatusRow(
                    _LEGACY_UI_LABELS[code],
                    status="ok" if result.legacy_counts.get(code) else "neutral",
                    value=format_int(result.legacy_counts.get(code, 0)),
                )
            )
        self._real_legacy_layout.addWidget(
            StatusRow(
                "Válidos sin regla legacy",
                status="neutral",
                value=format_int(result.non_target_records),
            )
        )
        note = QLabel(
            "Categorías legacy: reproducen la lógica del workbook actual, "
            "pendientes de validación funcional."
        )
        note.setProperty("role", "muted")
        note.setWordWrap(True)
        self._real_legacy_layout.addWidget(note)

        _clear(self._real_columns_layout, keep=1)
        detected = QLabel(", ".join(result.detected_columns) or "—")
        detected.setWordWrap(True)
        self._real_columns_layout.addWidget(detected)
        if result.missing_optional_columns:
            missing = QLabel(
                "Opcionales ausentes: " + ", ".join(result.missing_optional_columns)
            )
            missing.setProperty("role", "muted")
            self._real_columns_layout.addWidget(missing)

        _clear(self._real_diag_layout, keep=1)
        for diagnostic in result.diagnostics:
            if not diagnostic.present:
                self._real_diag_layout.addWidget(
                    StatusRow(diagnostic.field, status="pending", value="ausente")
                )
                continue
            self._real_diag_layout.addWidget(
                StatusRow(
                    diagnostic.field,
                    status="ok",
                    value=f"{format_int(diagnostic.non_empty)} con valor · "
                    f"{format_int(diagnostic.unique_values)} únicos",
                )
            )
            for name, count in diagnostic.distribution[:8]:
                sub = QLabel(f"    {name}  ·  {format_int(count)}")
                sub.setProperty("role", "muted")
                self._real_diag_layout.addWidget(sub)
