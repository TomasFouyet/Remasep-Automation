"""Pantalla 'Análisis'.

Dos modos que nunca deben confundirse:

- **DEMO**  -> datos de ``MockRemasepService`` (ficticios).
- **REAL**  -> ``MedinetAnalysisService`` sobre el archivo Medinet elegido.
"""

from __future__ import annotations

from PySide6.QtWidgets import QHBoxLayout, QLabel, QPushButton, QVBoxLayout, QWidget

from remasep.ui.components.status_card import StatusRow, make_card
from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.styles import format_int, refresh_style

_STEPS = ["Archivos", "Análisis", "Revisión", "Resultado"]

_LEGACY_UI_LABELS = {
    "AG": "Consultas médicas",
    "AH": "Controles odontológicos",
    "AI": "Evaluaciones odontológicas",
    "AJ": "Controles ortodoncia",
    "AK": "Controles ortopedia",
    "AL": "Instalaciones ortopedia",
}


def _muted(text: str) -> QLabel:
    label = QLabel(text)
    label.setProperty("role", "muted")
    label.setWordWrap(True)
    return label


class AnalysisScreen(QWidget):
    name = "analysis"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        self.steps = StepIndicator(_STEPS)
        self.badge = QLabel()
        self.period_label = QLabel()
        self.period_label.setProperty("role", "h2")
        self.count_label = QLabel()
        self.count_label.setProperty("role", "metric")
        self.error_label = _muted("")
        self.error_label.setProperty("status", "error")

        # Comunes / DEMO
        self._validations_card, self._validations_layout = make_card(title="VALIDACIONES")
        self._result_card, self._result_layout = make_card(title="RESULTADO")
        self._modules_card, self._modules_layout = make_card(title="MÓDULOS DETECTADOS")

        # REAL
        self._registros_card, self._registros_layout = make_card(title="REGISTROS")
        self._legacy_card, self._legacy_layout = make_card(title="CLASIFICACIÓN LEGACY")
        self._diag_card, self._diag_layout = make_card(title="DIAGNÓSTICOS FUNCIONALES")
        self._legacy_note = _muted(
            "Estas categorías reproducen la lógica del workbook actual y están "
            "pendientes de validación funcional. No son reglas oficiales MINSAL."
        )

        self.back_button = QPushButton("Atrás")
        self.back_button.clicked.connect(lambda: self._app.navigate("new_report"))
        self.review_button = QPushButton("Revisar problemas")
        self.review_button.setProperty("variant", "primary")
        self.review_button.clicked.connect(lambda: self._app.navigate("exceptions"))
        self.summary_button = QPushButton("Ver resumen")
        self.summary_button.setProperty("variant", "primary")
        self.summary_button.clicked.connect(lambda: self._app.navigate("summary"))

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addStretch()
        actions.addWidget(self.review_button)
        actions.addWidget(self.summary_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 24, 40, 24)
        layout.setSpacing(12)
        layout.addWidget(self.steps)
        layout.addWidget(self.badge)
        layout.addWidget(self.period_label)
        layout.addWidget(self.count_label)
        layout.addWidget(self.error_label)
        for widget in (
            self._validations_card,
            self._result_card,
            self._modules_card,
            self._registros_card,
            self._legacy_card,
            self._legacy_note,
            self._diag_card,
        ):
            layout.addWidget(widget)
        layout.addStretch()
        layout.addLayout(actions)

    # -- navegación ----------------------------------------------------

    def on_enter(self) -> None:
        self.steps.set_current(1)
        state = self._app.state
        is_real = state.is_real

        self.badge.setText("ANÁLISIS REAL" if is_real else "DATOS DE DEMOSTRACIÓN")
        self.badge.setProperty("badge", "real" if is_real else "demo")
        refresh_style(self.badge)

        error = state.analysis_error if is_real else None
        self.error_label.setText(error or "")
        self.error_label.setVisible(bool(error))

        demo_widgets = (self._result_card, self._modules_card, self.review_button)
        real_widgets = (
            self._registros_card,
            self._legacy_card,
            self._legacy_note,
            self._diag_card,
            self.summary_button,
        )
        show_real = is_real and not error
        for widget in demo_widgets:
            widget.setVisible(not is_real)
        for widget in real_widgets:
            widget.setVisible(show_real)
        self._validations_card.setVisible(not error)
        self.period_label.setVisible(not error)
        self.count_label.setVisible(not error)

        if error:
            return
        if is_real:
            self._render_real(state.medinet_analysis)
        else:
            self._render_demo(state.analysis)

    # -- DEMO --------------------------------------------------------

    def _render_demo(self, result) -> None:
        if result is None:
            return
        self.period_label.setText(result.period_label)
        self.count_label.setText(f"{format_int(result.total_records)} atenciones encontradas")

        _clear(self._validations_layout, keep=1)
        for validation in result.validations:
            self._validations_layout.addWidget(StatusRow(validation.name, status=validation.status))

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
        self._modules_layout.addWidget(_muted("Otras fuentes"))
        for source in result.other_sources:
            self._modules_layout.addWidget(
                StatusRow(source.name, status="pending", value=source.note)
            )

    # -- REAL --------------------------------------------------------

    def _render_real(self, result) -> None:
        if result is None:
            return
        self.period_label.setText(f"{result.period_label}  ·  {result.source_name}")
        self.count_label.setText(f"{format_int(result.total_records)} registros analizados")

        _clear(self._validations_layout, keep=1)
        for validation in result.validations:
            self._validations_layout.addWidget(
                StatusRow(
                    validation.name,
                    status=validation.status,
                    value=validation.message or None,
                )
            )

        _clear(self._registros_layout, keep=1)
        self._registros_layout.addWidget(
            StatusRow("Total", status="neutral", value=format_int(result.total_records))
        )
        self._registros_layout.addWidget(
            StatusRow("Válidos", status="ok", value=format_int(result.valid_records))
        )
        self._registros_layout.addWidget(
            StatusRow(
                "Inválidos",
                status="warning" if result.invalid_records else "ok",
                value=format_int(result.invalid_records),
            )
        )
        self._registros_layout.addWidget(
            StatusRow(
                "Dentro del período", status="ok", value=format_int(result.records_in_period)
            )
        )
        self._registros_layout.addWidget(
            StatusRow(
                "Fuera del período",
                status="neutral",
                value=format_int(result.records_outside_period),
            )
        )
        self._registros_layout.addWidget(
            StatusRow(
                f"Procesados para {result.period_label}",
                status="ok",
                value=format_int(result.processing_scope_records),
            )
        )
        if result.records_outside_period:
            self._registros_layout.addWidget(
                _muted(
                    f"{format_int(result.records_outside_period)} registros pertenecen a "
                    f"otros períodos y no se incluirán en el REMASEP de "
                    f"{result.period_label}. No se descartan del archivo."
                )
            )
        period_dates = "—"
        if result.min_service_date and result.max_service_date:
            period_dates = f"{result.min_service_date} → {result.max_service_date}"
        self._registros_layout.addWidget(
            StatusRow("Período observado", status="neutral", value=period_dates)
        )
        if result.structural_empty_rows:
            self._registros_layout.addWidget(
                _muted(
                    f"{format_int(result.structural_empty_rows)} filas estructurales vacías "
                    "ignoradas (fórmulas arrastradas más allá de las atenciones)."
                )
            )

        _clear(self._legacy_layout, keep=1)
        self._legacy_layout.addWidget(
            _muted(
                f"Clasificación sobre los {format_int(result.processing_scope_records)} "
                f"registros de {result.period_label} (no sobre todo el archivo)."
            )
        )
        for code in ("AG", "AH", "AI", "AJ", "AK", "AL"):
            self._legacy_layout.addWidget(
                StatusRow(
                    _LEGACY_UI_LABELS[code],
                    status="ok" if result.legacy_counts.get(code) else "neutral",
                    value=format_int(result.legacy_counts.get(code, 0)),
                )
            )
        self._legacy_layout.addWidget(
            StatusRow(
                "Sin regla legacy (en el período)",
                status="neutral",
                value=format_int(result.non_target_records),
            )
        )

        _clear(self._diag_layout, keep=1)
        for diagnostic in result.diagnostics:
            if not diagnostic.present:
                self._diag_layout.addWidget(
                    StatusRow(diagnostic.field, status="pending", value="ausente")
                )
                continue
            self._diag_layout.addWidget(
                StatusRow(
                    diagnostic.field,
                    status="ok",
                    value=f"{format_int(diagnostic.non_empty)} con valor · "
                    f"{format_int(diagnostic.unique_values)} únicos",
                )
            )
            for name, count in diagnostic.distribution[:6]:
                self._diag_layout.addWidget(_muted(f"    {name}  ·  {format_int(count)}"))
            if diagnostic.note:
                self._diag_layout.addWidget(_muted(f"    {diagnostic.note}"))


def _clear(layout: QVBoxLayout, *, keep: int) -> None:
    """Elimina los widgets del layout salvo los primeros ``keep`` (p.ej. el título)."""
    while layout.count() > keep:
        item = layout.takeAt(keep)
        widget = item.widget()
        if widget is not None:
            widget.deleteLater()
