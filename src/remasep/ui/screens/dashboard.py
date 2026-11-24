"""Preparación del borrador: alcance, pendientes y detalle agregado Medinet."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt
from PySide6.QtWidgets import (
    QFileDialog,
    QFrame,
    QGridLayout,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QScrollArea,
    QVBoxLayout,
    QWidget,
)

from remasep.services.medinet_summary import MANUAL_REVIEW, MonthlyMedinetSummary
from remasep.ui.charts import Bar, BarChartCard
from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.components.widgets import Card, MetricCard, SectionHeader, StatusBadge
from remasep.ui.operation_lifecycle import OperationKind
from remasep.ui.pdf_report import export_summary_pdf, suggested_pdf_name
from remasep.ui.styles import THEME, format_int

_STEPS = ["Datos", "Análisis", "Resumen", "Informe"]
_DISCLAIMER = (
    "Puedes preparar un borrador con Medinet. Antes de enviarlo debes completar "
    "los datos pendientes y revisar CONTROL en Excel."
)
_SHEET_DESCRIPTIONS = {
    "REMASEP 01": "Consultas y otras atenciones ambulatorias.",
    "REMASEP_OD": "Consultas, controles y tratamientos odontológicos.",
    "B2 ANEXO": "Procedimientos seleccionados que se pueden contar desde Medinet.",
}


class DashboardScreen(QWidget):
    name = "dashboard"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        self.steps = StepIndicator(_STEPS)

        header = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        self._title = QLabel("¿Qué incluirá el borrador?")
        self._title.setProperty("role", "h2")
        self._period = QLabel()
        self._period.setProperty("role", "subtitle")
        title_box.addWidget(self._title)
        title_box.addWidget(self._period)
        header.addLayout(title_box)
        header.addStretch()
        header.addWidget(StatusBadge("BORRADOR PARCIAL", kind="warning"), alignment=Qt.AlignmentFlag.AlignTop)

        disclaimer = QLabel(_DISCLAIMER)
        disclaimer.setProperty("role", "muted")
        disclaimer.setWordWrap(True)
        self._readiness = disclaimer

        self._validation_warning = QLabel()
        self._validation_warning.setObjectName("bannerError")
        self._validation_warning.setWordWrap(True)
        self._validation_warning.setVisible(False)

        # --- KPIs ---
        self.kpi_period = MetricCard(
            "Citas revisadas", hint="Antes de excluir por estado"
        )
        self.kpi_included = MetricCard(
            "Citas incluidas", hint="Según agenda, estado y período", accent=True
        )
        self.kpi_excluded = MetricCard(
            "Excluidas por estado", hint="Canceladas, no presentadas, etc."
        )
        self.kpi_writes = MetricCard("Celdas previstas", hint="No es un conteo de citas")
        kpi_row = QHBoxLayout()
        kpi_row.setSpacing(THEME.gap)
        for card in (self.kpi_included, self.kpi_writes, self.kpi_period, self.kpi_excluded):
            kpi_row.addWidget(card)

        self._scope_note = QLabel()
        self._scope_note.setWordWrap(True)
        self._scope_note.setProperty("role", "muted")
        self._included_states = QLabel()
        self._included_states.setWordWrap(True)
        self._included_states.setProperty("role", "faint")

        self._coverage_card = Card()
        self._coverage_card.body.addWidget(SectionHeader("La app completará", hint="Sólo los campos automatizados de estas hojas."))
        self._coverage_list = QVBoxLayout()
        self._coverage_card.body.addLayout(self._coverage_list)
        self._coverage_card.body.addWidget(_wrapped(
            "Usa la plantilla oficial en blanco. El resto de sus campos no se completa "
            "automáticamente. Las fórmulas y macros se conservan.", faint=True
        ))

        self._pending_card = Card(muted=True)
        self._pending_card.body.addWidget(SectionHeader("Completar o revisar antes del envío"))
        for text in MANUAL_REVIEW:
            self._pending_card.body.addWidget(_wrapped(f"• {text}"))
        self._pending_list = QVBoxLayout()
        self._pending_card.body.addLayout(self._pending_list)
        self._pending_card.body.addWidget(_wrapped(
            "CONTROL se comprueba después de generar. Esta pantalla no confirma "
            "que el informe esté completo ni libre de alertas.", faint=True
        ))
        preparation = QHBoxLayout()
        preparation.setSpacing(THEME.gap)
        preparation.addWidget(self._coverage_card, stretch=1)
        preparation.addWidget(self._pending_card, stretch=1)

        # --- gráficos ---
        self.chart_estado = BarChartCard(
            "Distribución por estado",
            subtitle="Estados incluidos y excluidos según la regla aplicada.",
            legend=[("Consideradas", THEME.chart_included), ("Excluidas", THEME.chart_excluded)],
        )
        self.chart_service = BarChartCard(
            "Atenciones consideradas por especialidad",
            subtitle="Sobre las atenciones consideradas para el REMASEP.",
        )
        self.chart_sex = BarChartCard("Consideradas por sexo")
        self.chart_age = BarChartCard("Consideradas por tramo de edad")
        self._charts_container = QWidget()
        charts = QGridLayout(self._charts_container)
        charts.setContentsMargins(0, 0, 0, 0)
        charts.setSpacing(THEME.gap)
        charts.addWidget(self.chart_estado, 0, 0)
        charts.addWidget(self.chart_service, 0, 1)
        charts.addWidget(self.chart_sex, 1, 0)
        charts.addWidget(self.chart_age, 1, 1)

        self.details_button = QPushButton("Ver detalle de las citas: estado, especialidad, sexo y edad")
        self.details_button.setCheckable(True)
        self.details_button.toggled.connect(self._charts_container.setVisible)
        self._charts_container.setVisible(False)

        # --- contenido desplazable ---
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 8, 0)
        content_layout.setSpacing(THEME.gap_lg)
        content_layout.addLayout(kpi_row)
        content_layout.addWidget(self._validation_warning)
        content_layout.addWidget(self._scope_note)
        content_layout.addWidget(self._included_states)
        content_layout.addLayout(preparation)
        content_layout.addWidget(self.details_button)
        content_layout.addWidget(self._charts_container)
        content_layout.addStretch()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(content)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        # --- acciones ---
        self.back_button = QPushButton("Volver")
        self.back_button.setProperty("variant", "ghost")
        self.back_button.clicked.connect(lambda: self._app.navigate("new_report"))
        self.pdf_button = QPushButton("Exportar estadísticas PDF")
        self.pdf_button.clicked.connect(self._export_pdf)
        self.generate_button = QPushButton("Generar borrador REMASEP")
        self.generate_button.setProperty("variant", "primary")
        self.generate_button.clicked.connect(lambda: self._app.navigate("generate"))
        self._app.operations.operation_finished.connect(self._sync_generate_button)

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addStretch()
        actions.addWidget(self.pdf_button)
        actions.addWidget(self.generate_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 24, 40, 20)
        layout.setSpacing(THEME.gap)
        layout.addWidget(self.steps)
        layout.addLayout(header)
        layout.addWidget(disclaimer)
        layout.addWidget(scroll, stretch=1)
        layout.addLayout(actions)

    # --- render --------------------------------------------------

    def on_enter(self) -> None:
        self.steps.set_current(2)
        summary = self._app.state.summary
        if summary is None:
            self._app.navigate("new_report")
            return
        self._render(summary)

    def _render(self, s: MonthlyMedinetSummary) -> None:
        self._period.setText(f"{s.period_label}  ·  fuente: {s.source}")
        self.kpi_period.set_value(s.period_scope_records)
        self.kpi_included.set_value(s.included_records)
        self.kpi_excluded.set_value(s.excluded_records)
        self.kpi_writes.set_value(
            s.planned_write_count if s.planned_write_count and not s.validation_blocked else "—"
        )
        self.kpi_writes.set_hint(
            "Generación bloqueada por datos inválidos" if s.validation_blocked else
            f"En {len(s.write_targets_by_sheet)} hojas; no son citas"
            if s.write_targets_by_sheet else "Alcance pendiente de cargar"
        )
        self._readiness.setText(
            "Corrige los registros inválidos de Medinet antes de preparar el borrador."
            if s.validation_blocked else _DISCLAIMER
        )
        self._coverage_card.setVisible(not s.validation_blocked)
        self.kpi_period.set_hint(
            "Ambulatorias, antes de excluir por estado"
            if s.agenda_filter_applied else "Agenda no informada; antes de excluir por estado"
        )
        if s.dated_period_records is not None:
            self._scope_note.setText(
                f"En el archivo: {format_int(s.dated_period_records)} citas con fecha en el mes. "
                f"Fuera por agenda: {format_int(s.agenda_excluded_records)}. "
                f"Fuera por estado: {format_int(s.excluded_records)}. "
                f"Registros inválidos relevantes: {format_int(s.invalid_records)}."
            )
        else:
            self._scope_note.setText("Resumen de citas del período seleccionado.")
        self._included_states.setText("Estados incluidos: " + ", ".join(s.estado_included_states) + ".")
        if s.validation_blocked:
            noun = "registro inválido" if s.invalid_records == 1 else "registros inválidos"
            self._validation_warning.setText(
                f"No se puede generar el REMASEP: hay {s.invalid_records} {noun} "
                "relevante(s) para este período. Corrige el archivo de Medinet y vuelve "
                "a analizarlo. Ningún registro inválido fue contado."
            )
            self._validation_warning.setVisible(True)
        else:
            self._validation_warning.clear()
            self._validation_warning.setVisible(False)
        self._sync_generate_button()

        self.chart_estado.set_bars([
            Bar(e.label, e.count, THEME.chart_included if e.included else THEME.chart_excluded)
            for e in s.estado_distribution
        ])
        self.chart_service.set_bars(_series_bars(s.service_distribution))
        self.chart_sex.set_bars(_series_bars(s.sex_distribution))
        self.chart_age.set_bars(_series_bars(s.age_distribution))

        while self._coverage_list.count():
            item = self._coverage_list.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        order = {name: i for i, name in enumerate(_SHEET_DESCRIPTIONS)}
        for entry in sorted(s.write_targets_by_sheet, key=lambda e: (order.get(e.label, 99), e.label)):
            group = QWidget()
            body = QVBoxLayout(group)
            body.setContentsMargins(0, 0, 0, 0)
            body.setSpacing(2)
            body.addWidget(_wrapped(f"{entry.label}: {format_int(entry.count)} celdas previstas"))
            if entry.label in _SHEET_DESCRIPTIONS:
                body.addWidget(_wrapped(_SHEET_DESCRIPTIONS[entry.label], faint=True))
            self._coverage_list.addWidget(group)
        if not s.write_targets_by_sheet:
            self._coverage_list.addWidget(_wrapped("Vuelve a analizar el archivo para cargar el alcance."))

        while self._pending_list.count():
            item = self._pending_list.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        if s.pending_sources:
            self._pending_list.addWidget(_wrapped(
                "Otras fuentes necesarias: " + "; ".join(s.pending_sources) + ".", faint=True
            ))

    def _sync_generate_button(self, *_args) -> None:
        summary = self._app.state.summary
        validation_blocked = summary is None or summary.validation_blocked
        generation_running = self._app.operations.has_running(OperationKind.GENERATION)
        self.generate_button.setEnabled(not validation_blocked and not generation_running)
        if generation_running:
            self.generate_button.setToolTip("Ya hay una generación en curso.")
        else:
            self.generate_button.setToolTip("")

    # --- PDF ----------------------------------------------------

    def _export_pdf(self) -> None:
        summary = self._app.state.summary
        if summary is None:
            return
        suggested = str(Path.home() / suggested_pdf_name(summary))
        path, _ = QFileDialog.getSaveFileName(
            self, "Guardar resumen PDF", suggested, "PDF (*.pdf)"
        )
        if not path:
            return
        target = Path(path)
        if target.suffix.lower() != ".pdf":
            target = target.with_suffix(".pdf")
        try:
            export_summary_pdf(summary, target)
        except FileExistsError:
            QMessageBox.warning(
                self,
                "El archivo ya existe",
                f"Ya existe un archivo llamado “{target.name}”. Elige otro nombre.",
            )
            return
        except Exception:  # noqa: BLE001
            QMessageBox.warning(
                self, "No se pudo exportar", "No pudimos generar el PDF. Vuelve a intentarlo."
            )
            return
        QMessageBox.information(
            self, "Resumen exportado", f"El resumen se guardó en:\n{target}"
        )


def _series_bars(items) -> list[Bar]:
    return [
        Bar(c.label, c.count, THEME.chart_series[i % len(THEME.chart_series)])
        for i, c in enumerate(items)
    ]


def _wrapped(text: str, *, faint: bool = False) -> QLabel:
    label = QLabel(text)
    label.setWordWrap(True)
    label.setProperty("role", "faint" if faint else "muted")
    return label
