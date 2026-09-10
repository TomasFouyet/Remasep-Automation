"""Pantalla 'Resumen mensual' — fuente **Medinet** únicamente.

KPIs + 4 gráficos agregados (sin PII). Desde aquí se exporta el PDF y se lanza
la generación del REMASEP.
"""

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

from remasep.services.medinet_summary import MonthlyMedinetSummary
from remasep.ui.charts import Bar, BarChartCard
from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.components.widgets import Card, MetricCard, SectionHeader, StatusBadge
from remasep.ui.pdf_report import export_summary_pdf, suggested_pdf_name
from remasep.ui.styles import THEME

_STEPS = ["Datos", "Análisis", "Resumen", "Informe"]
_DISCLAIMER = (
    "Este resumen corresponde a los datos disponibles en Medinet y no representa "
    "todavía la totalidad del informe REMASEP."
)


class DashboardScreen(QWidget):
    name = "dashboard"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        self.steps = StepIndicator(_STEPS)

        header = QHBoxLayout()
        title_box = QVBoxLayout()
        title_box.setSpacing(2)
        self._title = QLabel("Resumen mensual")
        self._title.setProperty("role", "h2")
        self._period = QLabel()
        self._period.setProperty("role", "subtitle")
        title_box.addWidget(self._title)
        title_box.addWidget(self._period)
        header.addLayout(title_box)
        header.addStretch()
        header.addWidget(StatusBadge("FUENTE: MEDINET", kind="info"), alignment=Qt.AlignmentFlag.AlignTop)

        disclaimer = QLabel(_DISCLAIMER)
        disclaimer.setProperty("role", "faint")
        disclaimer.setWordWrap(True)

        # --- KPIs ---
        self.kpi_period = MetricCard("Citas del período", hint="Válidas dentro del mes")
        self.kpi_included = MetricCard(
            "Consideradas para REMASEP", hint="Sólo estados confirmados", accent=True
        )
        self.kpi_excluded = MetricCard(
            "Excluidas por estado", hint="Canceladas, no presentadas, etc."
        )
        self.kpi_ratio = MetricCard("Porcentaje considerado")
        kpi_row = QHBoxLayout()
        kpi_row.setSpacing(THEME.gap)
        for card in (self.kpi_period, self.kpi_included, self.kpi_excluded, self.kpi_ratio):
            kpi_row.addWidget(card)

        # --- gráficos ---
        self.chart_estado = BarChartCard(
            "Distribución por estado",
            subtitle="Se consideran únicamente los estados confirmados para REMASEP.",
            legend=[("Consideradas", THEME.chart_included), ("Excluidas", THEME.chart_excluded)],
        )
        self.chart_service = BarChartCard(
            "Atenciones consideradas por especialidad",
            subtitle="Sobre las atenciones consideradas para el REMASEP.",
        )
        self.chart_sex = BarChartCard("Consideradas por sexo")
        self.chart_age = BarChartCard("Consideradas por tramo de edad")
        charts = QGridLayout()
        charts.setSpacing(THEME.gap)
        charts.addWidget(self.chart_estado, 0, 0)
        charts.addWidget(self.chart_service, 0, 1)
        charts.addWidget(self.chart_sex, 1, 0)
        charts.addWidget(self.chart_age, 1, 1)

        # --- pendientes de otras fuentes ---
        self._pending_card = Card(muted=True)
        self._pending_card.body.addWidget(
            SectionHeader(
                "Pendientes de otras fuentes",
                hint="Estas secciones del REMASEP se completarán en próximas versiones.",
            )
        )
        self._pending_list = QVBoxLayout()
        self._pending_list.setSpacing(4)
        self._pending_card.body.addLayout(self._pending_list)

        # --- contenido desplazable ---
        content = QWidget()
        content_layout = QVBoxLayout(content)
        content_layout.setContentsMargins(0, 0, 8, 0)
        content_layout.setSpacing(THEME.gap_lg)
        content_layout.addLayout(kpi_row)
        content_layout.addLayout(charts)
        content_layout.addWidget(self._pending_card)
        content_layout.addStretch()

        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setWidget(content)
        scroll.setFrameShape(QFrame.Shape.NoFrame)

        # --- acciones ---
        self.back_button = QPushButton("Volver")
        self.back_button.setProperty("variant", "ghost")
        self.back_button.clicked.connect(lambda: self._app.navigate("new_report"))
        self.pdf_button = QPushButton("Exportar resumen PDF")
        self.pdf_button.clicked.connect(self._export_pdf)
        self.generate_button = QPushButton("Generar REMASEP")
        self.generate_button.setProperty("variant", "primary")
        self.generate_button.clicked.connect(lambda: self._app.navigate("generate"))

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
        self.kpi_ratio.set_value(f"{s.included_percentage:g}%")
        self.kpi_ratio.set_hint(s.considered_ratio_label)

        self.chart_estado.set_bars([
            Bar(e.label, e.count, THEME.chart_included if e.included else THEME.chart_excluded)
            for e in s.estado_distribution
        ])
        self.chart_service.set_bars(_series_bars(s.service_distribution))
        self.chart_sex.set_bars(_series_bars(s.sex_distribution))
        self.chart_age.set_bars(_series_bars(s.age_distribution))

        while self._pending_list.count():
            item = self._pending_list.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        for name in s.pending_sources:
            row = QLabel(f"•  {name}")
            row.setProperty("role", "muted")
            self._pending_list.addWidget(row)

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
