"""Pantalla 'Nuevo reporte': período + fuentes + modo demostración."""

from __future__ import annotations

from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from remasep.services.mock_remasep import month_name
from remasep.ui.components.file_selector import FileSelector
from remasep.ui.components.status_card import make_card
from remasep.ui.components.step_indicator import StepIndicator

_STEPS = ["Archivos", "Análisis", "Revisión", "Resultado"]
_YEARS = list(range(2024, 2031))


class NewReportScreen(QWidget):
    name = "new_report"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app

        self.steps = StepIndicator(_STEPS)

        heading = QLabel("Nuevo reporte mensual")
        heading.setProperty("role", "h2")

        # --- Período ---
        period_card, period_layout = make_card(title="PERÍODO")
        self.month_combo = QComboBox()
        for index in range(1, 13):
            self.month_combo.addItem(month_name(index), index)
        self.year_spin = QSpinBox()
        self.year_spin.setRange(_YEARS[0], _YEARS[-1])

        period_row = QHBoxLayout()
        month_box = QVBoxLayout()
        month_box.addWidget(QLabel("Mes"))
        month_box.addWidget(self.month_combo)
        year_box = QVBoxLayout()
        year_box.addWidget(QLabel("Año"))
        year_box.addWidget(self.year_spin)
        period_row.addLayout(month_box)
        period_row.addLayout(year_box)
        period_row.addStretch()
        period_layout.addLayout(period_row)

        self.month_combo.currentIndexChanged.connect(self._sync_state)
        self.year_spin.valueChanged.connect(self._sync_state)

        # --- Fuentes ---
        sources_label = QLabel("FUENTES")
        sources_label.setProperty("role", "section")

        self.medinet_selector = FileSelector(
            "A) Reporte Medinet",
            extensions=["xlsx", "xlsm", "csv"],
            hint="Obligatorio en la versión final. Durante esta demo puede usarse el modo "
            "demostración sin archivo.",
        )
        self.egresos_selector = FileSelector(
            "B) Egresos hospitalarios",
            extensions=["xlsx", "xlsm", "csv"],
            badge="Pendiente integración",
            hint="Usado para información quirúrgica y otros datos por confirmar.",
        )
        self.recursos_selector = FileSelector(
            "C) Recursos / Pabellones",
            extensions=["xlsx", "xlsm", "csv"],
            badge="Pendiente definición de fuente",
        )
        self.medinet_selector.fileSelected.connect(self._sync_state)
        self.egresos_selector.fileSelected.connect(self._sync_state)
        self.recursos_selector.fileSelected.connect(self._sync_state)

        # --- Acciones ---
        self.back_button = QPushButton("Atrás")
        self.back_button.clicked.connect(lambda: self._app.navigate("home"))

        self.demo_button = QPushButton("Usar datos de demostración")
        self.demo_button.clicked.connect(self._start_demo)

        self.analyze_button = QPushButton("Analizar")
        self.analyze_button.setProperty("variant", "primary")
        self.analyze_button.clicked.connect(self._start_real)

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addStretch()
        actions.addWidget(self.demo_button)
        actions.addWidget(self.analyze_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(40, 28, 40, 24)
        layout.setSpacing(14)
        layout.addWidget(self.steps)
        layout.addWidget(heading)
        layout.addWidget(period_card)
        layout.addWidget(sources_label)
        layout.addWidget(self.medinet_selector)
        layout.addWidget(self.egresos_selector)
        layout.addWidget(self.recursos_selector)
        layout.addStretch()
        layout.addLayout(actions)

    # --- navegación / estado ------------------------------------------

    def on_enter(self) -> None:
        self.steps.set_current(0)
        state = self._app.state
        # Poblar desde el estado sin disparar _sync_state a mitad de camino.
        self.month_combo.blockSignals(True)
        self.year_spin.blockSignals(True)
        self.month_combo.setCurrentIndex(state.month - 1)
        self.year_spin.setValue(state.year)
        self.month_combo.blockSignals(False)
        self.year_spin.blockSignals(False)
        self._sync_state()

    def _sync_state(self, *_args: object) -> None:
        state = self._app.state
        state.month = self.month_combo.currentData()
        state.year = self.year_spin.value()
        state.medinet_path = self.medinet_selector.path
        state.egresos_path = self.egresos_selector.path
        state.recursos_path = self.recursos_selector.path
        self._refresh_analyze_enabled()

    def _refresh_analyze_enabled(self) -> None:
        self.analyze_button.setEnabled(self.medinet_selector.path is not None)

    def _start_demo(self) -> None:
        self._sync_state()
        self._app.state.demo = True
        self._app.run_analysis()
        self._app.navigate("analysis")

    def _start_real(self) -> None:
        self._sync_state()
        self._app.state.demo = False
        self._app.run_analysis()
        self._app.navigate("analysis")
