"""Pantalla 'Nuevo informe': período + archivo Medinet + plantilla REMASEP."""

from __future__ import annotations

from pathlib import Path

from PySide6.QtWidgets import (
    QComboBox,
    QHBoxLayout,
    QLabel,
    QPushButton,
    QSpinBox,
    QVBoxLayout,
    QWidget,
)

from remasep.core.errors import RemasepError
from remasep.services.common import Period, month_name
from remasep.services.excel_writer import check_template_compatibility
from remasep.services.medinet_analysis import detect_medinet_periods
from remasep.services.runtime_assets import RuntimeAssetError, load_runtime_bundle
from remasep.ui.components.file_selector import FileSelector
from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.components.widgets import Card, SectionHeader
from remasep.ui.errors import humanize_error

_STEPS = ["Datos", "Análisis", "Resumen", "Informe"]
_YEARS = list(range(2024, 2031))
_FALLBACK_FINGERPRINT = "stf:dc624775927d4d4d"


def _expected_fingerprint() -> str:
    try:
        return load_runtime_bundle().template_fingerprint_id
    except RuntimeAssetError:
        return _FALLBACK_FINGERPRINT


class NewReportScreen(QWidget):
    name = "new_report"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app
        # Períodos (mes/año) presentes en el archivo Medinet elegido.
        #   None  -> aún no se detectó / no se pudo leer (no bloquea; se revalida al analizar)
        #   ()    -> archivo leído pero sin ninguna cita con fecha válida (bloquea)
        #   (...) -> períodos reales disponibles
        self._available_periods: tuple[Period, ...] | None = None

        self.steps = StepIndicator(_STEPS)
        heading = QLabel("Nuevo informe mensual")
        heading.setProperty("role", "h2")

        # --- período ---
        period_card = Card()
        period_card.body.addWidget(SectionHeader("Período"))
        self.month_combo = QComboBox()
        for index in range(1, 13):
            self.month_combo.addItem(month_name(index), index)
        self.year_spin = QSpinBox()
        self.year_spin.setRange(_YEARS[0], _YEARS[-1])

        row = QHBoxLayout()
        row.setSpacing(16)
        mbox = QVBoxLayout()
        mbox.setSpacing(4)
        mbox.addWidget(QLabel("Mes"))
        mbox.addWidget(self.month_combo)
        ybox = QVBoxLayout()
        ybox.setSpacing(4)
        ybox.addWidget(QLabel("Año"))
        ybox.addWidget(self.year_spin)
        row.addLayout(mbox)
        row.addLayout(ybox)
        row.addStretch()
        period_card.body.addLayout(row)
        self.month_combo.currentIndexChanged.connect(self._sync_state)
        self.year_spin.valueChanged.connect(self._sync_state)

        # --- archivos ---
        self.medinet_selector = FileSelector(
            "Archivo de Medinet",
            extensions=["xlsx", "xlsm"],
            hint="Export directo “Detalle de citas” del período elegido.",
        )
        self.template_selector = FileSelector(
            "Plantilla REMASEP",
            extensions=["xlsm"],
            hint="Archivo oficial REMASEP (.xlsm) sobre el que se escribirá el informe.",
        )
        self.medinet_selector.fileSelected.connect(self._on_medinet_selected)
        self.template_selector.fileSelected.connect(self._on_template_selected)

        # --- acciones ---
        self.back_button = QPushButton("Volver")
        self.back_button.setProperty("variant", "ghost")
        self.back_button.clicked.connect(lambda: self._app.navigate("home"))

        self.analyze_button = QPushButton("Analizar datos")
        self.analyze_button.setProperty("variant", "primary")
        self.analyze_button.clicked.connect(self._analyze)

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addStretch()
        actions.addWidget(self.analyze_button)

        layout = QVBoxLayout(self)
        layout.setContentsMargins(48, 28, 48, 24)
        layout.setSpacing(16)
        layout.addWidget(self.steps)
        layout.addWidget(heading)
        layout.addWidget(period_card)
        layout.addWidget(self.medinet_selector)
        layout.addWidget(self.template_selector)
        layout.addStretch()
        layout.addLayout(actions)

    # --- estado ---------------------------------------------------

    def on_enter(self) -> None:
        self.steps.set_current(0)
        state = self._app.state
        self._apply_period(Period(state.month, state.year))
        if self.medinet_selector.path is None:
            self._available_periods = None
        self._sync_state()

    def _apply_period(self, period: Period) -> None:
        """Fija mes/año en los selectores sin disparar señales."""
        self.month_combo.blockSignals(True)
        self.year_spin.blockSignals(True)
        self.month_combo.setCurrentIndex(period.month - 1)
        self.year_spin.setValue(period.year)
        self.month_combo.blockSignals(False)
        self.year_spin.blockSignals(False)

    def _on_medinet_selected(self, *_args: object) -> None:
        """Detección asistida del período al elegir el archivo Medinet.

        Un export puede contener varios meses; el período real se obtiene de las
        fechas de cita, nunca del nombre del archivo.
        """
        self._available_periods = None
        path = self.medinet_selector.path
        if path is None or not self.medinet_selector.has_valid_extension:
            self._sync_state()
            return

        try:
            periods = detect_medinet_periods(Path(path))
        except RemasepError as exc:
            self.medinet_selector.set_status(humanize_error(exc).title, status="warning")
            self._sync_state()
            return
        except Exception:  # noqa: BLE001 - se revalida al analizar; nunca crashear aquí
            self.medinet_selector.set_status(
                "No pudimos leer los períodos del archivo ahora; se revisará al analizar.",
                status="warning",
            )
            self._sync_state()
            return

        self._available_periods = tuple(periods)
        if not periods:
            self.medinet_selector.set_status(
                "No encontramos citas con fecha válida en este archivo.",
                status="warning",
            )
            self._sync_state()
            return

        current = Period(self.month_combo.currentData(), self.year_spin.value())
        if len(periods) == 1:
            self._apply_period(periods[0])
            self.medinet_selector.set_status(
                f"✓  Período detectado: {periods[0].label}", status="ok"
            )
        elif current in periods:
            self.medinet_selector.set_status(
                f"El archivo contiene {len(periods)} períodos. "
                f"Se procesará {current.label} (puedes cambiarlo).",
                status="ok",
            )
        else:
            self._apply_period(periods[-1])  # más reciente: elección determinista
            self.medinet_selector.set_status(
                "El archivo contiene varios períodos. Confirma el mes que deseas procesar.",
                status="warning",
            )
        self._sync_state()

    def _sync_state(self, *_args: object) -> None:
        state = self._app.state
        state.month = self.month_combo.currentData()
        state.year = self.year_spin.value()
        state.medinet_path = self.medinet_selector.path
        state.template_path = self.template_selector.path

        files_ok = (
            self.medinet_selector.has_valid_extension
            and self.template_selector.has_valid_extension
        )
        period_ok = True
        if self._available_periods:  # tupla no vacía: hay períodos conocidos
            chosen = Period(state.month, state.year)
            period_ok = chosen in self._available_periods
            if not period_ok and self.medinet_selector.has_valid_extension:
                self.medinet_selector.set_status(
                    f"No encontramos citas de {chosen.label} en este archivo.",
                    status="warning",
                )
        elif self._available_periods == ():  # archivo leído, sin fechas válidas
            period_ok = False

        self.analyze_button.setEnabled(files_ok and period_ok)

    def _on_template_selected(self, *_args: object) -> None:
        self._sync_state()
        path = self.template_selector.path
        if path is None or not self.template_selector.has_valid_extension:
            return
        try:
            compat = check_template_compatibility(Path(path), _expected_fingerprint())
        except Exception:  # noqa: BLE001 - la compatibilidad real se revalida al generar
            self.template_selector.set_status(
                "No pudimos verificar la plantilla ahora; se revisará al generar.",
                status="warning",
            )
            return
        if compat.compatible:
            self.template_selector.set_status("✓  Plantilla compatible", status="ok")
        else:
            self.template_selector.set_status(
                "La plantilla no corresponde a la versión REMASEP compatible.",
                status="warning",
            )

    def _analyze(self) -> None:
        self._sync_state()
        self._app.state.summary = None
        self._app.state.analysis_error = None
        self._app.navigate("analysis")
