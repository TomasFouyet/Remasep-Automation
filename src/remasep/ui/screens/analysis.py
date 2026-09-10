"""Pantalla 'Análisis': progreso comprensible mientras se procesa Medinet.

No muestra logs. Al terminar navega al Resumen mensual; ante un error muestra un
mensaje humano con acción para elegir otro archivo.
"""

from __future__ import annotations

from PySide6.QtCore import QThread
from PySide6.QtWidgets import (
    QHBoxLayout,
    QLabel,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.components.widgets import Card, ErrorBanner
from remasep.ui.errors import HumanError
from remasep.ui.workers import ANALYSIS_STEPS, AnalysisWorker, compute_medinet_summary

_STEPS = ["Datos", "Análisis", "Resumen", "Informe"]


class AnalysisScreen(QWidget):
    name = "analysis"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app
        self._thread: QThread | None = None
        self._worker: AnalysisWorker | None = None

        self.steps = StepIndicator(_STEPS)
        heading = QLabel("Analizando los datos de Medinet")
        heading.setProperty("role", "h2")

        self._progress_card = Card()
        self._status_label = QLabel("Preparando…")
        self._status_label.setProperty("role", "h3")
        self._bar = QProgressBar()
        self._bar.setRange(0, len(ANALYSIS_STEPS))
        self._bar.setTextVisible(False)
        self._detail = QLabel("Esto puede tardar unos segundos.")
        self._detail.setProperty("role", "muted")
        self._progress_card.body.addWidget(self._status_label)
        self._progress_card.body.addWidget(self._bar)
        self._progress_card.body.addWidget(self._detail)

        self._error = ErrorBanner()
        self._error.actionClicked.connect(lambda: self._app.navigate("new_report"))

        self.back_button = QPushButton("Cancelar")
        self.back_button.setProperty("variant", "ghost")
        self.back_button.clicked.connect(self._cancel)

        actions = QHBoxLayout()
        actions.addWidget(self.back_button)
        actions.addStretch()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(48, 28, 48, 24)
        layout.setSpacing(16)
        layout.addWidget(self.steps)
        layout.addWidget(heading)
        layout.addWidget(self._error)
        layout.addWidget(self._progress_card)
        layout.addStretch()
        layout.addLayout(actions)

    # --- ciclo de vida --------------------------------------------

    def on_enter(self) -> None:
        self.steps.set_current(1)
        self._error.clear()
        self._progress_card.setVisible(True)
        self._bar.setValue(0)
        self._status_label.setText("Preparando…")
        self._start()

    def _start(self) -> None:
        state = self._app.state
        self._thread = QThread(self)
        self._worker = AnalysisWorker(state.medinet_path, state.period)
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.step.connect(self._on_step)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._thread.start()

    def _teardown(self) -> None:
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait(2000)
            self._thread = None
            self._worker = None

    def _cancel(self) -> None:
        self._teardown()
        self._app.navigate("new_report")

    # --- señales del worker --------------------------------------

    def _on_step(self, text: str, index: int, total: int) -> None:
        self._status_label.setText(text)
        self._bar.setValue(min(index + 1, total))

    def _on_done(self, summary) -> None:
        self._teardown()
        self._app.state.summary = summary
        self._app.state.analysis_error = None
        self._app.navigate("dashboard")

    def _on_failed(self, human: HumanError) -> None:
        self._teardown()
        self._app.state.summary = None
        self._app.state.analysis_error = human
        self._progress_card.setVisible(False)
        self._error.show_error(human.title, human.detail, human.action_label)

    # --- ruta síncrona para tests (sin hilos) --------------------

    def run_now(self) -> None:
        """Ejecuta el análisis de forma bloqueante (usado por los tests)."""
        state = self._app.state
        try:
            summary = compute_medinet_summary(state.medinet_path, state.period)
        except Exception as exc:  # noqa: BLE001
            from remasep.ui.errors import humanize_error

            self._on_failed(humanize_error(exc))
        else:
            self._on_done(summary)
