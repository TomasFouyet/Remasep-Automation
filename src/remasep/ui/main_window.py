"""Ventana principal: navegación por QStackedWidget + estado compartido.

El modo **demo** usa ``MockRemasepService`` (datos ficticios). El modo **real**
usa ``MedinetAnalysisService`` sobre un archivo Medinet seleccionado. Ambos
coexisten. No se genera REMASEP, no se usa COM ni macros.
"""

from __future__ import annotations

import datetime as _dt
import sys
from dataclasses import dataclass, field
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMainWindow, QStackedWidget, QWidget

from remasep.core.errors import RemasepError
from remasep.services.medinet_analysis import MedinetAnalysisResult, MedinetAnalysisService
from remasep.services.mock_remasep import (
    AnalysisResult,
    MockRemasepService,
    Period,
    ReviewOutcome,
    project_review,
)
from remasep.ui.screens.analysis import AnalysisScreen
from remasep.ui.screens.exceptions import ExceptionsScreen
from remasep.ui.screens.home import HomeScreen
from remasep.ui.screens.new_report import NewReportScreen
from remasep.ui.screens.summary import SummaryScreen
from remasep.ui.styles import STYLESHEET


def _default_period() -> tuple[int, int]:
    """Mes anterior al actual, según la fecha local del sistema."""
    today = _dt.date.today()  # noqa: DTZ011 - fecha civil local, granularidad de mes
    previous = today.replace(day=1) - _dt.timedelta(days=1)
    return previous.month, previous.year


ANALYSIS_MODE_DEMO = "DEMO"
ANALYSIS_MODE_REAL = "REAL"


@dataclass
class AppState:
    month: int
    year: int
    demo: bool = False
    medinet_path: Path | None = None
    egresos_path: Path | None = None
    recursos_path: Path | None = None
    analysis_mode: str = ""
    analysis: AnalysisResult | None = None  # solo modo DEMO
    medinet_analysis: MedinetAnalysisResult | None = None  # solo modo REAL
    analysis_error: str | None = None
    exception_decisions: dict[str, str] = field(default_factory=dict)

    @property
    def period(self) -> Period:
        return Period(self.month, self.year)

    @property
    def is_real(self) -> bool:
        return self.analysis_mode == ANALYSIS_MODE_REAL

    def review_outcome(self) -> ReviewOutcome | None:
        """Proyección del análisis tras aplicar las decisiones de la sesión (solo DEMO)."""
        if self.analysis is None:
            return None
        return project_review(self.analysis, self.exception_decisions)


class MainWindow(QMainWindow):
    def __init__(self, service: MockRemasepService | None = None) -> None:
        super().__init__()
        self.setWindowTitle("REMASEP Automático — Fundación Gantz")
        self.resize(880, 720)
        self.setStyleSheet(STYLESHEET)

        self.service = service or MockRemasepService()
        self.medinet_service = MedinetAnalysisService()
        month, year = _default_period()
        self.state = AppState(month=month, year=year)

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.screens: dict[str, QWidget] = {}
        for screen_cls in (
            HomeScreen,
            NewReportScreen,
            AnalysisScreen,
            ExceptionsScreen,
            SummaryScreen,
        ):
            screen = screen_cls(self)
            self.screens[screen.name] = screen
            self.stack.addWidget(screen)

        self.navigate("home")

    # --- navegación --------------------------------------------------

    def navigate(self, name: str) -> None:
        screen = self.screens[name]
        self.stack.setCurrentWidget(screen)
        on_enter = getattr(screen, "on_enter", None)
        if callable(on_enter):
            on_enter()

    @property
    def current_screen_name(self) -> str:
        return self.stack.currentWidget().name

    # --- acciones del flujo ---------------------------------------

    def run_analysis(self) -> None:
        """Ejecuta el análisis según haya archivo Medinet (REAL) o no (DEMO)."""
        self.state.exception_decisions = {}
        self.state.analysis = None
        self.state.medinet_analysis = None
        self.state.analysis_error = None

        if self.state.demo or self.state.medinet_path is None:
            self.state.analysis_mode = ANALYSIS_MODE_DEMO
            self.state.analysis = self.service.analyze(
                self.state.period,
                demo=True,
                medinet_file=self.state.medinet_path,
                egresos_file=self.state.egresos_path,
                recursos_file=self.state.recursos_path,
            )
            return

        self.state.analysis_mode = ANALYSIS_MODE_REAL
        try:
            self.state.medinet_analysis = self.medinet_service.analyze(
                self.state.medinet_path, self.state.period
            )
        except RemasepError as exc:
            self.state.analysis_error = str(exc)

    def start_analysis(self, *, demo: bool) -> str:
        """Lanza el análisis y devuelve el nombre de la pantalla a mostrar."""
        self.state.demo = demo
        self.run_analysis()
        return "analysis"

    def reset_flow(self) -> None:
        """Vuelve al inicio conservando el período elegido."""
        self.state = AppState(month=self.state.month, year=self.state.year)
        for selector_name in ("medinet_selector", "egresos_selector", "recursos_selector"):
            selector = getattr(self.screens["new_report"], selector_name, None)
            if selector is not None:
                selector.clear()
        self.navigate("home")


def run_app() -> None:
    app = QApplication.instance() or QApplication(sys.argv or ["remasep"])
    app.setApplicationName("REMASEP Automático")
    window = MainWindow()
    window.show()
    app.exec()
