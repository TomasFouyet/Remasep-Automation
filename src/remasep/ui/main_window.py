"""Ventana principal: navegación por ``QStackedWidget`` + estado compartido.

Flujo MEDINET (Sprint 3.10):

    Inicio → Nuevo informe → Análisis → Resumen mensual → Generar → Resultado

Consume el backend **validado** (`medinet_summary`, `production_pipeline`,
`GenerationService`); no lo modifica.
"""

from __future__ import annotations

import datetime as _dt
import sys
from dataclasses import dataclass
from pathlib import Path

from PySide6.QtWidgets import QApplication, QMainWindow, QStackedWidget, QWidget

from remasep.services.common import Period
from remasep.services.medinet_summary import MonthlyMedinetSummary
from remasep.ui.errors import HumanError
from remasep.ui.screens.analysis import AnalysisScreen
from remasep.ui.screens.dashboard import DashboardScreen
from remasep.ui.screens.generate import GenerateScreen
from remasep.ui.screens.home import HomeScreen
from remasep.ui.screens.new_report import NewReportScreen
from remasep.ui.styles import STYLESHEET
from remasep.ui.workers import GenerationOutcome

APP_NAME = "REMASEP Automation"
APP_VERSION = "0.1.0"


def _default_period() -> tuple[int, int]:
    today = _dt.date.today()  # noqa: DTZ011 - fecha civil local, granularidad de mes
    previous = today.replace(day=1) - _dt.timedelta(days=1)
    return previous.month, previous.year


@dataclass
class AppState:
    month: int
    year: int
    medinet_path: Path | None = None
    template_path: Path | None = None
    output_path: Path | None = None
    summary: MonthlyMedinetSummary | None = None
    analysis_error: HumanError | None = None
    generation: GenerationOutcome | None = None

    @property
    def period(self) -> Period:
        return Period(self.month, self.year)


class MainWindow(QMainWindow):
    def __init__(self) -> None:
        super().__init__()
        self.setWindowTitle(f"{APP_NAME} — Fundación Gantz")
        self.resize(1040, 720)
        self.setMinimumSize(880, 600)
        self.setStyleSheet(STYLESHEET)

        month, year = _default_period()
        self.state = AppState(month=month, year=year)

        self.stack = QStackedWidget()
        self.setCentralWidget(self.stack)

        self.screens: dict[str, QWidget] = {}
        for screen_cls in (
            HomeScreen,
            NewReportScreen,
            AnalysisScreen,
            DashboardScreen,
            GenerateScreen,
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

    # --- flujo ------------------------------------------------------

    def reset_flow(self) -> None:
        """Vuelve al inicio conservando el período elegido."""
        month, year = self.state.month, self.state.year
        self.state = AppState(month=month, year=year)
        new_report = self.screens["new_report"]
        for attr in ("medinet_selector", "template_selector"):
            selector = getattr(new_report, attr, None)
            if selector is not None:
                selector.clear()
        self.navigate("home")


def run_app() -> None:  # pragma: no cover - punto de entrada
    app = QApplication.instance() or QApplication(sys.argv or ["remasep"])
    app.setApplicationName(APP_NAME)
    window = MainWindow()
    window.show()
    app.exec()
