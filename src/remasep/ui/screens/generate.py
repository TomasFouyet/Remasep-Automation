"""Pantalla 'Generar REMASEP': progreso → resultado.

Usa el ``GenerationService`` existente (Excel COM). No reimplementa nada. Al
terminar ofrece abrir el Excel y su carpeta con APIs seguras del sistema.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QThread
from PySide6.QtWidgets import (
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QProgressBar,
    QPushButton,
    QVBoxLayout,
    QWidget,
)

from remasep.ui.components.step_indicator import StepIndicator
from remasep.ui.components.widgets import Card, ErrorBanner, StatusBadge
from remasep.ui.errors import HumanError
from remasep.ui.open_location import open_containing_folder, open_path
from remasep.ui.styles import format_int
from remasep.ui.workers import GENERATION_STEPS, GenerationOutcome, GenerationWorker, run_generation

_STEPS = ["Datos", "Análisis", "Resumen", "Informe"]
_NOT_FOR_SUBMISSION = (
    "Este informe contiene los datos automatizados desde Medinet. Algunas "
    "secciones del REMASEP todavía pueden requerir información adicional "
    "(tabla quirúrgica, egresos, recursos)."
)


class GenerateScreen(QWidget):
    name = "generate"

    def __init__(self, app: QWidget) -> None:
        super().__init__()
        self._app = app
        self._thread: QThread | None = None
        self._worker: GenerationWorker | None = None
        self._output_path: str | None = None
        self._last_human: HumanError | None = None

        self.steps = StepIndicator(_STEPS)
        self._heading = QLabel("Generando el REMASEP")
        self._heading.setProperty("role", "h2")

        # progreso
        self._progress_card = Card()
        self._status_label = QLabel("Preparando…")
        self._status_label.setProperty("role", "h3")
        self._bar = QProgressBar()
        self._bar.setRange(0, len(GENERATION_STEPS))
        self._bar.setTextVisible(False)
        self._progress_hint = QLabel("Microsoft Excel realizará el cálculo. No cierres la aplicación.")
        self._progress_hint.setProperty("role", "muted")
        self._progress_card.body.addWidget(self._status_label)
        self._progress_card.body.addWidget(self._bar)
        self._progress_card.body.addWidget(self._progress_hint)

        # error
        self._error = ErrorBanner()
        self._error.actionClicked.connect(self._retry_or_back)

        # resultado
        self._result_card = Card()
        ok_row = QHBoxLayout()
        self._result_title = QLabel("REMASEP generado correctamente")
        self._result_title.setProperty("role", "h2")
        ok_row.addWidget(StatusBadge("✓ LISTO", kind="ok"))
        ok_row.addWidget(self._result_title, stretch=1)
        self._result_card.body.addLayout(ok_row)
        self._result_lines = QVBoxLayout()
        self._result_lines.setSpacing(4)
        self._result_card.body.addLayout(self._result_lines)
        self._nfs = QLabel(_NOT_FOR_SUBMISSION)
        self._nfs.setProperty("role", "muted")
        self._nfs.setWordWrap(True)
        self._result_card.body.addSpacing(4)
        self._result_card.body.addWidget(self._nfs)

        self._open_excel = QPushButton("Abrir Excel")
        self._open_excel.setProperty("variant", "primary")
        self._open_excel.clicked.connect(self._do_open_excel)
        self._open_folder = QPushButton("Abrir carpeta")
        self._open_folder.clicked.connect(self._do_open_folder)
        self._home = QPushButton("Volver al inicio")
        self._home.setProperty("variant", "ghost")
        self._home.clicked.connect(self._app.reset_flow)
        result_actions = QHBoxLayout()
        result_actions.addWidget(self._open_excel)
        result_actions.addWidget(self._open_folder)
        result_actions.addStretch()
        result_actions.addWidget(self._home)
        self._result_card.body.addSpacing(6)
        self._result_card.body.addLayout(result_actions)

        self.back_button = QPushButton("Volver al resumen")
        self.back_button.setProperty("variant", "ghost")
        self.back_button.clicked.connect(lambda: self._app.navigate("dashboard"))
        bottom = QHBoxLayout()
        bottom.addWidget(self.back_button)
        bottom.addStretch()

        layout = QVBoxLayout(self)
        layout.setContentsMargins(48, 28, 48, 24)
        layout.setSpacing(16)
        layout.addWidget(self.steps)
        layout.addWidget(self._heading)
        layout.addWidget(self._error)
        layout.addWidget(self._progress_card)
        layout.addWidget(self._result_card)
        layout.addStretch()
        layout.addLayout(bottom)

    # --- ciclo de vida --------------------------------------------

    def on_enter(self) -> None:
        self.steps.set_current(3)
        self._error.clear()
        self._result_card.setVisible(False)
        self._progress_card.setVisible(False)

        state = self._app.state
        # "Guardar como" ANTES de iniciar Excel: si no hay ruta, o la ruta ya
        # existe (p. ej. una segunda generación del mismo mes), pedirla.
        if state.output_path is None or Path(state.output_path).exists():
            target = self._prompt_output_path()
            if target is None:  # el usuario canceló: volver al Resumen, sin error
                self._app.navigate("dashboard")
                return
            state.output_path = target

        self._progress_card.setVisible(True)
        self._heading.setText("Generando el REMASEP")
        self._bar.setValue(0)
        self._status_label.setText("Preparando…")
        self._start()

    def _prompt_output_path(self) -> Path | None:
        """Diálogo "Guardar informe REMASEP". Devuelve la ruta o ``None`` si se cancela.

        No sobrescribe: si el nombre elegido ya existe, lo explica y vuelve a
        pedir. Nunca borra archivos existentes.
        """
        period = self._app.state.period
        default_name = f"REMASEP_{period.year:04d}_{period.month:02d}_DRAFT.xlsm"
        previous = self._app.state.output_path
        start_dir = previous.parent if previous is not None else Path("outputs")
        if not start_dir.is_dir():
            start_dir = Path.home()
        suggested = str(start_dir / default_name)
        while True:
            path, _ = QFileDialog.getSaveFileName(
                self,
                "Guardar informe REMASEP",
                suggested,
                "Excel con macros (*.xlsm)",
                options=QFileDialog.Option.DontConfirmOverwrite,
            )
            if not path:
                return None
            target = Path(path)
            if target.suffix.lower() != ".xlsm":
                target = target.with_suffix(".xlsm")
            if target.exists():
                QMessageBox.warning(
                    self,
                    "Ya existe un archivo",
                    f"Ya existe un archivo llamado “{target.name}”. Elige otro nombre.",
                )
                suggested = str(target.parent / default_name)
                continue
            return target

    def _start(self) -> None:
        state = self._app.state
        out = str(state.output_path) if state.output_path is not None else None
        self._thread = QThread(self)
        self._worker = GenerationWorker(
            state.medinet_path, state.period, state.template_path, out
        )
        self._worker.moveToThread(self._thread)
        self._thread.started.connect(self._worker.run)
        self._worker.step.connect(self._on_step)
        self._worker.done.connect(self._on_done)
        self._worker.failed.connect(self._on_failed)
        self._thread.start()

    def _teardown(self) -> None:
        if self._thread is not None:
            self._thread.quit()
            self._thread.wait(5000)
            self._thread = None
            self._worker = None

    # --- señales ------------------------------------------------

    def _on_step(self, text: str, index: int, total: int) -> None:
        self._status_label.setText(text)
        self._bar.setValue(min(index + 1, total))

    def _on_done(self, outcome: GenerationOutcome) -> None:
        self._teardown()
        self._app.state.generation = outcome
        if not outcome.ok:
            self._show_error(outcome.human_error)
            return
        self._show_result(outcome)

    def _on_failed(self, human: HumanError) -> None:
        self._teardown()
        self._show_error(human)

    # --- vistas -----------------------------------------------

    def _show_error(self, human: HumanError | None) -> None:
        human = human or HumanError("No pudimos generar el informe.", "Vuelve a intentarlo.", "Reintentar")
        self._last_human = human
        self._progress_card.setVisible(False)
        self._result_card.setVisible(False)
        self._heading.setText("No se pudo generar el informe")
        self._error.show_error(human.title, human.detail, human.action_label)

    def _show_result(self, o: GenerationOutcome) -> None:
        self._progress_card.setVisible(False)
        self._heading.setText("Generación completada")
        self._output_path = o.output_path
        while self._result_lines.count():
            item = self._result_lines.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        name = Path(o.output_path).name if o.output_path else "—"
        rows = [
            ("Nombre del archivo", name),
            ("Período", self._app.state.period.label),
            ("Atenciones consideradas", format_int(o.considered_records)),
            ("Integridad del archivo", "Correcta" if o.integrity_ok else "Con observaciones"),
        ]
        for label, value in rows:
            line = QLabel(f"{label}:  {value}")
            line.setProperty("role", "muted")
            self._result_lines.addWidget(line)
        self._open_excel.setEnabled(bool(o.output_path))
        self._open_folder.setEnabled(bool(o.output_path))
        self._result_card.setVisible(True)

    # --- acciones -------------------------------------------

    def _retry_or_back(self) -> None:
        label = (self._last_human.action_label if self._last_human else "").lower()
        if "nombre" in label:
            # OUTPUT_ALREADY_EXISTS: reabrir "Guardar como" directamente, sin
            # perder Medinet / plantilla / período / análisis ya válidos.
            self._app.state.output_path = None
            self._app.navigate("generate")
        elif label.startswith("elegir otro archivo"):
            self._app.navigate("new_report")
        elif label.startswith("reintentar"):
            self._app.navigate("generate")
        else:
            self._app.navigate("dashboard")

    def _do_open_excel(self) -> None:
        if self._output_path:
            open_path(self._output_path)

    def _do_open_folder(self) -> None:
        if self._output_path:
            open_containing_folder(self._output_path)

    # --- ruta síncrona para tests --------------------------

    def run_now(self, output_path: str | Path | None = None) -> None:
        state = self._app.state
        out = output_path if output_path is not None else state.output_path
        outcome = run_generation(
            state.medinet_path,
            state.period,
            state.template_path,
            str(out) if out is not None else None,
        )
        self._on_done(outcome)
