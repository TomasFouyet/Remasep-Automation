"""Contratos de concurrencia de las operaciones Qt (F04).

La sincronización usa ``threading.Event`` y señales Qt; no depende de sleeps.
"""

from __future__ import annotations

import os
import subprocess
import sys
import textwrap
import threading
from pathlib import Path

from PySide6.QtCore import QEventLoop, QObject, QTimer, Signal, Slot
from PySide6.QtTest import QSignalSpy

from remasep.services.common import Period
from remasep.ui.main_window import AppState
from remasep.ui.operation_lifecycle import (
    OperationKind,
    OperationManager,
    OperationSnapshot,
    StartResult,
)


class _EventWorker(QObject):
    step = Signal(str, str, int, int)
    done = Signal(str, object)
    failed = Signal(str, object)

    def __init__(
        self,
        operation: OperationSnapshot,
        started: threading.Event,
        release: threading.Event,
        payload: object,
        invocations: list[str],
        observed: list[tuple[Period, Path | None, Path | None, Path | None]] | None = None,
    ) -> None:
        super().__init__()
        self.operation = operation
        self.started = started
        self.release = release
        self.payload = payload
        self.invocations = invocations
        self.observed = observed

    @Slot()
    def run(self) -> None:
        self.invocations.append(self.operation.run_id)
        self.started.set()
        if not self.release.wait(5):
            self.failed.emit(self.operation.run_id, RuntimeError("test release timeout"))
            return
        if self.observed is not None:
            self.observed.append(
                (
                    self.operation.period,
                    self.operation.input_path,
                    self.operation.template_path,
                    self.operation.output_path,
                )
            )
        self.done.emit(self.operation.run_id, self.payload)


class _LateFailureWorker(QObject):
    step = Signal(str, str, int, int)
    done = Signal(str, object)
    failed = Signal(str, object)

    def __init__(
        self,
        operation: OperationSnapshot,
        started: threading.Event,
        release: threading.Event,
    ) -> None:
        super().__init__()
        self.operation = operation
        self.started = started
        self.release = release

    @Slot()
    def run(self) -> None:
        self.started.set()
        if not self.release.wait(5):
            return
        self.step.emit(self.operation.run_id, "stale progress", 1, 1)
        self.failed.emit(self.operation.run_id, RuntimeError("stale failure"))


def _operation(kind: OperationKind, month: int, tmp_path: Path) -> OperationSnapshot:
    return OperationSnapshot.create(
        operation_type=kind,
        period=Period(month, 2026),
        input_path=tmp_path / f"input-{month}.xlsx",
        template_path=tmp_path / f"template-{month}.xlsm",
        output_path=tmp_path / f"output-{month}.xlsm",
    )


def _wait_for(manager: OperationManager, spy: QSignalSpy, count: int = 1) -> None:
    while spy.count() < count:
        loop = QEventLoop()
        timed_out = []

        def timeout(loop=loop, timed_out=timed_out) -> None:
            timed_out.append(True)
            loop.quit()

        timer = QTimer()
        timer.setSingleShot(True)
        timer.timeout.connect(timeout)
        manager.operation_finished.connect(loop.quit)
        timer.start(5000)
        loop.exec()
        manager.operation_finished.disconnect(loop.quit)
        assert not timed_out, "la señal Qt esperada no llegó"


def test_cancelled_july_result_cannot_replace_current_june_state(qapp, tmp_path):
    manager = OperationManager()
    finished = QSignalSpy(manager.operation_finished)
    invocations: list[str] = []
    applied: list[tuple[Period, object]] = []
    state = AppState(month=7, year=2026)

    def apply_result(operation: OperationSnapshot, value: object) -> None:
        state.month = operation.period.month
        state.year = operation.period.year
        state.summary = value  # type: ignore[assignment] - payload controlado del worker fake
        applied.append((operation.period, value))

    july = _operation(OperationKind.ANALYSIS, 7, tmp_path)
    july_started, july_release = threading.Event(), threading.Event()
    assert manager.start(
        july,
        lambda op: _EventWorker(
            op, july_started, july_release, "summary-july", invocations
        ),
        on_done=apply_result,
    ) is StartResult.STARTED
    assert july_started.wait(5)

    assert manager.cancel(july.run_id)
    state.month = 6
    state.year = 2026
    state.summary = "summary-june"  # type: ignore[assignment]
    june = _operation(OperationKind.ANALYSIS, 6, tmp_path)
    june_started, june_release = threading.Event(), threading.Event()
    assert manager.start(
        june,
        lambda op: _EventWorker(
            op, june_started, june_release, "summary-june", invocations
        ),
        on_done=apply_result,
    ) is StartResult.STARTED
    assert june_started.wait(5)
    june_release.set()
    _wait_for(manager, finished)

    july_release.set()
    _wait_for(manager, finished, 2)
    assert applied == [(Period(6, 2026), "summary-june")]
    assert state.period == Period(6, 2026)
    assert state.summary == "summary-june"
    assert not manager.has_running_operations


def test_generation_guard_rejects_second_backend_invocation(qapp, tmp_path):
    manager = OperationManager()
    finished = QSignalSpy(manager.operation_finished)
    invocations: list[str] = []
    first = _operation(OperationKind.GENERATION, 7, tmp_path)
    started, release = threading.Event(), threading.Event()

    assert manager.start(
        first,
        lambda op: _EventWorker(op, started, release, "first", invocations),
    ) is StartResult.STARTED
    assert started.wait(5)

    second = _operation(OperationKind.GENERATION, 6, tmp_path)
    second_factory_called = threading.Event()

    def second_factory(op):
        second_factory_called.set()
        return _EventWorker(op, threading.Event(), threading.Event(), "second", invocations)

    assert manager.start(second, second_factory) is StartResult.ALREADY_RUNNING
    assert not second_factory_called.is_set()
    assert invocations == [first.run_id]

    release.set()
    _wait_for(manager, finished)
    assert not manager.has_running_operations


def test_cancelled_progress_and_failure_signals_are_ignored(qapp, tmp_path):
    manager = OperationManager()
    finished = QSignalSpy(manager.operation_finished)
    operation = _operation(OperationKind.ANALYSIS, 7, tmp_path)
    started, release = threading.Event(), threading.Event()
    progress: list[str] = []
    failures: list[object] = []

    assert manager.start(
        operation,
        lambda op: _LateFailureWorker(op, started, release),
        on_step=lambda _op, text, _index, _total: progress.append(text),
        on_failed=lambda _op, error: failures.append(error),
    ) is StartResult.STARTED
    assert started.wait(5)
    assert manager.cancel(operation.run_id)
    release.set()
    _wait_for(manager, finished)

    assert progress == []
    assert failures == []


def test_operation_parameters_are_immutable_and_cancel_discards_result(qapp, tmp_path):
    manager = OperationManager()
    finished = QSignalSpy(manager.operation_finished)
    original = _operation(OperationKind.GENERATION, 7, tmp_path)
    started, release = threading.Event(), threading.Event()
    observed: list[tuple[Period, Path | None, Path | None, Path | None]] = []
    applied: list[object] = []
    mutable_state = {
        "period": original.period,
        "input": original.input_path,
        "template": original.template_path,
        "output": original.output_path,
    }

    assert manager.start(
        original,
        lambda op: _EventWorker(
            op, started, release, "late", [], observed
        ),
        on_done=lambda _op, value: applied.append(value),
    ) is StartResult.STARTED
    assert started.wait(5)

    mutable_state.update(
        period=Period(6, 2026),
        input=tmp_path / "other.xlsx",
        template=tmp_path / "other.xlsm",
        output=tmp_path / "other-output.xlsm",
    )
    assert manager.cancel(original.run_id)
    release.set()
    _wait_for(manager, finished)

    assert observed == [
        (
            Period(7, 2026),
            tmp_path / "input-7.xlsx",
            tmp_path / "template-7.xlsm",
            tmp_path / "output-7.xlsm",
        )
    ]
    assert applied == []


def test_close_with_blocked_qthread_exits_without_abort(tmp_path):
    script = textwrap.dedent(
        """
        import threading

        from PySide6.QtCore import QObject, Signal, Slot
        from PySide6.QtWidgets import QApplication

        from remasep.services.common import Period
        from remasep.ui.main_window import MainWindow
        from remasep.ui.operation_lifecycle import (
            OperationKind, OperationSnapshot, StartResult,
        )

        started = threading.Event()
        release = threading.Event()
        close_seen = threading.Event()

        class Worker(QObject):
            step = Signal(str, str, int, int)
            done = Signal(str, object)
            failed = Signal(str, object)

            def __init__(self, operation):
                super().__init__()
                self.operation = operation

            @Slot()
            def run(self):
                started.set()
                if not release.wait(5):
                    self.failed.emit(self.operation.run_id, RuntimeError("timeout"))
                    return
                self.done.emit(self.operation.run_id, "done")

        class Window(MainWindow):
            def closeEvent(self, event):
                close_seen.set()
                super().closeEvent(event)

        class Bridge(QObject):
            request_close = Signal()

        app = QApplication([])
        window = Window()
        bridge = Bridge()
        bridge.request_close.connect(window.close)
        operation = OperationSnapshot.create(
            operation_type=OperationKind.GENERATION,
            period=Period(7, 2026),
            input_path="input.xlsx",
            template_path="template.xlsm",
            output_path="output.xlsm",
        )
        assert window.operations.start(operation, Worker) is StartResult.STARTED

        def coordinate():
            if not started.wait(5):
                release.set()
                bridge.request_close.emit()
                return
            bridge.request_close.emit()
            close_seen.wait(5)
            release.set()

        threading.Thread(target=coordinate, daemon=True).start()
        window.show()
        exit_code = app.exec()
        assert not window.operations.has_running_operations
        raise SystemExit(exit_code)
        """
    )
    env = dict(os.environ)
    env["QT_QPA_PLATFORM"] = "offscreen"
    src = str(Path(__file__).resolve().parents[1] / "src")
    env["PYTHONPATH"] = os.pathsep.join(filter(None, (src, env.get("PYTHONPATH"))))
    completed = subprocess.run(
        [sys.executable, "-c", script],
        cwd=tmp_path,
        env=env,
        text=True,
        capture_output=True,
        timeout=15,
        check=False,
    )
    assert completed.returncode == 0, completed.stdout + completed.stderr
