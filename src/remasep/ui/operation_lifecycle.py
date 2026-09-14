"""Ownership and authorization rules for background Qt operations.

The manager owns every worker/thread pair until ``QThread.finished``.  Cancelling
an operation revokes its permission to update UI state; it never terminates the
underlying thread.
"""

from __future__ import annotations

import uuid
from collections.abc import Callable
from dataclasses import dataclass
from enum import Enum
from pathlib import Path

from PySide6.QtCore import QObject, QThread, Signal, Slot

from remasep.services.common import Period


class OperationKind(str, Enum):
    ANALYSIS = "analysis"
    GENERATION = "generation"


class OperationStatus(str, Enum):
    RUNNING = "running"
    CANCELLED = "cancelled"
    SUCCEEDED = "succeeded"
    FAILED = "failed"


class StartResult(str, Enum):
    STARTED = "started"
    ALREADY_RUNNING = "already_running"


@dataclass(frozen=True)
class OperationSnapshot:
    """Immutable identity and inputs captured before a worker starts."""

    run_id: str
    operation_type: OperationKind
    period: Period
    input_path: Path | None
    template_path: Path | None
    output_path: Path | None

    @classmethod
    def create(
        cls,
        *,
        operation_type: OperationKind,
        period: Period,
        input_path: str | Path | None,
        template_path: str | Path | None = None,
        output_path: str | Path | None = None,
    ) -> OperationSnapshot:
        def frozen_path(value: str | Path | None) -> Path | None:
            return None if value is None else Path(value)

        return cls(
            run_id=uuid.uuid4().hex,
            operation_type=operation_type,
            period=period,
            input_path=frozen_path(input_path),
            template_path=frozen_path(template_path),
            output_path=frozen_path(output_path),
        )


DoneCallback = Callable[[OperationSnapshot, object], None]
FailedCallback = Callable[[OperationSnapshot, object], None]
StepCallback = Callable[[OperationSnapshot, str, int, int], None]
WorkerFactory = Callable[[OperationSnapshot], QObject]


@dataclass
class _ManagedOperation:
    operation: OperationSnapshot
    thread: QThread
    worker: QObject
    on_step: StepCallback | None
    on_done: DoneCallback | None
    on_failed: FailedCallback | None
    status: OperationStatus = OperationStatus.RUNNING
    authorized: bool = True


class OperationManager(QObject):
    """Central owner for worker lifetimes and stale-signal authorization."""

    operation_finished = Signal(str, str)  # run_id, final status
    all_operations_finished = Signal()

    def __init__(self, parent: QObject | None = None) -> None:
        super().__init__(parent)
        self._operations: dict[str, _ManagedOperation] = {}
        self._current: dict[OperationKind, str] = {}

    @property
    def has_running_operations(self) -> bool:
        return bool(self._operations)

    def has_running(self, operation_type: OperationKind) -> bool:
        return any(
            managed.operation.operation_type is operation_type
            for managed in self._operations.values()
        )

    def start(
        self,
        operation: OperationSnapshot,
        worker_factory: WorkerFactory,
        *,
        on_step: StepCallback | None = None,
        on_done: DoneCallback | None = None,
        on_failed: FailedCallback | None = None,
    ) -> StartResult:
        """Start once allowed, without constructing a rejected worker.

        Generation stays exclusive until its backend thread has really ended,
        including after cancellation. Analysis can be superseded after its
        result has been revoked, while the old thread remains manager-owned.
        """
        current = self._current.get(operation.operation_type)
        generation_busy = (
            operation.operation_type is OperationKind.GENERATION
            and self.has_running(OperationKind.GENERATION)
        )
        if current is not None or generation_busy:
            return StartResult.ALREADY_RUNNING

        worker = worker_factory(operation)
        thread = QThread(self)
        thread.setProperty("remasep_run_id", operation.run_id)
        managed = _ManagedOperation(
            operation=operation,
            thread=thread,
            worker=worker,
            on_step=on_step,
            on_done=on_done,
            on_failed=on_failed,
        )
        self._operations[operation.run_id] = managed
        self._current[operation.operation_type] = operation.run_id

        worker.moveToThread(thread)
        thread.started.connect(worker.run)
        worker.step.connect(self._on_step)
        worker.done.connect(self._on_done)
        worker.failed.connect(self._on_failed)
        worker.done.connect(worker.deleteLater)
        worker.failed.connect(worker.deleteLater)
        worker.done.connect(thread.quit)
        worker.failed.connect(thread.quit)
        thread.finished.connect(self._on_thread_finished)
        thread.start()
        return StartResult.STARTED

    def cancel(self, run_id: str) -> bool:
        """Revoke result authorization without interrupting backend execution."""
        managed = self._operations.get(run_id)
        if managed is None or managed.status is not OperationStatus.RUNNING:
            return False
        managed.authorized = False
        managed.status = OperationStatus.CANCELLED
        kind = managed.operation.operation_type
        if self._current.get(kind) == run_id:
            self._current.pop(kind, None)
        return True

    def cancel_current(self, operation_type: OperationKind) -> bool:
        run_id = self._current.get(operation_type)
        return False if run_id is None else self.cancel(run_id)

    def cancel_all(self) -> None:
        for run_id in tuple(self._operations):
            self.cancel(run_id)

    @Slot(str, str, int, int)
    def _on_step(self, run_id: str, text: str, index: int, total: int) -> None:
        managed = self._authorized_sender(run_id)
        if managed is not None and managed.on_step is not None:
            managed.on_step(managed.operation, text, index, total)

    @Slot(str, object)
    def _on_done(self, run_id: str, value: object) -> None:
        managed = self._operations.get(run_id)
        if managed is None:
            return
        managed.thread.quit()
        authorized = self._is_authorized(managed)
        if authorized:
            managed.status = OperationStatus.SUCCEEDED
            self._clear_current(managed)
            if managed.on_done is not None:
                managed.on_done(managed.operation, value)

    @Slot(str, object)
    def _on_failed(self, run_id: str, error: object) -> None:
        managed = self._operations.get(run_id)
        if managed is None:
            return
        managed.thread.quit()
        authorized = self._is_authorized(managed)
        if authorized:
            managed.status = OperationStatus.FAILED
            self._clear_current(managed)
            if managed.on_failed is not None:
                managed.on_failed(managed.operation, error)

    @Slot()
    def _on_thread_finished(self) -> None:
        thread = self.sender()
        if not isinstance(thread, QThread):
            return
        run_id = str(thread.property("remasep_run_id"))
        managed = self._operations.pop(run_id, None)
        if managed is None:
            return
        self._clear_current(managed)
        thread.deleteLater()
        self.operation_finished.emit(run_id, managed.status.value)
        if not self._operations:
            self.all_operations_finished.emit()

    def _authorized_sender(self, run_id: str) -> _ManagedOperation | None:
        managed = self._operations.get(run_id)
        if managed is None:
            return None
        return managed if self._is_authorized(managed) else None

    def _is_authorized(self, managed: _ManagedOperation) -> bool:
        operation = managed.operation
        return (
            managed.authorized
            and managed.status is OperationStatus.RUNNING
            and self._current.get(operation.operation_type) == operation.run_id
        )

    def _clear_current(self, managed: _ManagedOperation) -> None:
        operation = managed.operation
        if self._current.get(operation.operation_type) == operation.run_id:
            self._current.pop(operation.operation_type, None)
