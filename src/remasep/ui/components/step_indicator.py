"""Indicador de pasos del asistente (1 Archivos · 2 Análisis · …)."""

from __future__ import annotations

from PySide6.QtWidgets import QHBoxLayout, QLabel, QWidget

from remasep.ui.styles import refresh_style


class StepIndicator(QWidget):
    def __init__(self, steps: list[str], parent: QWidget | None = None) -> None:
        super().__init__(parent)
        self._steps = list(steps)
        self._labels: list[QLabel] = []
        self._current = 0

        layout = QHBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)
        layout.setSpacing(8)
        for index, text in enumerate(self._steps):
            label = QLabel(f"{index + 1}  {text}")
            self._labels.append(label)
            layout.addWidget(label)
            if index < len(self._steps) - 1:
                sep = QLabel("→")
                sep.setProperty("role", "muted")
                layout.addWidget(sep)
        layout.addStretch()
        self.set_current(0)

    def set_current(self, index: int) -> None:
        self._current = index
        for i, label in enumerate(self._labels):
            if i < index:
                label.setObjectName("stepDone")
            elif i == index:
                label.setObjectName("stepActive")
            else:
                label.setObjectName("stepInactive")
            refresh_style(label)

    @property
    def current(self) -> int:
        return self._current
