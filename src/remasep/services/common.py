"""Tipos compartidos por los servicios de análisis (mock y real)."""

from __future__ import annotations

from dataclasses import dataclass

_MONTHS_ES = [
    "Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio",
    "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre",
]


def month_name(month: int) -> str:
    return _MONTHS_ES[(month - 1) % 12]


@dataclass(frozen=True)
class Period:
    month: int
    year: int

    @property
    def label(self) -> str:
        return f"{month_name(self.month)} {self.year}"

    def contains(self, year: int, month: int) -> bool:
        return year == self.year and month == self.month


@dataclass(frozen=True)
class ValidationResult:
    name: str
    status: str  # "ok" | "warning" | "error"
    message: str = ""
