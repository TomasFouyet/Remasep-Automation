"""Servicio mock para la maqueta de UI.

NO implementa lógica REMASEP real: devuelve datos ficticios y estables para que
la interfaz pueda recorrerse de principio a fin mientras se desarrolla el
backend. No lee Excel, no toca archivos de pacientes, no persiste nada.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass, field
from pathlib import Path

# Period / ValidationResult / month_name viven en un módulo neutral y se
# re-exportan aquí para no romper los imports existentes de la UI.
from remasep.services.common import Period, ValidationResult, month_name

__all__ = [
    "APP_VERSION",
    "CLASSIFICATION_OPTIONS",
    "IGNORE_OPTION",
    "TEMPLATE_VERSION",
    "AnalysisResult",
    "ExceptionItem",
    "MockRemasepService",
    "ModuleStatus",
    "Period",
    "ReviewOutcome",
    "SourceStatus",
    "ValidationResult",
    "month_name",
    "project_review",
]

TEMPLATE_VERSION = "REMASEP 2026 V1.4"
APP_VERSION = "0.1.0"

# Opción especial: marca los registros de una excepción como ignorados.
IGNORE_OPTION = "Ignorar justificadamente"

# Categorías que se ofrecen al reclasificar una excepción (solo visual, no se
# persisten ni modifican reglas reales).
CLASSIFICATION_OPTIONS = [
    "Consulta médica",
    "Control odontológico de especialidad",
    "Evaluación odontológica de especialidad",
    "Prestación odontológica",
    IGNORE_OPTION,
]


@dataclass(frozen=True)
class ModuleStatus:
    name: str
    detected: bool
    note: str = ""


@dataclass(frozen=True)
class SourceStatus:
    name: str
    available: bool
    note: str = ""


@dataclass(frozen=True)
class ExceptionItem:
    value: str
    count: int
    reason: str
    possible_categories: list[str] = field(default_factory=lambda: list(CLASSIFICATION_OPTIONS))


@dataclass(frozen=True)
class AnalysisResult:
    period: Period
    total_records: int
    classified_records: int
    ignored_records: int
    unclassified_records: int
    validations: list[ValidationResult]
    modules: list[ModuleStatus]
    other_sources: list[SourceStatus]
    exceptions: list[ExceptionItem]
    template_version: str = TEMPLATE_VERSION
    demo: bool = False

    def __post_init__(self) -> None:
        parts = self.classified_records + self.ignored_records + self.unclassified_records
        if parts != self.total_records:
            raise ValueError(
                "AnalysisResult inconsistente: "
                f"{self.classified_records} + {self.ignored_records} + "
                f"{self.unclassified_records} = {parts} != total_records={self.total_records}"
            )
        exceptions_total = sum(item.count for item in self.exceptions)
        if exceptions_total != self.unclassified_records:
            raise ValueError(
                f"Los ExceptionItem suman {exceptions_total}, se esperaba "
                f"unclassified_records={self.unclassified_records}"
            )

    @property
    def period_label(self) -> str:
        return self.period.label

    @property
    def review_groups(self) -> int:
        """Cantidad de tipos/grupos distintos que requieren revisión humana."""
        return len(self.exceptions)

    def pending(self, resolved_values: set[str] | None = None) -> int:
        """Grupos de excepción sin una decisión tomada en la UI (no se persiste nada)."""
        resolved = resolved_values or set()
        return sum(1 for item in self.exceptions if item.value not in resolved)

    def can_continue(self, resolved_values: set[str] | None = None) -> bool:
        return self.pending(resolved_values) == 0


@dataclass(frozen=True)
class ReviewOutcome:
    """Estado de los registros DESPUÉS de aplicar las decisiones de la sesión UI.

    No modifica ``AnalysisResult``: es una proyección derivada. Se mantiene la
    invariante ``total == final_classified + final_ignored + final_unclassified``.
    """

    total_records: int
    final_classified: int
    final_ignored: int
    final_unclassified: int
    resolved_groups: int
    pending_groups: int

    def __post_init__(self) -> None:
        parts = self.final_classified + self.final_ignored + self.final_unclassified
        if parts != self.total_records:
            raise ValueError(
                "ReviewOutcome inconsistente: "
                f"{self.final_classified} + {self.final_ignored} + "
                f"{self.final_unclassified} = {parts} != total_records={self.total_records}"
            )

    @property
    def fully_reviewed(self) -> bool:
        return self.pending_groups == 0


def project_review(
    analysis: AnalysisResult, decisions: Mapping[str, str] | None = None
) -> ReviewOutcome:
    """Reparte los registros de las excepciones según las decisiones tomadas.

    - Excepción con una categoría asignada -> sus registros pasan a *classified*.
    - Excepción marcada "Ignorar justificadamente" -> pasan a *ignored*.
    - Excepción sin decisión -> siguen contando como *unclassified* / pendientes.
    """
    decisions = decisions or {}
    extra_classified = 0
    extra_ignored = 0
    still_unclassified = 0
    resolved_groups = 0
    pending_groups = 0

    for item in analysis.exceptions:
        decision = (decisions.get(item.value) or "").strip()
        if not decision:
            still_unclassified += item.count
            pending_groups += 1
        elif decision == IGNORE_OPTION:
            extra_ignored += item.count
            resolved_groups += 1
        else:
            extra_classified += item.count
            resolved_groups += 1

    return ReviewOutcome(
        total_records=analysis.total_records,
        final_classified=analysis.classified_records + extra_classified,
        final_ignored=analysis.ignored_records + extra_ignored,
        final_unclassified=still_unclassified,
        resolved_groups=resolved_groups,
        pending_groups=pending_groups,
    )


class MockRemasepService:
    """Genera un ``AnalysisResult`` ficticio pero coherente para la maqueta."""

    template_version = TEMPLATE_VERSION

    def analyze(
        self,
        period: Period,
        *,
        demo: bool = False,
        medinet_file: Path | None = None,
        egresos_file: Path | None = None,
        recursos_file: Path | None = None,
    ) -> AnalysisResult:
        validations = [
            ValidationResult("Archivo leído correctamente", "ok"),
            ValidationResult("Estructura compatible", "ok"),
            ValidationResult("Período correcto", "ok"),
            ValidationResult("Fechas válidas", "ok"),
            ValidationResult("Clasificación ejecutada", "ok"),
        ]

        modules = [
            ModuleStatus("REMASEP 01", True),
            ModuleStatus("B2 ANEXO", True),
            ModuleStatus("REMASEP OD", True),
        ]

        other_sources = [
            SourceStatus("Egresos hospitalarios", False, "pendiente integración"),
            SourceStatus("Recursos / pabellones", False, "pendiente definición de fuente"),
        ]

        exceptions = [
            ExceptionItem(
                value="CONTROL ORTODONCIA ESPECIALIDAD NUEVA",
                count=7,
                reason="No existe una clasificación conocida.",
            ),
            ExceptionItem(
                value="EVALUACIÓN FONOAUDIOLÓGICA COMPLEMENTARIA",
                count=2,
                reason="No existe una clasificación conocida.",
            ),
            ExceptionItem(
                value="TELECONSULTA GENÉTICA (MODALIDAD SIN MAPEO)",
                count=1,
                reason="La modalidad observada no está mapeada.",
            ),
        ]

        # Cifras ficticias coherentes con la invariante de AnalysisResult:
        # total = clasificadas + ignoradas + sin_clasificar, y los ExceptionItem
        # suman exactamente las "sin clasificar" (7 + 2 + 1 = 10).
        total = 2006
        ignored = 8
        unclassified = sum(item.count for item in exceptions)
        classified = total - ignored - unclassified

        return AnalysisResult(
            period=period,
            total_records=total,
            classified_records=classified,
            ignored_records=ignored,
            unclassified_records=unclassified,
            validations=validations,
            modules=modules,
            other_sources=other_sources,
            exceptions=exceptions,
            demo=demo or medinet_file is None,
        )
