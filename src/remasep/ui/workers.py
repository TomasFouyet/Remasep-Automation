"""Trabajo en segundo plano de la UI: análisis Medinet y generación REMASEP.

Cada worker es un ``QObject`` que corre en un ``QThread`` y emite señales de
progreso legibles (nunca logs Python) + un resultado o un ``HumanError``.

La lógica pura vive en funciones (``compute_medinet_summary`` /
``run_generation``) reutilizables por los tests sin hilos ni Excel.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QObject, Signal

from remasep.core.errors import RemasepError
from remasep.services.common import Period
from remasep.services.excel_writer import (
    MODE_PRODUCTION,
    STATUS_GENERATED_DRAFT,
    WRITER_INTEGRITY_PASS,
)
from remasep.services.generation_service import GenerationService
from remasep.services.medinet_summary import (
    MonthlyMedinetSummary,
    build_monthly_medinet_summary,
)
from remasep.services.production_pipeline import build_production_pending_writes
from remasep.ui.errors import HumanError, humanize_error, humanize_generation_status

ANALYSIS_STEPS: tuple[str, ...] = (
    "Leyendo el archivo de Medinet…",
    "Validando el período…",
    "Aplicando los estados confirmados…",
    "Calculando las métricas del mes…",
    "Preparando el resumen…",
)

GENERATION_STEPS: tuple[str, ...] = (
    "Preparando la plantilla…",
    "Calculando los datos de Medinet…",
    "Escribiendo los datos en el Excel…",
    "Recalculando el Excel…",
    "Verificando la integridad del archivo…",
)


# ---------------------------------------------------------------------------
# Lógica pura (sin Qt)
# ---------------------------------------------------------------------------


def compute_medinet_summary(
    medinet_path: str | Path, period: Period
) -> MonthlyMedinetSummary:
    return build_monthly_medinet_summary(medinet_path, period)


@dataclass(frozen=True)
class GenerationOutcome:
    ok: bool
    status: str
    submission_label: str
    output_path: str | None
    written_cells: int
    considered_records: int
    integrity_ok: bool
    control_status: str
    warnings: tuple[str, ...]
    human_error: HumanError | None = None


def run_generation(
    medinet_path: str | Path,
    period: Period,
    template_path: str | Path,
    output_path: str | Path | None,
    *,
    service: GenerationService | None = None,
) -> GenerationOutcome:
    """Construye los PendingWrites y llama a ``GenerationService`` (Excel COM).

    Reutiliza el pipeline validado; no reimplementa nada. En Linux/sin Excel
    devuelve ``ok=False`` con un ``HumanError`` legible (no lanza).
    """
    from remasep.services.runtime_assets import load_runtime_bundle

    service = service or GenerationService()
    bundle = load_runtime_bundle()
    production = build_production_pending_writes(medinet_path, period, bundle=bundle)
    if production.run.value_type_conflicts or production.run.unsupported:
        return GenerationOutcome(
            ok=False, status="ABORTED_PREFLIGHT", submission_label="",
            output_path=None, written_cells=0,
            considered_records=production.scope.processing_scope_records,
            integrity_ok=False, control_status="", warnings=(),
            human_error=HumanError(
                "No pudimos preparar los datos de Medinet.",
                "Algunas métricas del mes no se pudieron calcular. "
                "Revisa el archivo de Medinet.",
                "Elegir otro archivo",
            ),
        )

    sr = service.generate(
        mode=MODE_PRODUCTION,
        template_path=template_path,
        output_path=output_path,
        pending_writes=production.pending_writes,
        manifest_instruction_ids=bundle.instruction_ids,
        period_year=period.year,
        period_month=period.month,
        zero_write_policy=production.zero_write_policy,
        expected_fingerprint_id=production.template_fingerprint_id,
    )
    r = sr.result
    ok = r.status in (STATUS_GENERATED_DRAFT, "GENERATED_DIAGNOSTIC_REFERENCE")
    human = None if ok else humanize_generation_status(r.status, errors=list(r.errors))
    return GenerationOutcome(
        ok=ok,
        status=r.status,
        submission_label=sr.submission_label,
        output_path=r.output_path,
        written_cells=r.written_cells,
        considered_records=production.scope.processing_scope_records,
        integrity_ok=r.writer_integrity_status == WRITER_INTEGRITY_PASS,
        control_status=r.control_status,
        warnings=tuple(r.warnings),
        human_error=human,
    )


# ---------------------------------------------------------------------------
# Workers (Qt)
# ---------------------------------------------------------------------------


class AnalysisWorker(QObject):
    step = Signal(str, int, int)          # texto, índice, total
    done = Signal(object)                 # MonthlyMedinetSummary
    failed = Signal(object)              # HumanError

    def __init__(self, medinet_path: str | Path, period: Period) -> None:
        super().__init__()
        self._medinet_path = medinet_path
        self._period = period

    def run(self) -> None:
        try:
            for i, label in enumerate(ANALYSIS_STEPS):
                self.step.emit(label, i, len(ANALYSIS_STEPS))
            summary = compute_medinet_summary(self._medinet_path, self._period)
        except RemasepError as exc:
            self.failed.emit(humanize_error(exc))
        except Exception as exc:  # noqa: BLE001 - a mensaje humano, sin traceback al usuario
            self.failed.emit(humanize_error(exc))
        else:
            self.step.emit("Listo", len(ANALYSIS_STEPS), len(ANALYSIS_STEPS))
            self.done.emit(summary)


class GenerationWorker(QObject):
    step = Signal(str, int, int)
    done = Signal(object)                 # GenerationOutcome
    failed = Signal(object)              # HumanError

    def __init__(
        self,
        medinet_path: str | Path,
        period: Period,
        template_path: str | Path,
        output_path: str | Path | None,
    ) -> None:
        super().__init__()
        self._args = (medinet_path, period, template_path, output_path)

    def run(self) -> None:
        try:
            for i, label in enumerate(GENERATION_STEPS):
                self.step.emit(label, i, len(GENERATION_STEPS))
            outcome = run_generation(*self._args)
        except RemasepError as exc:
            self.failed.emit(humanize_error(exc))
        except Exception as exc:  # noqa: BLE001
            self.failed.emit(humanize_error(exc))
        else:
            self.step.emit("Listo", len(GENERATION_STEPS), len(GENERATION_STEPS))
            self.done.emit(outcome)
