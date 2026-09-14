"""Modelo de datos del **resumen mensual Medinet** (Sprint 3.10).

Adaptador de **sólo lectura** sobre el pipeline validado: no recalcula reglas
REMASEP, no modifica nada del backend. Reutiliza `processing_scope_frame`
(válidos ∩ período) y `apply_estado_filter` (regla ESTADO confirmada, Sprint
3.9), y calcula distribuciones **agregadas** para alimentar la UI y el PDF.

**Sin PII**: sólo conteos por categoría. Nunca nombres, RUT, fechas de
nacimiento individuales, teléfonos ni identificadores personales.

La UI y el export PDF consumen el mismo :class:`MonthlyMedinetSummary`.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path

import pandas as pd

from remasep.services.common import Period
from remasep.services.legacy_transform import legacy_age_years
from remasep.services.medinet_analysis import select_period_records
from remasep.services.production_pipeline import apply_estado_filter
from remasep.services.runtime_assets import RuntimeBundle, load_runtime_bundle

_SERVICE_TOP_N = 8
_MISSING = "(sin dato)"

# Tramos de edad para el dashboard (años). Amplios y legibles; la edad usa la
# regla legacy existente (`legacy_age_years`).
_AGE_BANDS: tuple[tuple[str, int, int], ...] = (
    ("0 a 4 años", 0, 5),
    ("5 a 9 años", 5, 10),
    ("10 a 14 años", 10, 15),
    ("15 a 19 años", 15, 20),
    ("20 a 39 años", 20, 40),
    ("40 a 64 años", 40, 65),
    ("65 años o más", 65, 10_000),
)

PENDING_SOURCES: tuple[str, ...] = (
    "Tabla quirúrgica",
    "Egresos hospitalarios",
    "Recursos / capacidad de pabellón",
)


@dataclass(frozen=True)
class CategoryCount:
    label: str
    count: int


@dataclass(frozen=True)
class EstadoCategoryCount:
    label: str
    count: int
    included: bool  # True = se considera para el REMASEP


@dataclass(frozen=True)
class MonthlyMedinetSummary:
    """Resumen agregado del mes, fuente **Medinet** únicamente. Sin PII, sin Qt."""

    period_label: str
    period_year: int
    period_month: int
    period_scope_records: int          # válidos + inválidos relevantes del período
    valid_records: int                  # válidos del período, antes de ESTADO
    invalid_records: int                # inválidos relevantes; nunca contabilizados
    problem_counts: tuple[tuple[str, int], ...]  # diagnóstico agregado, sin PII
    included_records: int              # tras la regla ESTADO confirmada
    excluded_records: int              # excluidos por ESTADO
    estado_distribution: tuple[EstadoCategoryCount, ...]
    sex_distribution: tuple[CategoryCount, ...]
    age_distribution: tuple[CategoryCount, ...]
    service_distribution: tuple[CategoryCount, ...]
    estado_included_states: tuple[str, ...]
    estado_excluded_states: tuple[str, ...]
    estado_filter_status: str
    generated_at: str
    pending_sources: tuple[str, ...] = field(default_factory=lambda: PENDING_SOURCES)
    source: str = "Medinet"

    @property
    def validation_blocked(self) -> bool:
        return self.invalid_records > 0

    @property
    def excluded_by_estado(self) -> int:
        return self.excluded_records

    @property
    def processing_scope_records(self) -> int:
        return self.included_records

    @property
    def included_percentage(self) -> float:
        if not self.period_scope_records:
            return 0.0
        return round(100.0 * self.included_records / self.period_scope_records, 1)

    @property
    def considered_ratio_label(self) -> str:
        return f"{self.included_records} / {self.period_scope_records}"

    def as_dict(self) -> dict:
        return {
            "source": self.source,
            "period": self.period_label,
            "period_scope_records": self.period_scope_records,
            "valid_records": self.valid_records,
            "invalid_records": self.invalid_records,
            "problem_counts": [
                {"code": code, "count": count} for code, count in self.problem_counts
            ],
            "validation_blocked": self.validation_blocked,
            "excluded_by_estado": self.excluded_by_estado,
            "processing_scope_records": self.processing_scope_records,
            "included_records": self.included_records,
            "excluded_records": self.excluded_records,
            "included_percentage": self.included_percentage,
            "estado_filter_status": self.estado_filter_status,
            "estado_included_states": list(self.estado_included_states),
            "estado_excluded_states": list(self.estado_excluded_states),
            "estado_distribution": [
                {"estado": s.label, "cantidad": s.count, "considerado": s.included}
                for s in self.estado_distribution
            ],
            "sex_distribution": [{"categoria": c.label, "cantidad": c.count} for c in self.sex_distribution],
            "age_distribution": [{"tramo": c.label, "cantidad": c.count} for c in self.age_distribution],
            "service_distribution": [
                {"especialidad": c.label, "cantidad": c.count} for c in self.service_distribution
            ],
            "pending_sources": list(self.pending_sources),
            "generated_at": self.generated_at,
        }


def _clean_series(frame: pd.DataFrame, column: str) -> pd.Series:
    values = frame[column].astype("string").str.strip()
    return values.mask(values.eq("") | values.isna(), _MISSING)


def _category_counts(frame: pd.DataFrame, column: str) -> tuple[CategoryCount, ...]:
    counts = _clean_series(frame, column).value_counts()
    return tuple(CategoryCount(str(label), int(n)) for label, n in counts.items())


def _service_counts(frame: pd.DataFrame) -> tuple[CategoryCount, ...]:
    counts = _clean_series(frame, "ESPECIALIDAD").value_counts()
    top = [CategoryCount(str(label), int(n)) for label, n in counts.head(_SERVICE_TOP_N).items()]
    rest = int(counts.iloc[_SERVICE_TOP_N:].sum())
    if rest:
        top.append(CategoryCount("Otras especialidades", rest))
    return tuple(top)


def _age_counts(frame: pd.DataFrame) -> tuple[CategoryCount, ...]:
    buckets: Counter[str] = Counter()
    for nac, cita in zip(frame["FECHA_NACIMIENTO"], frame["DIA_CITA"], strict=False):
        years = legacy_age_years(nac, cita)
        if years is None:
            buckets[_MISSING] += 1
            continue
        for label, low, high in _AGE_BANDS:
            if low <= years < high:
                buckets[label] += 1
                break
    ordered = [CategoryCount(label, buckets[label]) for label, _, _ in _AGE_BANDS if buckets.get(label)]
    if buckets.get(_MISSING):
        ordered.append(CategoryCount(_MISSING, buckets[_MISSING]))
    return tuple(ordered)


def build_monthly_medinet_summary(
    medinet_path: str | Path,
    period: Period,
    *,
    bundle: RuntimeBundle | None = None,
    sheet_name: str | int | None = None,
) -> MonthlyMedinetSummary:
    """Construye el resumen agregado del mes desde el export Medinet.

    Reutiliza exactamente las funciones del pipeline (`processing_scope_frame`,
    `apply_estado_filter`): los KPIs coinciden con
    ``build_production_pending_writes(...).scope``.
    """
    bundle = bundle or load_runtime_bundle()
    selection = select_period_records(medinet_path, period, sheet_name=sheet_name)
    period_frame = selection.frame
    included_frame, excluded = apply_estado_filter(period_frame, bundle.estado_filter)

    included_norm = bundle.estado_filter.included_normalized()
    estado_counts = _clean_series(period_frame, "ESTADO").value_counts()
    estado_dist = tuple(
        EstadoCategoryCount(
            label=str(label),
            count=int(n),
            included=str(label).strip().casefold() in included_norm,
        )
        for label, n in estado_counts.items()
    )

    return MonthlyMedinetSummary(
        period_label=period.label,
        period_year=period.year,
        period_month=period.month,
        period_scope_records=selection.period_scope_records,
        valid_records=selection.valid_records,
        invalid_records=selection.invalid_records,
        problem_counts=selection.problem_counts,
        included_records=len(included_frame),
        excluded_records=int(excluded),
        estado_distribution=estado_dist,
        sex_distribution=_category_counts(included_frame, "SEXO"),
        age_distribution=_age_counts(included_frame),
        service_distribution=_service_counts(included_frame),
        estado_included_states=tuple(bundle.estado_filter.included_states),
        estado_excluded_states=tuple(bundle.estado_filter.excluded_states),
        estado_filter_status=bundle.estado_filter.status,
        generated_at=datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
    )
