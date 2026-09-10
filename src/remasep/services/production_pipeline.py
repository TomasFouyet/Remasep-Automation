"""Path de PRODUCCIÓN Medinet → PendingWrite (Sprint 3.8 · filtro ESTADO 3.9 fase 2).

Construye ``MetricValue`` y ``PendingWrite`` a partir de **sólo**:

1. el export directo de Medinet (input real del usuario),
2. los assets de runtime versionados (``config/runtime_2026/``),
3. la plantilla oficial (la aporta el ``GenerationService`` aguas abajo).

**No** abre ``GENERACION DATOS REMASEP.xlsx``. **No** lee ``artifacts/``.

Filtro por ESTADO: la regla vive en ``config/runtime_2026/estado_filter.yaml``
(``status: CONFIRMED``, `included_states`). Se aplica **después** de limitar a
filas estructuralmente válidas ∩ período y **antes** de calcular los
``MetricValue``. Para Julio 2026: 2114 → 1364 registros.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from remasep.services.common import Period
from remasep.services.legacy_rules import load_legacy_rules
from remasep.services.medinet_analysis import processing_scope_frame
from remasep.services.metric_value_producer import (
    MODE_PRODUCTION,
    CompletenessReport,
    MetricValueProducer,
    ProducerRun,
    build_rows,
    check_write_completeness,
    join_metric_values_with_manifest,
)
from remasep.services.runtime_assets import (
    EstadoFilterRule,
    RuntimeBundle,
    load_runtime_bundle,
)


@dataclass(frozen=True)
class ProductionScope:
    """Desglose del alcance de procesamiento (auditable)."""

    period_scope_records: int          # estructuralmente válidos ∩ período (pre-ESTADO)
    estado_filter_status: str          # CONFIRMED / PENDING_FUNCTIONAL_CONFIRMATION
    estado_filter_applied: bool
    estado_included_states: tuple[str, ...]
    estado_excluded_records: int
    processing_scope_records: int       # post-ESTADO (lo que se procesa)


@dataclass
class ProductionResult:
    period: str
    scope: ProductionScope
    bundle_version: str
    template_fingerprint_id: str
    zero_write_policy: str
    run: ProducerRun
    pending_writes: list
    completeness: CompletenessReport

    @property
    def scope_records(self) -> int:
        return self.scope.processing_scope_records

    @property
    def estado_filter_status(self) -> str:
        return self.scope.estado_filter_status

    @property
    def write_ready_values(self) -> int:
        return len(self.run.metric_values)


def apply_estado_filter(
    frame: pd.DataFrame, rule: EstadoFilterRule
) -> tuple[pd.DataFrame, int]:
    """Filtra por ESTADO según la regla versionada. Devuelve ``(frame, excluidos)``.

    Si la regla no está ``CONFIRMED`` (o no tiene estados) devuelve el frame sin
    tocar. La normalización (`str.strip().casefold()`) la define la propia regla:
    no se repite la lista de estados ni la normalización en otro lugar.
    """
    if not rule.confirmed:
        return frame, 0
    included = rule.included_normalized()
    normalized = frame["ESTADO"].astype("string").str.strip().str.casefold()
    keep = normalized.isin(included).fillna(False).astype(bool)
    filtered = frame.loc[keep].reset_index(drop=True)
    return filtered, int(len(frame) - len(filtered))


def build_production_metric_values(
    medinet_path: str | Path,
    period: Period,
    *,
    bundle: RuntimeBundle | None = None,
    sheet_name: str | int | None = None,
) -> tuple[ProducerRun, RuntimeBundle, ProductionScope]:
    """Corre el productor en modo PRODUCCIÓN usando el runtime bundle.

    Registros = ``processing_scope_records`` = (válidos ∩ período) ∩ estados
    confirmados.
    """
    bundle = bundle or load_runtime_bundle()
    period_frame = processing_scope_frame(medinet_path, period, sheet_name=sheet_name)
    frame, excluded = apply_estado_filter(period_frame, bundle.estado_filter)

    scope = ProductionScope(
        period_scope_records=len(period_frame),
        estado_filter_status=bundle.estado_filter.status,
        estado_filter_applied=bundle.estado_filter.confirmed,
        estado_included_states=tuple(bundle.estado_filter.included_states),
        estado_excluded_records=excluded,
        processing_scope_records=len(frame),
    )

    rows = build_rows(frame.to_dict("records"), bundle.detail_columns_map, load_legacy_rules())
    producer = MetricValueProducer(
        formula_index=bundle.formula_index,
        detail_sheet=bundle.detail_sheet,
        resolver_by_sheet=bundle.resolver_by_sheet(),
        rows=rows,
        period=period,
        scope_records=len(frame),
        mode=MODE_PRODUCTION,
    )
    source_ids = [w.source_metric_id for w in bundle.write_instructions]
    run = producer.produce_run(source_ids, expected_types=bundle.expected_value_types())
    return run, bundle, scope


def build_production_pending_writes(
    medinet_path: str | Path,
    period: Period,
    *,
    bundle: RuntimeBundle | None = None,
    sheet_name: str | int | None = None,
) -> ProductionResult:
    run, bundle, scope = build_production_metric_values(
        medinet_path, period, bundle=bundle, sheet_name=sheet_name
    )
    pending = join_metric_values_with_manifest(run.metric_values, bundle.write_instructions)
    completeness = check_write_completeness(run.metric_values, bundle.write_instructions)
    return ProductionResult(
        period=period.label,
        scope=scope,
        bundle_version=bundle.version,
        template_fingerprint_id=bundle.template_fingerprint_id,
        zero_write_policy=bundle.zero_write_policy,
        run=run,
        pending_writes=pending,
        completeness=completeness,
    )
