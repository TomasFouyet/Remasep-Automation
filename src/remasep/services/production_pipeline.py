"""Path de PRODUCCIÓN Medinet → PendingWrite (Sprint 3.8).

Construye ``MetricValue`` y ``PendingWrite`` a partir de **sólo**:

1. el export directo de Medinet (input real del usuario),
2. los assets de runtime versionados (``config/runtime_2026/``),
3. la plantilla oficial (la aporta el ``GenerationService`` aguas abajo).

**No** abre ``GENERACION DATOS REMASEP.xlsx``. **No** lee ``artifacts/``. **No**
aplica ningún filtro por ESTADO (sigue ``PENDING_FUNCTIONAL_CONFIRMATION``). El
camino de reverse-engineering / diagnóstico vive aparte en
``scripts/build_metric_values.py`` (dev-only).
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

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
from remasep.services.runtime_assets import RuntimeBundle, load_runtime_bundle


@dataclass
class ProductionResult:
    period: str
    scope_records: int
    bundle_version: str
    template_fingerprint_id: str
    zero_write_policy: str
    estado_filter_status: str
    run: ProducerRun
    pending_writes: list
    completeness: CompletenessReport

    @property
    def write_ready_values(self) -> int:
        return len(self.run.metric_values)


def build_production_metric_values(
    medinet_path: str | Path,
    period: Period,
    *,
    bundle: RuntimeBundle | None = None,
    sheet_name: str | int | None = None,
) -> tuple[ProducerRun, RuntimeBundle]:
    """Corre el productor en modo PRODUCCIÓN usando el runtime bundle.

    Registros = ``processing_scope_records`` (válidos ∩ período), **sin** filtro
    por ESTADO.
    """
    bundle = bundle or load_runtime_bundle()
    frame = processing_scope_frame(medinet_path, period, sheet_name=sheet_name)
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
    return run, bundle


def build_production_pending_writes(
    medinet_path: str | Path,
    period: Period,
    *,
    bundle: RuntimeBundle | None = None,
    sheet_name: str | int | None = None,
) -> ProductionResult:
    run, bundle = build_production_metric_values(
        medinet_path, period, bundle=bundle, sheet_name=sheet_name
    )
    pending = join_metric_values_with_manifest(run.metric_values, bundle.write_instructions)
    completeness = check_write_completeness(run.metric_values, bundle.write_instructions)
    return ProductionResult(
        period=period.label,
        scope_records=run.scope_records,
        bundle_version=bundle.version,
        template_fingerprint_id=bundle.template_fingerprint_id,
        zero_write_policy=bundle.zero_write_policy,
        estado_filter_status=bundle.estado_filter_status,
        run=run,
        pending_writes=pending,
        completeness=completeness,
    )
