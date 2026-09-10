"""Sprint 3.7A — Productor real de ``MetricValue`` para MEDINET.

Responde **qué valor** alimenta cada ``WriteInstruction`` del manifiesto de
escritura (Sprint 3.6). Reutiliza el motor de agregación legacy ya existente:

- parsing/evaluación de ``COUNTIF(S)(...) ± celda`` — :mod:`legacy_aggregation`;
- columnas derivadas ``AC:AL`` — :func:`legacy_transform.legacy_derived_record`;
- ruleset de clasificación legacy — :mod:`legacy_rules`;
- grafo de dependencias y fórmulas por celda — ``close_legacy_aggregations``
  (lo invoca el script; aquí sólo se recibe el índice ya extraído).

**No** reimplementa ninguna fórmula. **No** escribe Excel, **no** usa COM, **no**
modifica la plantilla. Sin PII: opera sobre agregados
(``processing_scope_records``) y produce conteos, nunca filas.

Dos modos, mutuamente excluyentes:

``PRODUCTION_PERIOD_SCOPE``
    Modo productivo. Registros = ``processing_scope_records`` (válidos ∩ período).
    **No** aplica ningún filtro por ESTADO. ``input_scope = PERIOD_ONLY``,
    ``estado_filter_applied = false``.

``LEGACY_EQUIVALENCE_DIAGNOSTIC``
    Modo diagnóstico, fuera del flujo productivo. Requiere opt-in explícito.
    Registros = ``processing_scope_records`` filtrados por la hipótesis ESTADO
    observada (``CURRENT_LEGACY_BEHAVIOR_PENDING_FUNCTIONAL_CONFIRMATION``).
    Sólo sirve para comparar contra el snapshot legacy de julio. **No** cambia
    la configuración productiva ni convierte la hipótesis en regla.
"""

from __future__ import annotations

import math
from collections import Counter
from collections.abc import Callable, Iterable, Mapping, Sequence
from dataclasses import dataclass, field

import pandas as pd

from remasep.adapters.medinet import OPTIONAL_FIELDS, REQUIRED_FIELDS
from remasep.core.errors import RemasepError
from remasep.services.common import Period
from remasep.services.legacy_aggregation import (
    DerivedFormula,
    UnsupportedFormulaError,
    parse_derived_formula,
)
from remasep.services.legacy_rules import LegacyRuleSet
from remasep.services.legacy_transform import legacy_derived_record
from remasep.services.writable_target_mapping import MetricValue

PRODUCER_VERSION = "medinet_value_producer_2026.v1"

MODE_PRODUCTION = "PRODUCTION_PERIOD_SCOPE"
MODE_DIAGNOSTIC = "LEGACY_EQUIVALENCE_DIAGNOSTIC"
_MODES = (MODE_PRODUCTION, MODE_DIAGNOSTIC)

INPUT_SCOPE_PERIOD_ONLY = "PERIOD_ONLY"
INPUT_SCOPE_PERIOD_AND_ESTADO_HYPOTHESIS = "PERIOD_AND_ESTADO_HYPOTHESIS"

# Campos semánticos del export (nunca letras de columna, nunca PII).
SEMANTIC_FIELDS: tuple[str, ...] = REQUIRED_FIELDS + OPTIONAL_FIELDS

# --- validación de tipo de valor (Sprint 3.7A §6) --------------------------
EXPECTED_INTEGER_COUNT = "INTEGER_COUNT"
EXPECTED_NUMERIC = "NUMERIC"
VALUE_TYPE_OK = "OK"
VALUE_TYPE_CONFLICT = "VALUE_TYPE_CONFLICT"

# --- estados de comparación de equivalencia (Sprint 3.7A §8) ---------------
CMP_MATCH = "MATCH"
CMP_ZERO_VS_BLANK = "ZERO_VS_BLANK_EQUIVALENT"
CMP_MISMATCH = "MISMATCH"
CMP_TARGET_UNAVAILABLE = "TARGET_UNAVAILABLE"
CMP_SOURCE_UNAVAILABLE = "SOURCE_UNAVAILABLE"

# --- semántica del cero (Sprint 3.7A §9) ----------------------------------
ZERO_SOURCE_ZERO_TARGET_ZERO = "SOURCE_ZERO_TARGET_ZERO"
ZERO_SOURCE_ZERO_TARGET_BLANK = "SOURCE_ZERO_TARGET_BLANK"
ZERO_SOURCE_ZERO_TARGET_OTHER = "SOURCE_ZERO_TARGET_OTHER"
ZERO_NOT_APPLICABLE = "NOT_APPLICABLE"

# --- política de escritura del cero (Sprint 3.7A §10) --------------------
ZERO_POLICY_WRITE = "WRITE_ZERO"
ZERO_POLICY_PRESERVE_BLANK = "PRESERVE_BLANK_FOR_ZERO"
ZERO_POLICY_FORM_SPECIFIC = "FORM_SPECIFIC"
ZERO_POLICY_UNRESOLVED = "UNRESOLVED"


class MetricValueProducerError(RemasepError):
    """Error del productor de valores (id desconocido, fórmula no soportada...)."""


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class LegacyFormulaSpec:
    """Fórmula legacy de una celda agregada, extraída del cierre de dependencias."""

    source_metric_id: str
    sheet: str
    cell: str
    kind: str
    formula: str
    value_ref_coords: tuple[str, ...] = ()


@dataclass(frozen=True)
class ValueTypeConflict:
    source_metric_id: str
    raw_value_repr: str
    reason: str


@dataclass(frozen=True)
class PendingWrite:
    """Unión ``MetricValue`` × ``WriteInstruction``. Todavía NO se escribe."""

    instruction_id: str
    source_metric_id: str
    target_sheet: str
    target_cell: str
    value: object
    expected_value_type: str
    period: str


@dataclass
class ProducerRun:
    mode: str
    period: str
    producer_version: str
    input_scope: str
    estado_filter_applied: bool
    diagnostic: bool
    scope_records: int
    metric_values: list[MetricValue] = field(default_factory=list)
    value_type_conflicts: list[ValueTypeConflict] = field(default_factory=list)
    unsupported: list[tuple[str, str]] = field(default_factory=list)

    @property
    def by_id(self) -> dict[str, MetricValue]:
        return {mv.source_metric_id: mv for mv in self.metric_values}

    def write_ready_values(self, write_instructions: Iterable) -> int:
        produced = self.by_id
        return sum(1 for wi in write_instructions if wi.source_metric_id in produced)


@dataclass(frozen=True)
class CompletenessReport:
    required_metric_count: int
    produced_metric_value_count: int
    missing: tuple[str, ...]
    duplicate: tuple[str, ...]
    orphan: tuple[str, ...]

    @property
    def ok(self) -> bool:
        return not (self.missing or self.duplicate or self.orphan)


# ---------------------------------------------------------------------------
# Construcción de filas (semántico -> letra de columna legacy + derivadas AC:AL)
# ---------------------------------------------------------------------------


def _isna(value: object) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def build_rows(
    records: Iterable[Mapping[str, object]],
    detail_columns_map: Mapping[str, str],
    ruleset: LegacyRuleSet,
) -> list[dict[str, object]]:
    """Filas listas para el evaluador legacy: ``{letra_columna: valor} + AC:AL``.

    ``detail_columns_map`` mapea campo semántico -> letra de columna **en la hoja
    de detalle legacy** (las fórmulas ``COUNTIF(S)`` referencian columnas por
    letra, p.ej. ``!$C:$C``). No entra ninguna columna PII: sólo los 10 campos
    semánticos y las 10 derivadas.
    """
    rows: list[dict[str, object]] = []
    for record in records:
        clean = {
            field_name: (None if _isna(record.get(field_name)) else record.get(field_name))
            for field_name in SEMANTIC_FIELDS
        }
        derived = legacy_derived_record(clean, ruleset)
        row: dict[str, object] = {
            letter: clean.get(field_name)
            for field_name, letter in detail_columns_map.items()
        }
        row.update(derived)
        rows.append(row)
    return rows


# ---------------------------------------------------------------------------
# Validación de tipo de valor
# ---------------------------------------------------------------------------


def validate_value_type(
    value: object, expected_value_type: str
) -> tuple[bool, object, str]:
    """``(ok, valor_coercionado, motivo)``. Sin truncamiento silencioso."""
    if expected_value_type == EXPECTED_INTEGER_COUNT:
        if isinstance(value, bool):
            return False, None, "BOOL_NOT_INTEGER_COUNT"
        if isinstance(value, int):
            return (True, value, VALUE_TYPE_OK) if value >= 0 else (False, None, "NEGATIVE_COUNT")
        if isinstance(value, float):
            if not math.isfinite(value):
                return False, None, "NON_FINITE_FLOAT"
            if value.is_integer():
                coerced = int(value)
                if coerced < 0:
                    return False, None, "NEGATIVE_COUNT"
                return True, coerced, "COERCED_EXACT_FLOAT"
            return False, None, "NON_INTEGER_FLOAT"
        return False, None, "NON_NUMERIC"
    if expected_value_type == EXPECTED_NUMERIC:
        if isinstance(value, bool):
            return False, None, "BOOL_NOT_NUMERIC"
        if isinstance(value, (int, float)) and math.isfinite(value):
            return True, value, VALUE_TYPE_OK
        return False, None, "NON_NUMERIC"
    # TEXT / UNKNOWN: sin coerción.
    return True, value, VALUE_TYPE_OK


# ---------------------------------------------------------------------------
# Productor
# ---------------------------------------------------------------------------


class MetricValueProducer:
    """Produce ``MetricValue`` por ``source_metric_id`` reevaluando la fórmula
    legacy contra las filas del alcance del período (o del subconjunto
    diagnóstico)."""

    def __init__(
        self,
        *,
        formula_index: Mapping[str, LegacyFormulaSpec],
        detail_sheet: str,
        resolver_by_sheet: Mapping[str, Callable[[str], object]],
        rows: Sequence[Mapping[str, object]],
        period: Period,
        scope_records: int | None = None,
        mode: str = MODE_PRODUCTION,
        producer_version: str = PRODUCER_VERSION,
    ) -> None:
        if mode not in _MODES:
            raise MetricValueProducerError(f"modo no soportado: {mode!r}")
        self._formula_index = dict(formula_index)
        self._detail_sheet = detail_sheet
        self._resolver_by_sheet = resolver_by_sheet
        self._rows = rows
        self._period = period
        self._mode = mode
        self._producer_version = producer_version
        self._scope_records = len(rows) if scope_records is None else scope_records
        self._coord_to_id: dict[tuple[str, str], str] = {
            (spec.sheet, spec.cell): mid for mid, spec in self._formula_index.items()
        }
        self._parsed: dict[str, DerivedFormula] = {}
        self._value_cache: dict[str, object] = {}
        self._evaluating: set[str] = set()

    # -- propiedades del modo ---------------------------------------------
    @property
    def mode(self) -> str:
        return self._mode

    @property
    def is_diagnostic(self) -> bool:
        return self._mode == MODE_DIAGNOSTIC

    @property
    def estado_filter_applied(self) -> bool:
        return self._mode == MODE_DIAGNOSTIC

    @property
    def input_scope(self) -> str:
        return (
            INPUT_SCOPE_PERIOD_AND_ESTADO_HYPOTHESIS
            if self._mode == MODE_DIAGNOSTIC
            else INPUT_SCOPE_PERIOD_ONLY
        )

    # -- evaluación ------------------------------------------------------
    def _parse(self, spec: LegacyFormulaSpec) -> DerivedFormula:
        cached = self._parsed.get(spec.source_metric_id)
        if cached is not None:
            return cached
        resolver = self._resolver_by_sheet.get(spec.sheet)
        if resolver is None:
            raise MetricValueProducerError(
                f"sin resolver de referencias para la hoja {spec.sheet!r}"
            )
        try:
            parsed = parse_derived_formula(spec.formula, self._detail_sheet, resolver)
        except UnsupportedFormulaError as exc:
            raise MetricValueProducerError(
                f"fórmula fuera del subset soportado para {spec.source_metric_id}: {exc}"
            ) from exc
        self._parsed[spec.source_metric_id] = parsed
        return parsed

    def _eval(self, spec: LegacyFormulaSpec) -> object:
        if spec.source_metric_id in self._value_cache:
            return self._value_cache[spec.source_metric_id]
        if spec.source_metric_id in self._evaluating:
            raise MetricValueProducerError(
                f"ciclo de dependencias en {spec.source_metric_id}"
            )
        self._evaluating.add(spec.source_metric_id)
        try:
            parsed = self._parse(spec)
            lookup: dict[str, float] = {}
            for coord in parsed.value_ref_coords:
                dep_id = self._coord_to_id.get((spec.sheet, coord))
                if dep_id is None:
                    raise MetricValueProducerError(
                        f"{spec.source_metric_id} depende de {spec.sheet}!{coord}, "
                        "ausente del índice de fórmulas"
                    )
                lookup[coord] = float(self._eval(self._formula_index[dep_id]))
            value = parsed.evaluate(self._rows, lookup)
        finally:
            self._evaluating.discard(spec.source_metric_id)
        self._value_cache[spec.source_metric_id] = value
        return value

    def produce(self, source_metric_id: str) -> MetricValue:
        spec = self._formula_index.get(source_metric_id)
        if spec is None:
            raise MetricValueProducerError(
                f"source_metric_id sin fórmula legacy: {source_metric_id}"
            )
        value = self._eval(spec)
        return MetricValue(
            source_metric_id=source_metric_id,
            value=value,
            period=self._period.label,
            producer_version=self._producer_version,
        )

    def produce_run(
        self,
        source_metric_ids: Iterable[str],
        *,
        expected_types: Mapping[str, str] | None = None,
    ) -> ProducerRun:
        run = ProducerRun(
            mode=self._mode,
            period=self._period.label,
            producer_version=self._producer_version,
            input_scope=self.input_scope,
            estado_filter_applied=self.estado_filter_applied,
            diagnostic=self.is_diagnostic,
            scope_records=self._scope_records,
        )
        seen: set[str] = set()
        for mid in source_metric_ids:
            if mid in seen:
                continue
            seen.add(mid)
            try:
                metric_value = self.produce(mid)
            except MetricValueProducerError as exc:
                run.unsupported.append((mid, str(exc)))
                continue
            expected = (expected_types or {}).get(mid, EXPECTED_INTEGER_COUNT)
            ok, coerced, reason = validate_value_type(metric_value.value, expected)
            if not ok:
                run.value_type_conflicts.append(
                    ValueTypeConflict(mid, repr(metric_value.value), reason)
                )
                continue
            run.metric_values.append(
                MetricValue(
                    source_metric_id=mid,
                    value=coerced,
                    period=metric_value.period,
                    producer_version=metric_value.producer_version,
                )
            )
        return run


# ---------------------------------------------------------------------------
# Unión con el manifiesto + completitud (Sprint 3.7A §16 / §17)
# ---------------------------------------------------------------------------


def join_metric_values_with_manifest(
    metric_values: Iterable[MetricValue],
    write_instructions: Iterable,
) -> list[PendingWrite]:
    """``MetricValue`` × ``WriteInstruction`` -> ``PendingWrite``. Sin last-one-wins.

    Lanza si un ``source_metric_id`` trae más de un ``MetricValue`` (ambigüedad de
    valor: hay que resolverla, no elegir el último).
    """
    by_id: dict[str, MetricValue] = {}
    for metric_value in metric_values:
        if metric_value.source_metric_id in by_id:
            raise MetricValueProducerError(
                "MetricValue duplicado para "
                f"{metric_value.source_metric_id} (no se aplica last-one-wins)"
            )
        by_id[metric_value.source_metric_id] = metric_value

    pending: list[PendingWrite] = []
    for instruction in write_instructions:
        metric_value = by_id.get(instruction.source_metric_id)
        if metric_value is None:
            continue
        pending.append(
            PendingWrite(
                instruction_id=instruction.instruction_id,
                source_metric_id=instruction.source_metric_id,
                target_sheet=instruction.target_sheet,
                target_cell=instruction.target_cell,
                value=metric_value.value,
                expected_value_type=instruction.expected_value_type,
                period=metric_value.period,
            )
        )
    return pending


def check_write_completeness(
    metric_values: Iterable[MetricValue],
    write_instructions: Iterable,
) -> CompletenessReport:
    """Cada instrucción WRITE_READY debe tener exactamente un ``MetricValue``."""
    required = {wi.source_metric_id for wi in write_instructions}
    counts = Counter(mv.source_metric_id for mv in metric_values)
    missing = tuple(sorted(m for m in required if m not in counts))
    duplicate = tuple(sorted(m for m, c in counts.items() if c > 1))
    orphan = tuple(sorted(m for m in counts if m not in required))
    return CompletenessReport(
        required_metric_count=len(required),
        produced_metric_value_count=int(sum(counts.values())),
        missing=missing,
        duplicate=duplicate,
        orphan=orphan,
    )


# ---------------------------------------------------------------------------
# Equivalencia de valores de referencia (Sprint 3.7A §7 / §8 / §9)
# ---------------------------------------------------------------------------


def _is_zero(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and value == 0


def _numeric_equal(source: object, target: object) -> bool:
    if isinstance(source, bool) or isinstance(target, bool):
        return source is target
    if isinstance(source, (int, float)) and isinstance(target, (int, float)):
        return float(source) == float(target)
    return source == target


def compare_reference_values(
    source_value: object,
    target_value: object,
    *,
    target_present: bool,
) -> tuple[str, str]:
    """``(comparison_status, zero_semantics)`` para un par (legacy, plantilla final).

    No se asume ``0 == blanco``: se devuelve el estado explícito
    ``ZERO_VS_BLANK_EQUIVALENT`` y se clasifica la semántica del cero para que la
    política se derive de la evidencia, no al revés.
    """
    if source_value is None:
        return CMP_SOURCE_UNAVAILABLE, ZERO_NOT_APPLICABLE
    if not target_present:
        if _is_zero(source_value):
            return CMP_ZERO_VS_BLANK, ZERO_SOURCE_ZERO_TARGET_BLANK
        return CMP_TARGET_UNAVAILABLE, ZERO_NOT_APPLICABLE
    if _numeric_equal(source_value, target_value):
        if _is_zero(source_value):
            return CMP_MATCH, ZERO_SOURCE_ZERO_TARGET_ZERO
        return CMP_MATCH, ZERO_NOT_APPLICABLE
    if _is_zero(source_value):
        return CMP_MISMATCH, ZERO_SOURCE_ZERO_TARGET_OTHER
    return CMP_MISMATCH, ZERO_NOT_APPLICABLE
