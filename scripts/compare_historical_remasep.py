"""Validación histórica del motor MEDINET (diagnóstico, sólo lectura).

Compara el pipeline PRODUCTIVO actual (``build_production_pending_writes``) contra
cuatro REMASEP históricos reales (Abril/Mayo/Junio/Julio 2026), usando el único
export directo de Medinet disponible (``detalle_citas``) como fuente para los
cuatro meses — no existen exports Medinet históricos separados.

**No** reimplementa la lógica MEDINET: reutiliza directamente
``build_production_pending_writes`` (Sprint 3.8/3.9) y
``compare_reference_values`` (Sprint 3.7A) tal cual están, sin tocarlas. **No**
modifica reglas de clasificación, el filtro ESTADO, los runtime assets, el
``write_manifest.csv`` ni ningún archivo histórico: los REMASEP de referencia se
abren siempre con ``openpyxl`` en modo ``data_only=True, read_only=True`` y se
verifica su SHA256 antes/después de leer.

Alcance de la comparación: **sólo** las 1122 celdas del ``write_manifest``
(la automatización basada en ``detalle_citas``). Egresos hospitalarios, tabla
quirúrgica, recursos/capacidad, etc. quedan explícitamente fuera de alcance.

Categorías de comparación (celda a celda). Un valor ``None``/vacío LEÍDO CON ÉXITO
(archivo existe, hoja existe, celda legible) **es** un valor de referencia válido
— no es lo mismo que "no se pudo obtener la referencia":

- ``EXACT_MATCH`` — valor calculado == valor histórico.
- ``ZERO_VS_BLANK_EQUIVALENT`` — calculado 0 y referencia vacía/None (o ambos 0).
- ``REAL_MISMATCH`` — ambos tienen valores comparables y son distintos,
  incluyendo calculado no-cero frente a referencia vacía/None (una celda vacía
  frente a un conteo esperado no-cero es una discrepancia real, no una
  referencia indisponible).
- ``REFERENCE_UNAVAILABLE`` — la referencia realmente no se pudo obtener:
  archivo ausente/ambiguo, hoja inexistente, celda ilegible o cualquier otro
  fallo real de lectura. Nunca se usa sólo porque la celda esté vacía.

Si algún invariante de producción falla (1122 PendingWrites, completitud OK,
filtro ESTADO=CONFIRMED) el mes se marca ``BLOCKED`` y no se fabrica ninguna
comparación para ese mes.

Uso::

    python scripts/compare_historical_remasep.py
    python scripts/compare_historical_remasep.py --months 7 --output /tmp/out
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sys
from collections import Counter, defaultdict
from collections.abc import Callable, Iterable
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path

import openpyxl

from remasep.core.errors import RemasepError
from remasep.services.common import Period
from remasep.services.legacy_aggregation import DerivedFormula, parse_derived_formula
from remasep.services.metric_value_producer import (
    CMP_MATCH,
    CMP_MISMATCH,
    CMP_SOURCE_UNAVAILABLE,
    CMP_TARGET_UNAVAILABLE,
    CMP_ZERO_VS_BLANK,
    compare_reference_values,
)
from remasep.services.production_pipeline import ProductionResult, build_production_pending_writes
from remasep.services.runtime_assets import RuntimeBundle, load_runtime_bundle

DEFAULT_OUTPUT = "artifacts/historical_regression"
DEFAULT_EXPORT = "data/local/detalle_citas - 2026-09-07T123630.940.xlsx"
DEFAULT_DATA_DIR = "data/local"
YEAR_IN_SCOPE = 2026
MONTHS_IN_SCOPE = (4, 5, 6, 7)

STATUS_EXACT_MATCH = "EXACT_MATCH"
STATUS_ZERO_VS_BLANK = "ZERO_VS_BLANK_EQUIVALENT"
STATUS_REAL_MISMATCH = "REAL_MISMATCH"
STATUS_REFERENCE_UNAVAILABLE = "REFERENCE_UNAVAILABLE"
ALL_STATUSES = (
    STATUS_EXACT_MATCH,
    STATUS_ZERO_VS_BLANK,
    STATUS_REAL_MISMATCH,
    STATUS_REFERENCE_UNAVAILABLE,
)

# Mapea el vocabulario ya existente de metric_value_producer al vocabulario de
# este diagnóstico (mismas categorías semánticas, sólo se relabelan) — SIEMPRE
# que la referencia se haya podido leer (ver `classify_comparison`): con la
# referencia disponible, `CMP_TARGET_UNAVAILABLE` de `compare_reference_values`
# significa "celda leída con éxito pero vacía/None frente a un calculado
# no-cero", que aquí es una discrepancia real (REAL_MISMATCH), NO una
# referencia indisponible — esa categoría se reserva para cuando la lectura en
# sí falla (archivo/hoja/celda ausente o inaccesible).
_CMP_STATUS_MAP: dict[str, str] = {
    CMP_MATCH: STATUS_EXACT_MATCH,
    CMP_ZERO_VS_BLANK: STATUS_ZERO_VS_BLANK,
    CMP_MISMATCH: STATUS_REAL_MISMATCH,
    CMP_TARGET_UNAVAILABLE: STATUS_REAL_MISMATCH,
    CMP_SOURCE_UNAVAILABLE: STATUS_REFERENCE_UNAVAILABLE,
}

MONTH_LABELS_ES = {1: "Enero", 2: "Febrero", 3: "Marzo", 4: "Abril", 5: "Mayo", 6: "Junio",
                    7: "Julio", 8: "Agosto", 9: "Septiembre", 10: "Octubre",
                    11: "Noviembre", 12: "Diciembre"}
MONTH_TOKENS_ES = {m: name.casefold() for m, name in MONTH_LABELS_ES.items()}

# Nombres solicitados explícitamente por el usuario para cada mes. Si el archivo
# real difiere ligeramente se intenta una detección acotada (ver
# `resolve_reference_file`), pero nunca se elige entre varios candidatos en
# silencio.
REQUESTED_REFERENCE_FILES: dict[int, str] = {
    4: "REMASEP 2026 V1.4 Abril 2026.xlsm",
    5: "2026-5 REMASEP_V1.4.xlsm",
    6: "REMASEP 2026 V1.4 Junio 2026.xlsm",
    7: "REMASEP_V1.4 Julio 2026.xlsm",
}

# Composición semántica documentada de las columnas derivadas AC:AL
# (`remasep.services.legacy_transform`, docstring del módulo). No se reimplementa
# el cálculo: sólo se usa para etiquetar de qué campos depende cada fórmula.
_DERIVED_SEMANTIC_FIELDS: dict[str, tuple[str, ...]] = {
    "AC": ("TIPO_DE_CITA", "SUCURSAL"),
    "AD": ("TIPO_DE_CITA", "SEXO"),
    "AE": ("PRESTACION", "ESPECIALIDAD", "SEXO"),
    "AF": ("FECHA_NACIMIENTO", "DIA_CITA"),
    "AG": ("TIPO_DE_CITA", "SEXO"),
    "AH": ("TIPO_DE_CITA", "SEXO"),
    "AI": ("TIPO_DE_CITA", "SEXO"),
    "AJ": ("PRESTACION", "SEXO"),
    "AK": ("PRESTACION", "SEXO"),
    "AL": ("PRESTACION", "SEXO"),
}
_AG_AL_COLUMNS = frozenset({"AG", "AH", "AI", "AJ", "AK", "AL"})
AGE_COLUMN = "AF"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# Resolución del archivo histórico (sin matching ambiguo silencioso)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ReferenceResolution:
    path: Path | None
    match_kind: str
    requested_name: str
    ignored_candidates: tuple[str, ...] = ()

    @property
    def available(self) -> bool:
        return self.path is not None


def resolve_reference_file(
    data_dir: Path, month: int, requested_name: str
) -> ReferenceResolution:
    """Localiza el REMASEP histórico de ``month`` bajo ``data_dir``.

    Orden: (1) nombre exacto solicitado; (2) nombre exacto salvo mayúsculas; (3)
    único archivo ``.xlsm`` cuyo nombre contiene el mes (en español) y "remasep".
    Si el paso (3) encuentra más de un candidato, se marca ``AMBIGUOUS`` y no se
    elige ninguno. Nunca se sobreescribe ni se abre en modo escritura.
    """
    exact = data_dir / requested_name
    if exact.is_file():
        return ReferenceResolution(exact, "EXACT_NAME_MATCH", requested_name)

    if not data_dir.is_dir():
        return ReferenceResolution(None, "DATA_DIR_MISSING", requested_name)

    candidates = sorted(
        p for p in data_dir.iterdir() if p.is_file() and p.suffix.casefold() == ".xlsm"
    )
    lower_target = requested_name.casefold()
    ci_matches = [p for p in candidates if p.name.casefold() == lower_target]
    if len(ci_matches) == 1:
        return ReferenceResolution(ci_matches[0], "CASE_INSENSITIVE_MATCH", requested_name)

    token = MONTH_TOKENS_ES[month]
    fuzzy = [
        p for p in candidates
        if token in p.name.casefold() and "remasep" in p.name.casefold()
    ]
    if len(fuzzy) == 1:
        return ReferenceResolution(fuzzy[0], "FUZZY_SINGLE_CANDIDATE", requested_name)
    if len(fuzzy) > 1:
        return ReferenceResolution(
            None, "AMBIGUOUS", requested_name, tuple(p.name for p in fuzzy)
        )
    return ReferenceResolution(None, "NOT_FOUND", requested_name)


# ---------------------------------------------------------------------------
# Lectura de celdas históricas (sólo lectura) — distingue "vacío pero leído"
# de "no se pudo obtener la referencia"
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CellRead:
    """Resultado de leer UNA celda del REMASEP histórico.

    ``available=True, value=None`` = la celda se leyó con éxito y está vacía:
    es un valor de referencia válido (blanco), no una referencia indisponible.
    ``available=False`` = la referencia realmente no se pudo obtener (hoja
    inexistente, celda ilegible u otro fallo real de lectura); ``reason``
    documenta el motivo.
    """

    available: bool
    value: object = None
    reason: str = ""


def read_target_cells(
    path: Path, sheet_cells: set[tuple[str, str]]
) -> dict[tuple[str, str], CellRead]:
    """Lee ``sheet_cells`` del REMASEP histórico en ``path``. Sólo lectura
    (``data_only=True, read_only=True``): nunca escribe, nunca recalcula.

    Cualquier fallo real (archivo corrupto, hoja ausente, celda ilegible) se
    captura por celda y se reporta como ``CellRead(available=False, ...)``, sin
    interrumpir la comparación del resto de las celdas ni del resto de los
    meses. Un valor ``None`` leído con éxito se conserva como
    ``CellRead(available=True, value=None)``.
    """
    try:
        wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    except Exception as exc:  # noqa: BLE001 - archivo corrupto/no soportado, no debe tumbar la corrida
        reason = f"WORKBOOK_OPEN_ERROR: {exc}"
        return dict.fromkeys(sheet_cells, CellRead(False, None, reason))

    try:
        present_sheets = set(wb.sheetnames)
        out: dict[tuple[str, str], CellRead] = {}
        for sheet, cell in sheet_cells:
            if sheet not in present_sheets:
                out[(sheet, cell)] = CellRead(False, None, "SHEET_NOT_FOUND")
                continue
            try:
                value = wb[sheet][cell].value
            except Exception as exc:  # noqa: BLE001 - celda/rango realmente inaccesible
                out[(sheet, cell)] = CellRead(False, None, f"CELL_READ_ERROR: {exc}")
                continue
            out[(sheet, cell)] = CellRead(True, value)
        return out
    finally:
        wb.close()


# ---------------------------------------------------------------------------
# Contexto diagnóstico por métrica (fórmula, columnas referenciadas, ejes)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class MetricContext:
    source_metric_id: str
    sheet: str
    cell: str
    formula: str
    kind: str
    referenced_columns: tuple[str, ...]
    depends_on_age_af: bool
    depends_on_tipo_de_cita: bool
    depends_on_prestacion: bool
    depends_on_especialidad: bool
    depends_on_sexo: bool
    depends_on_estado: bool
    depends_on_ag_al_rules: bool


def build_metric_contexts(bundle: RuntimeBundle) -> dict[str, MetricContext]:
    """Contexto diagnóstico (sin PII) por ``source_metric_id``: fórmula legacy,
    columnas de la hoja de detalle referenciadas de forma **transitiva** (a
    través de las dependencias entre celdas agregadas) y ejes de investigación.

    Sólo lectura del catálogo ya validado (``bundle.formula_index``); no
    recalcula ningún valor y no reimplementa el evaluador de fórmulas — usa
    ``parse_derived_formula`` (motor productivo) para obtener las columnas
    referenciadas de cada fórmula.
    """
    formula_index = bundle.formula_index
    resolver_by_sheet = bundle.resolver_by_sheet()
    coord_to_id = {(spec.sheet, spec.cell): mid for mid, spec in formula_index.items()}
    letter_to_field = {letter: name for name, letter in bundle.detail_columns_map.items()}

    parsed_cache: dict[str, DerivedFormula] = {}
    columns_cache: dict[str, frozenset[str]] = {}

    def parse(mid: str) -> DerivedFormula:
        cached = parsed_cache.get(mid)
        if cached is not None:
            return cached
        spec = formula_index[mid]
        resolver = resolver_by_sheet[spec.sheet]
        parsed = parse_derived_formula(spec.formula, bundle.detail_sheet, resolver)
        parsed_cache[mid] = parsed
        return parsed

    def transitive_columns(mid: str, seen: set[str]) -> frozenset[str]:
        cached = columns_cache.get(mid)
        if cached is not None:
            return cached
        if mid in seen:
            return frozenset()
        seen = seen | {mid}
        parsed = parse(mid)
        cols = set(parsed.referenced_columns)
        spec = formula_index[mid]
        for coord in parsed.value_ref_coords:
            dep_mid = coord_to_id.get((spec.sheet, coord))
            if dep_mid is not None:
                cols |= transitive_columns(dep_mid, seen)
        result = frozenset(cols)
        columns_cache[mid] = result
        return result

    def semantic_fields(letters: frozenset[str]) -> set[str]:
        fields: set[str] = set()
        for letter in letters:
            if letter in _DERIVED_SEMANTIC_FIELDS:
                fields.update(_DERIVED_SEMANTIC_FIELDS[letter])
            elif letter in letter_to_field:
                fields.add(letter_to_field[letter])
        return fields

    contexts: dict[str, MetricContext] = {}
    for mid, spec in formula_index.items():
        letters = transitive_columns(mid, set())
        fields = semantic_fields(letters)
        contexts[mid] = MetricContext(
            source_metric_id=mid,
            sheet=spec.sheet,
            cell=spec.cell,
            formula=spec.formula,
            kind=spec.kind,
            referenced_columns=tuple(sorted(letters)),
            depends_on_age_af=AGE_COLUMN in letters,
            depends_on_tipo_de_cita="TIPO_DE_CITA" in fields,
            depends_on_prestacion="PRESTACION" in fields,
            depends_on_especialidad="ESPECIALIDAD" in fields,
            depends_on_sexo="SEXO" in fields,
            depends_on_estado="ESTADO" in fields,
            depends_on_ag_al_rules=bool(letters & _AG_AL_COLUMNS),
        )
    return contexts


# ---------------------------------------------------------------------------
# Clasificación celda a celda (reutiliza compare_reference_values)
# ---------------------------------------------------------------------------


def classify_comparison(
    calculated_value: object, reference_value: object, *, reference_available: bool
) -> str:
    """Clasifica una celda en una de las 4 categorías del diagnóstico.

    ``reference_available=False`` significa que la referencia realmente no se
    pudo obtener (archivo/hoja ausente, celda ilegible u otro fallo real de
    lectura): siempre ``REFERENCE_UNAVAILABLE``, nunca se asume equivalencia con
    0. ``reference_available=True`` con ``reference_value=None`` significa que
    la celda se leyó con éxito y está vacía — es un valor de referencia válido,
    no una referencia indisponible. Delega en ``compare_reference_values``
    (motor productivo, sin reimplementar): un calculado 0 frente a una
    referencia vacía es ``ZERO_VS_BLANK_EQUIVALENT``; un calculado no-cero
    frente a una referencia vacía es ``REAL_MISMATCH``.
    """
    if not reference_available:
        return STATUS_REFERENCE_UNAVAILABLE
    target_present = reference_value is not None
    cmp_status, _zero_semantics = compare_reference_values(
        calculated_value, reference_value, target_present=target_present
    )
    return _CMP_STATUS_MAP[cmp_status]


# ---------------------------------------------------------------------------
# Ejecución por mes
# ---------------------------------------------------------------------------

_DETAIL_FIELDS = (
    "month", "year", "period_label", "instruction_id", "source_metric_id",
    "target_sheet", "target_cell", "expected_value_type", "calculated_value",
    "reference_value", "reference_available", "reference_unavailable_reason",
    "comparison_status", "delta",
    "source_formula", "source_legacy_sheet", "source_legacy_cell", "source_kind",
    "referenced_detail_columns", "depends_on_age_AF", "depends_on_tipo_de_cita",
    "depends_on_prestacion", "depends_on_especialidad", "depends_on_sexo",
    "depends_on_estado", "depends_on_ag_al_rules",
    "processing_scope_records", "period_scope_records", "estado_excluded_records",
    "reference_file", "reference_match_kind",
)


@dataclass
class MonthOutcome:
    month: int
    year: int
    period_label: str
    status: str  # "COMPARED" | "BLOCKED"
    block_reason: str
    reference: ReferenceResolution
    scope_processing_records: int
    scope_period_records: int
    scope_estado_excluded: int
    completeness_ok: bool
    expected_writes: int
    pending_write_count: int
    detail_rows: list[dict] = field(default_factory=list)


def _empty_context_row() -> dict:
    return {
        "source_formula": "", "source_legacy_sheet": "", "source_legacy_cell": "",
        "source_kind": "", "referenced_detail_columns": "",
        "depends_on_age_AF": False, "depends_on_tipo_de_cita": False,
        "depends_on_prestacion": False, "depends_on_especialidad": False,
        "depends_on_sexo": False, "depends_on_estado": False,
        "depends_on_ag_al_rules": False,
    }


def _context_row(ctx: MetricContext | None) -> dict:
    if ctx is None:
        return _empty_context_row()
    return {
        "source_formula": ctx.formula,
        "source_legacy_sheet": ctx.sheet,
        "source_legacy_cell": ctx.cell,
        "source_kind": ctx.kind,
        "referenced_detail_columns": "|".join(ctx.referenced_columns),
        "depends_on_age_AF": ctx.depends_on_age_af,
        "depends_on_tipo_de_cita": ctx.depends_on_tipo_de_cita,
        "depends_on_prestacion": ctx.depends_on_prestacion,
        "depends_on_especialidad": ctx.depends_on_especialidad,
        "depends_on_sexo": ctx.depends_on_sexo,
        "depends_on_estado": ctx.depends_on_estado,
        "depends_on_ag_al_rules": ctx.depends_on_ag_al_rules,
    }


def check_invariants(result: ProductionResult, expected_writes: int) -> list[str]:
    """Invariantes que DEBEN cumplirse antes de comparar. Si alguno falla, el mes
    se marca BLOCKED y no se fabrica ninguna comparación (celda a celda)."""
    failures: list[str] = []
    if len(result.pending_writes) != expected_writes:
        failures.append(
            f"pending_writes={len(result.pending_writes)} != esperado {expected_writes}"
        )
    if not result.completeness.ok:
        failures.append(
            "completeness no OK: "
            f"missing={len(result.completeness.missing)} "
            f"duplicate={len(result.completeness.duplicate)} "
            f"orphan={len(result.completeness.orphan)}"
        )
    if result.scope.estado_filter_status != "CONFIRMED" or not result.scope.estado_filter_applied:
        failures.append(
            f"estado_filter_status={result.scope.estado_filter_status!r} "
            f"applied={result.scope.estado_filter_applied}"
        )
    return failures


def run_month(
    export_path: Path,
    data_dir: Path,
    month: int,
    year: int,
    bundle: RuntimeBundle,
    contexts: dict[str, MetricContext],
    *,
    pipeline: Callable[..., ProductionResult] = build_production_pending_writes,
) -> MonthOutcome:
    """Ejecuta el pipeline PRODUCTIVO para ``(month, year)`` y compara sus 1122
    PendingWrites contra el REMASEP histórico correspondiente. Sólo lectura."""
    period = Period(month, year)
    requested_name = REQUESTED_REFERENCE_FILES[month]
    reference = resolve_reference_file(data_dir, month, requested_name)
    expected_writes = bundle.expected_instruction_count

    result = pipeline(export_path, period, bundle=bundle)
    invariant_failures = check_invariants(result, expected_writes)

    base = MonthOutcome(
        month=month, year=year, period_label=period.label,
        status="BLOCKED" if invariant_failures else "COMPARED",
        block_reason="; ".join(invariant_failures),
        reference=reference,
        scope_processing_records=result.scope.processing_scope_records,
        scope_period_records=result.scope.period_scope_records,
        scope_estado_excluded=result.scope.estado_excluded_records,
        completeness_ok=result.completeness.ok,
        expected_writes=expected_writes,
        pending_write_count=len(result.pending_writes),
    )
    if invariant_failures:
        return base

    cell_reads: dict[tuple[str, str], CellRead] = {}
    if reference.available:
        sha_before = _sha256(reference.path)
        sheet_cells = {(pw.target_sheet, pw.target_cell) for pw in result.pending_writes}
        cell_reads = read_target_cells(reference.path, sheet_cells)
        sha_after = _sha256(reference.path)
        if sha_before != sha_after:
            raise RemasepError(
                f"el archivo histórico cambió durante la lectura: {reference.path}"
            )

    detail_rows: list[dict] = []
    for pw in sorted(result.pending_writes, key=lambda p: p.instruction_id):
        cell_read = cell_reads.get((pw.target_sheet, pw.target_cell))
        if not reference.available:
            cell_read = CellRead(False, None, reference.match_kind)
        elif cell_read is None:
            # no debería ocurrir (se piden todas las celdas de este mes), pero
            # nunca se fabrica disponibilidad si por algún motivo falta la clave.
            cell_read = CellRead(False, None, "CELL_NOT_REQUESTED")

        ref_value = cell_read.value if cell_read.available else None
        status = classify_comparison(pw.value, ref_value, reference_available=cell_read.available)
        delta = None
        if isinstance(pw.value, (int, float)) and isinstance(ref_value, (int, float)):
            delta = ref_value - pw.value

        row = {
            "month": month,
            "year": year,
            "period_label": period.label,
            "instruction_id": pw.instruction_id,
            "source_metric_id": pw.source_metric_id,
            "target_sheet": pw.target_sheet,
            "target_cell": pw.target_cell,
            "expected_value_type": pw.expected_value_type,
            "calculated_value": pw.value,
            "reference_value": ref_value,
            "reference_available": cell_read.available,
            "reference_unavailable_reason": "" if cell_read.available else cell_read.reason,
            "comparison_status": status,
            "delta": delta,
            "processing_scope_records": result.scope.processing_scope_records,
            "period_scope_records": result.scope.period_scope_records,
            "estado_excluded_records": result.scope.estado_excluded_records,
            "reference_file": reference.path.name if reference.available else "",
            "reference_match_kind": reference.match_kind,
        }
        row.update(_context_row(contexts.get(pw.source_metric_id)))
        detail_rows.append(row)

    base.detail_rows = detail_rows
    return base


# ---------------------------------------------------------------------------
# Resúmenes agregados
# ---------------------------------------------------------------------------


def _status_counts_row(rows: Iterable[dict]) -> dict:
    counts = Counter(r["comparison_status"] for r in rows)
    compared = sum(counts.values())
    exact = counts.get(STATUS_EXACT_MATCH, 0)
    zvb = counts.get(STATUS_ZERO_VS_BLANK, 0)
    mismatch = counts.get(STATUS_REAL_MISMATCH, 0)
    unavailable = counts.get(STATUS_REFERENCE_UNAVAILABLE, 0)
    rate = (exact + zvb) / compared if compared else None
    return {
        "compared": compared, "exact_match": exact, "zero_vs_blank": zvb,
        "real_mismatch": mismatch, "unavailable": unavailable, "equivalence_rate": rate,
    }


def month_summary_row(outcome: MonthOutcome) -> dict:
    row = {
        "group_type": "MONTH",
        "group_key": outcome.period_label,
        "month": outcome.month,
        "year": outcome.year,
        "status": outcome.status,
        "block_reason": outcome.block_reason,
        "reference_file": outcome.reference.path.name if outcome.reference.available else "",
        "reference_match_kind": outcome.reference.match_kind,
        "expected_writes": outcome.expected_writes,
        "pending_write_count": outcome.pending_write_count,
        "completeness_ok": outcome.completeness_ok,
        "processing_scope_records": outcome.scope_processing_records,
        "period_scope_records": outcome.scope_period_records,
        "estado_excluded_records": outcome.scope_estado_excluded,
    }
    row.update(_status_counts_row(outcome.detail_rows))
    return row


_GROUP_KEY_FNS: dict[str, Callable[[dict], object]] = {
    "TARGET_SHEET": lambda r: r["target_sheet"],
    "SOURCE_KIND": lambda r: r["source_kind"],
    "DEPENDS_ON_AGE_AF": lambda r: r["depends_on_age_AF"],
    "AXIS_TIPO_DE_CITA": lambda r: r["depends_on_tipo_de_cita"],
    "AXIS_PRESTACION": lambda r: r["depends_on_prestacion"],
    "AXIS_ESPECIALIDAD": lambda r: r["depends_on_especialidad"],
    "AXIS_SEXO": lambda r: r["depends_on_sexo"],
    "AXIS_ESTADO": lambda r: r["depends_on_estado"],
    "AXIS_AG_AL_RULES": lambda r: r["depends_on_ag_al_rules"],
}


def grouped_summary_rows(all_detail_rows: list[dict]) -> list[dict]:
    """Resumen agregado (sobre TODOS los meses ``COMPARED``) por hoja destino,
    tipo de métrica, dependencia de AF/edad y cada eje de investigación
    solicitado (TIPO_DE_CITA, PRESTACION, ESPECIALIDAD, SEXO, ESTADO, AG:AL)."""
    out: list[dict] = []
    if all_detail_rows:
        total = {"group_type": "TOTAL", "group_key": "ALL_MONTHS"}
        total.update(_status_counts_row(all_detail_rows))
        out.append(total)
    for group_type, key_fn in _GROUP_KEY_FNS.items():
        buckets: dict[str, list[dict]] = defaultdict(list)
        for r in all_detail_rows:
            buckets[str(key_fn(r))].append(r)
        for key in sorted(buckets):
            row = {"group_type": group_type, "group_key": key}
            row.update(_status_counts_row(buckets[key]))
            out.append(row)
    return out


_SUMMARY_FIELDS = (
    "group_type", "group_key", "month", "year", "status", "block_reason",
    "reference_file", "reference_match_kind", "expected_writes", "pending_write_count",
    "completeness_ok", "processing_scope_records", "period_scope_records",
    "estado_excluded_records", "compared", "exact_match", "zero_vs_blank",
    "real_mismatch", "unavailable", "equivalence_rate",
)


# ---------------------------------------------------------------------------
# Salidas
# ---------------------------------------------------------------------------


def _write_csv(path: Path, fieldnames: tuple[str, ...], rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames, restval="")
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def real_mismatch_full_stats(all_detail_rows: list[dict]) -> dict:
    """Estadísticas completas de los ``REAL_MISMATCH`` (todos los meses juntos):
    total, por hoja, dependencia de AF/edad (con *mismatch rate* AF vs no-AF
    sobre la población COMPARADA, no sólo sobre los mismatches), distribución
    del ``delta`` numérico y cuántos mismatches vienen de una referencia vacía
    (sin ``delta`` numérico, por diseño: no se fabrica un 0 para una celda
    histórica en blanco)."""
    compared_rows = [r for r in all_detail_rows if r["comparison_status"] != STATUS_REFERENCE_UNAVAILABLE]
    mismatches = [r for r in compared_rows if r["comparison_status"] == STATUS_REAL_MISMATCH]

    af_compared = [r for r in compared_rows if r["depends_on_age_AF"] is True]
    non_af_compared = [r for r in compared_rows if r["depends_on_age_AF"] is not True]
    af_mismatches = [r for r in mismatches if r["depends_on_age_AF"] is True]
    non_af_mismatches = [r for r in mismatches if r["depends_on_age_AF"] is not True]

    numeric_deltas = [r["delta"] for r in mismatches if isinstance(r["delta"], (int, float))]
    blank_reference_mismatches = [r for r in mismatches if r["reference_value"] is None]

    return {
        "total_real_mismatch": len(mismatches),
        "by_target_sheet": dict(sorted(Counter(r["target_sheet"] for r in mismatches).items())),
        "depends_on_age_af_count": len(af_mismatches),
        "does_not_depend_on_age_af_count": len(non_af_mismatches),
        "non_af_mismatch_instruction_ids": sorted(r["instruction_id"] for r in non_af_mismatches),
        "af_compared_population": len(af_compared),
        "non_af_compared_population": len(non_af_compared),
        "mismatch_rate_af": (len(af_mismatches) / len(af_compared)) if af_compared else None,
        "mismatch_rate_non_af": (
            (len(non_af_mismatches) / len(non_af_compared)) if non_af_compared else None
        ),
        "delta_numeric_count": len(numeric_deltas),
        "delta_blank_reference_count": len(blank_reference_mismatches),
        "delta_min": min(numeric_deltas) if numeric_deltas else None,
        "delta_max": max(numeric_deltas) if numeric_deltas else None,
        "delta_mean": (sum(numeric_deltas) / len(numeric_deltas)) if numeric_deltas else None,
        "reference_gt_calculated": sum(1 for d in numeric_deltas if d > 0),
        "reference_lt_calculated": sum(1 for d in numeric_deltas if d < 0),
        "reference_eq_calculated_flagged_as_mismatch": sum(1 for d in numeric_deltas if d == 0),
    }


def real_mismatch_axis_breakdown(all_detail_rows: list[dict]) -> dict:
    """De los ``REAL_MISMATCH``, qué fracción depende de cada eje de
    investigación (AF/edad, ESTADO, TIPO_DE_CITA, PRESTACION, ESPECIALIDAD,
    SEXO, reglas AG:AL). Sólo describe correlación con la fórmula legacy — no
    prueba causalidad y no cambia ninguna regla."""
    mismatches = [r for r in all_detail_rows if r["comparison_status"] == STATUS_REAL_MISMATCH]
    total = len(mismatches)
    axes = {
        "depends_on_age_AF": "AF / edad",
        "depends_on_estado": "ESTADO",
        "depends_on_tipo_de_cita": "TIPO_DE_CITA",
        "depends_on_prestacion": "PRESTACION",
        "depends_on_especialidad": "ESPECIALIDAD",
        "depends_on_sexo": "SEXO",
        "depends_on_ag_al_rules": "reglas AG:AL",
    }
    return {
        "total_real_mismatch": total,
        "by_axis": {
            label: sum(1 for r in mismatches if r[field] is True)
            for field, label in axes.items()
        },
    }


def build_report_md(
    outcomes: list[MonthOutcome], summary_rows: list[dict], all_detail_rows: list[dict]
) -> str:
    lines = [
        "# Validación histórica MEDINET — Abril–Julio 2026",
        "",
        (
            "Diagnóstico de sólo lectura. Compara el pipeline PRODUCTIVO actual "
            "(`build_production_pending_writes`) contra los REMASEP históricos reales, "
            "restringido a las 1122 celdas del `write_manifest` (automatización "
            "basada en `detalle_citas`). No se modificó ninguna lógica productiva, "
            "regla de clasificación, filtro ESTADO ni archivo histórico."
        ),
        "",
        "## Por mes",
        "",
        (
            "| Mes | Estado | Referencia | expected | compared | EXACT_MATCH | "
            "ZERO_VS_BLANK | REAL_MISMATCH | UNAVAILABLE | equivalence_rate |"
        ),
        "| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |",
    ]
    for outcome in outcomes:
        row = month_summary_row(outcome)
        rate = f"{row['equivalence_rate']:.4f}" if row["equivalence_rate"] is not None else "—"
        ref = row["reference_file"] or f"NO ENCONTRADO ({row['reference_match_kind']})"
        lines.append(
            f"| {outcome.period_label} | {outcome.status} | {ref} | "
            f"{row['expected_writes']} | {row['compared']} | {row['exact_match']} | "
            f"{row['zero_vs_blank']} | {row['real_mismatch']} | {row['unavailable']} | {rate} |"
        )
    lines.append("")

    blocked = [o for o in outcomes if o.status == "BLOCKED"]
    if blocked:
        lines.append("## Meses BLOCKED (invariantes de producción no cumplidos)")
        lines.append("")
        for o in blocked:
            lines.append(f"- **{o.period_label}**: {o.block_reason}")
        lines.append("")

    missing_ref = [o for o in outcomes if o.status == "COMPARED" and not o.reference.available]
    if missing_ref:
        lines.append("## Meses sin REMASEP histórico disponible")
        lines.append("")
        for o in missing_ref:
            lines.append(
                f"- **{o.period_label}**: se esperaba `{o.reference.requested_name}` en "
                f"`data/local/`. Resultado de la búsqueda: `{o.reference.match_kind}`."
                + (
                    f" Candidatos ambiguos ignorados: {', '.join(o.reference.ignored_candidates)}."
                    if o.reference.ignored_candidates
                    else ""
                )
            )
        lines.append("")

    lines.append("## Resumen agrupado")
    lines.append("")
    lines.append(
        "| group_type | group_key | compared | EXACT_MATCH | ZERO_VS_BLANK | "
        "REAL_MISMATCH | UNAVAILABLE | equivalence_rate |"
    )
    lines.append("| --- | --- | ---: | ---: | ---: | ---: | ---: | ---: |")
    for row in summary_rows:
        rate = f"{row['equivalence_rate']:.4f}" if row.get("equivalence_rate") is not None else "—"
        lines.append(
            f"| {row['group_type']} | {row['group_key']} | {row['compared']} | "
            f"{row['exact_match']} | {row['zero_vs_blank']} | {row['real_mismatch']} | "
            f"{row['unavailable']} | {rate} |"
        )
    lines.append("")

    full_stats = real_mismatch_full_stats(all_detail_rows)
    if full_stats["total_real_mismatch"]:
        fs = full_stats
        lines.append("## REAL_MISMATCH — estadísticas completas")
        lines.append("")
        lines.append(f"- total: **{fs['total_real_mismatch']}**")
        lines.append(
            "- por hoja: "
            + ", ".join(f"{sheet}={count}" for sheet, count in fs["by_target_sheet"].items())
        )
        lines.append(
            f"- dependen de AF/edad: {fs['depends_on_age_af_count']} · "
            f"NO dependen de AF/edad: {fs['does_not_depend_on_age_af_count']}"
        )
        if fs["non_af_mismatch_instruction_ids"]:
            lines.append(
                "  - instruction_id sin dependencia de AF: "
                + ", ".join(fs["non_af_mismatch_instruction_ids"])
            )
        af_rate = f"{fs['mismatch_rate_af']:.2%}" if fs["mismatch_rate_af"] is not None else "—"
        non_af_rate = (
            f"{fs['mismatch_rate_non_af']:.2%}" if fs["mismatch_rate_non_af"] is not None else "—"
        )
        lines.append(
            f"- mismatch rate sobre población comparada: AF-dependientes "
            f"{fs['depends_on_age_af_count']}/{fs['af_compared_population']} = {af_rate} · "
            f"no-AF {fs['does_not_depend_on_age_af_count']}/{fs['non_af_compared_population']} "
            f"= {non_af_rate}"
        )
        if fs["delta_numeric_count"]:
            delta_line = (
                f"- delta numérico (ambos valores numéricos): {fs['delta_numeric_count']} casos, "
                f"min={fs['delta_min']} max={fs['delta_max']} mean={fs['delta_mean']:.2f}"
            )
        else:
            delta_line = f"- delta numérico: {fs['delta_numeric_count']} casos"
        lines.append(delta_line)
        lines.append(
            f"  - reference > calculated: {fs['reference_gt_calculated']} · "
            f"reference < calculated: {fs['reference_lt_calculated']}"
        )
        lines.append(
            f"- REAL_MISMATCH con referencia histórica vacía/None (calculado no-cero, sin "
            f"delta numérico — no se fabrica un 0): {fs['delta_blank_reference_count']}"
        )
        lines.append("")

    axis_breakdown = real_mismatch_axis_breakdown(all_detail_rows)
    if axis_breakdown["total_real_mismatch"]:
        total = axis_breakdown["total_real_mismatch"]
        lines.append("## REAL_MISMATCH — correlación con ejes de investigación")
        lines.append("")
        lines.append(
            f"{total} REAL_MISMATCH en total (todos los meses). Correlación (NO causalidad "
            "confirmada) con cada eje — cuántos de esos mismatches dependen, de forma "
            "transitiva, de cada campo:"
        )
        lines.append("")
        lines.append("| eje | mismatches que dependen de él | % del total |")
        lines.append("| --- | ---: | ---: |")
        for label, count in axis_breakdown["by_axis"].items():
            pct = f"{count / total:.1%}"
            lines.append(f"| {label} | {count} | {pct} |")
        lines.append("")
        lines.append(
            "No se cambió ninguna regla a partir de esta correlación: es evidencia "
            "diagnóstica para investigación posterior, no una conclusión aplicada."
        )
        lines.append("")

    lines.append("## Limitaciones del método")
    lines.append("")
    lines.append(
        "- No existen exports Medinet históricos por mes: los cuatro meses se "
        "recalculan sobre el mismo `detalle_citas` actual, seleccionando cada "
        "período con la lógica productiva. Si el detalle histórico real difería "
        "(altas/bajas de registros posteriores, correcciones de datos en Medinet), "
        "esta comparación no puede detectarlo — sólo compara la lógica del motor, "
        "no la fidelidad histórica del dato fuente."
    )
    lines.append(
        "- Una celda histórica vacía/None LEÍDA CON ÉXITO es un valor de "
        "referencia válido, no `REFERENCE_UNAVAILABLE`: calculado 0 → "
        "`ZERO_VS_BLANK_EQUIVALENT`; calculado no-cero → `REAL_MISMATCH` (ver "
        "`reference_available`/`reference_unavailable_reason` en el CSV de "
        "detalle). `REFERENCE_UNAVAILABLE` se reserva para cuando la lectura en "
        "sí falla: archivo ausente/ambiguo, hoja inexistente o celda ilegible."
    )
    lines.append(
        "- El campo `delta` sólo se calcula cuando ambos valores son numéricos "
        "(`reference_value - calculated_value`); un `REAL_MISMATCH` originado por "
        "una referencia vacía frente a un calculado no-cero queda con `delta` "
        "vacío a propósito — no se fabrica un 0 para la celda en blanco."
    )
    lines.append("")
    return "\n".join(lines)


def write_outputs(output_dir: Path, outcomes: list[MonthOutcome]) -> dict:
    output_dir.mkdir(parents=True, exist_ok=True)
    all_detail_rows = [row for o in outcomes for row in o.detail_rows]
    summary_rows = [month_summary_row(o) for o in outcomes] + grouped_summary_rows(all_detail_rows)

    _write_csv(output_dir / "historical_regression_detail.csv", _DETAIL_FIELDS, all_detail_rows)
    _write_csv(output_dir / "historical_regression_summary.csv", _SUMMARY_FIELDS, summary_rows)
    (output_dir / "historical_regression_report.md").write_text(
        build_report_md(outcomes, summary_rows, all_detail_rows), encoding="utf-8"
    )

    axis_breakdown = real_mismatch_axis_breakdown(all_detail_rows)
    full_stats = real_mismatch_full_stats(all_detail_rows)

    summary_json = {
        "generated_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "scope": "solo las 1122 celdas del write_manifest (automatización detalle_citas)",
        "months": [month_summary_row(o) for o in outcomes],
        "total_compared": sum(len(o.detail_rows) for o in outcomes),
        "total_real_mismatches": axis_breakdown["total_real_mismatch"],
        "real_mismatch_axis_breakdown": axis_breakdown["by_axis"],
        "real_mismatch_full_stats": full_stats,
    }
    (output_dir / "summary.json").write_text(
        json.dumps(summary_json, indent=2, ensure_ascii=False, default=str) + "\n",
        encoding="utf-8",
    )
    return {
        "outcomes": outcomes,
        "detail_rows": all_detail_rows,
        "summary_rows": summary_rows,
        "summary_json": summary_json,
        "files": [
            output_dir / "historical_regression_detail.csv",
            output_dir / "historical_regression_summary.csv",
            output_dir / "historical_regression_report.md",
            output_dir / "summary.json",
        ],
    }


# ---------------------------------------------------------------------------
# Orquestación + CLI
# ---------------------------------------------------------------------------


def historical_regression(
    export_path: str | Path,
    data_dir: str | Path,
    months: Iterable[int] = MONTHS_IN_SCOPE,
    year: int = YEAR_IN_SCOPE,
    *,
    bundle: RuntimeBundle | None = None,
) -> list[MonthOutcome]:
    export_path = Path(export_path)
    data_dir = Path(data_dir)
    bundle = bundle or load_runtime_bundle()
    contexts = build_metric_contexts(bundle)
    return [
        run_month(export_path, data_dir, month, year, bundle, contexts)
        for month in months
    ]


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="compare_historical_remasep.py",
        description=(
            "Validación histórica MEDINET contra REMASEP reales (Abril-Julio 2026). "
            "Diagnóstico, sólo lectura, no modifica lógica productiva ni archivos "
            "históricos."
        ),
    )
    parser.add_argument("--export", default=DEFAULT_EXPORT)
    parser.add_argument("--data-dir", default=DEFAULT_DATA_DIR)
    parser.add_argument("--months", type=int, nargs="+", default=list(MONTHS_IN_SCOPE))
    parser.add_argument("--year", type=int, default=YEAR_IN_SCOPE)
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        outcomes = historical_regression(
            args.export, args.data_dir, args.months, args.year
        )
    except RemasepError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    outcome = write_outputs(Path(args.output), outcomes)
    for o in outcomes:
        row = month_summary_row(o)
        print(f"{o.period_label}: status={o.status}")
        if o.status == "BLOCKED":
            print(f"  BLOCKED: {o.block_reason}")
            continue
        print(f"  referencia: {row['reference_file'] or '(no encontrada: ' + row['reference_match_kind'] + ')'}")
        print(
            f"  compared={row['compared']} EXACT_MATCH={row['exact_match']} "
            f"ZERO_VS_BLANK={row['zero_vs_blank']} REAL_MISMATCH={row['real_mismatch']} "
            f"UNAVAILABLE={row['unavailable']} equivalence_rate={row['equivalence_rate']}"
        )
    print(f"Artefactos en {args.output}:")
    for f in outcome["files"]:
        print(f"  - {f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
