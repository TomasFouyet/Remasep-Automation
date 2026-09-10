"""Sprint 3.7A — Medinet Metric Value Producer & Reference Value Equivalence.

Responde **qué valor** alimenta cada ``WriteInstruction`` del manifiesto de
escritura (Sprint 3.6) y compara ese valor, celda a celda, contra la referencia
legacy y la plantilla final oficial.

Reutiliza el motor legacy existente (``close_legacy_aggregations`` para el grafo
y las fórmulas por celda; :mod:`remasep.services.metric_value_producer` para la
evaluación, la unión con el manifiesto y la equivalencia). **No escribe Excel**,
no usa COM, no modifica la plantilla. Verifica el SHA256 de todas las fuentes
antes y después. Sin PII: los artefactos contienen coordenadas y **un conteo
agregado** por métrica, nunca filas.

Modos:

- ``PRODUCTION_PERIOD_SCOPE`` (siempre): corrida real sobre el export directo
  ``detalle_citas`` del período, ``processing_scope_records`` (válidos ∩ mes),
  **sin** filtro por ESTADO.
- ``LEGACY_EQUIVALENCE_DIAGNOSTIC`` (``--diagnostic``): fuera del flujo
  productivo; aplica la hipótesis ESTADO observada sólo para comparar contra el
  snapshot legacy de julio. No convierte la hipótesis en regla.

Uso:

    python scripts/build_metric_values.py \\
        "data/local/detalle_citas - 2026-09-07T123630.940.xlsx" \\
        "data/local/GENERACION DATOS REMASEP.xlsx" \\
        "data/local/REMASEP_V1.4 Julio 2026.xlsm" \\
        --period 2026-07 --diagnostic
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sys
from collections import Counter, defaultdict
from datetime import UTC, datetime
from pathlib import Path

import openpyxl
import yaml
from close_legacy_aggregations import KIND_BASE, KIND_DERIVED, close_legacy_aggregations
from compare_legacy_aggregations import _make_resolver

from remasep.core.errors import RemasepError
from remasep.services.common import Period
from remasep.services.legacy_rules import load_legacy_rules
from remasep.services.medinet_analysis import MedinetAnalysisService, processing_scope_frame
from remasep.services.medinet_input_reconciliation import LEGACY_KEPT_STATES
from remasep.services.metric_value_producer import (
    CMP_MATCH,
    CMP_MISMATCH,
    CMP_SOURCE_UNAVAILABLE,
    CMP_TARGET_UNAVAILABLE,
    CMP_ZERO_VS_BLANK,
    MODE_DIAGNOSTIC,
    MODE_PRODUCTION,
    PRODUCER_VERSION,
    ZERO_POLICY_PRESERVE_BLANK,
    ZERO_POLICY_UNRESOLVED,
    ZERO_POLICY_WRITE,
    ZERO_SOURCE_ZERO_TARGET_BLANK,
    ZERO_SOURCE_ZERO_TARGET_OTHER,
    ZERO_SOURCE_ZERO_TARGET_ZERO,
    LegacyFormulaSpec,
    MetricValue,
    MetricValueProducer,
    build_rows,
    check_write_completeness,
    compare_reference_values,
    join_metric_values_with_manifest,
)

DEFAULT_OUTPUT = "artifacts/metric_value_producer"
DEFAULT_EXPORT = "data/local/detalle_citas - 2026-09-07T123630.940.xlsx"
DEFAULT_GENERATOR = "data/local/GENERACION DATOS REMASEP.xlsx"
DEFAULT_TEMPLATE = "data/local/REMASEP_V1.4 Julio 2026.xlsm"
DEFAULT_BASE_TEMPLATE = "data/local/2026-7 REMASEP_V1.4.xlsm"
DEFAULT_MANIFEST = "artifacts/writable_target_mapping/write_manifest_ready.csv"
ZERO_POLICY_CONFIG = "config/metric_value_producer_2026/zero_write_policy.yaml"

ESTADO_FILTER_STATUS = "PENDING_FUNCTIONAL_CONFIRMATION"

_FORM_BY_SHEET = {
    "REMASEP_OD": "REMASEP_OD",
    "REMASEP 01": "REMASEP_01",
    "B2 ANEXO": "B2_ANEXO",
}


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _parse_period(text: str) -> Period:
    try:
        year_s, month_s = text.split("-", 1)
        return Period(month=int(month_s), year=int(year_s))
    except (ValueError, TypeError) as exc:
        raise RemasepError(f"período inválido {text!r}, se esperaba AAAA-MM") from exc


def normalize_form(sheet: str) -> str:
    return _FORM_BY_SHEET.get(sheet, sheet.replace(" ", "_"))


def _section_of(signature: str) -> str:
    parts = [p.strip() for p in signature.split("::")]
    return parts[1] if len(parts) > 1 else ""


class ManifestRow:
    """Fila del manifiesto ready con la interfaz que consumen el join y la
    verificación de completitud (duck-typing sobre ``WriteInstruction``)."""

    __slots__ = (
        "expected_value_type",
        "instruction_id",
        "source_metric_id",
        "source_semantic_signature",
        "target_cell",
        "target_sheet",
    )

    def __init__(self, row: dict[str, str]) -> None:
        self.instruction_id = row["instruction_id"]
        self.source_metric_id = row["source_metric_id"]
        self.source_semantic_signature = row.get("source_semantic_signature", "")
        self.target_sheet = row["target_sheet"]
        self.target_cell = row["target_cell"]
        self.expected_value_type = row.get("expected_value_type", "INTEGER_COUNT")

    @property
    def form(self) -> str:
        return normalize_form(self.target_sheet)


def load_manifest(path: Path) -> list[ManifestRow]:
    with path.open(encoding="utf-8", newline="") as handle:
        rows = [ManifestRow(row) for row in csv.DictReader(handle)]
    if not rows:
        raise RemasepError(f"manifiesto vacío: {path}")
    return rows


# ---------------------------------------------------------------------------
# Índice de fórmulas legacy + valores de referencia legacy
# ---------------------------------------------------------------------------


class LegacyReference:
    """Fórmulas por celda agregada y valor legacy de referencia (reevaluado)."""

    def __init__(self, closure) -> None:
        self.detail_sheet = closure.detail_sheet
        self.detail_columns_map = dict(closure.detail_columns_map)
        self.active_records = closure.active_records
        self.resolver_by_sheet = {
            sheet: _make_resolver(cached) for sheet, cached in closure.cached.items()
        }
        self.formula_index: dict[str, LegacyFormulaSpec] = {}
        self.reference_value: dict[str, object] = {}
        self.cached_value: dict[str, object] = {}
        self.cache_status: dict[str, str] = {}
        self.kind: dict[str, str] = {}
        for (sheet, cell), node in closure.nodes.items():
            if node.kind not in (KIND_BASE, KIND_DERIVED) or not node.medinet:
                continue
            mid = node.metric_id
            self.formula_index[mid] = LegacyFormulaSpec(
                source_metric_id=mid,
                sheet=sheet,
                cell=cell,
                kind=node.kind,
                formula=node.formula,
                value_ref_coords=tuple(node.value_ref_coords),
            )
            self.reference_value[mid] = node.python_value
            self.cached_value[mid] = node.cached_value
            self.cache_status[mid] = node.cache_status or ""
            self.kind[mid] = node.kind


def build_legacy_reference(generator_path: str | Path) -> LegacyReference:
    return LegacyReference(close_legacy_aggregations(generator_path))


# ---------------------------------------------------------------------------
# Lectura de celdas destino en la plantilla (sólo lectura, sin recalcular)
# ---------------------------------------------------------------------------


def read_target_values(path: Path, sheet_cells: set[tuple[str, str]]) -> dict[tuple[str, str], object]:
    wb = openpyxl.load_workbook(path, data_only=True, read_only=True)
    try:
        present = set(wb.sheetnames)
        out: dict[tuple[str, str], object] = {}
        for sheet, cell in sheet_cells:
            if sheet not in present:
                out[(sheet, cell)] = None
                continue
            out[(sheet, cell)] = wb[sheet][cell].value
    finally:
        wb.close()
    return out


# ---------------------------------------------------------------------------
# Corridas del productor
# ---------------------------------------------------------------------------


def _diagnostic_frame(frame, kept_states: tuple[str, ...]):
    kept = {s.strip().casefold() for s in kept_states}
    estado = frame["ESTADO"].astype("string").str.strip().str.casefold()
    return frame.loc[estado.isin(kept)].reset_index(drop=True)


def run_producer(
    frame,
    reference: LegacyReference,
    manifest: list[ManifestRow],
    period: Period,
    *,
    mode: str,
):
    ruleset = load_legacy_rules()
    rows = build_rows(frame.to_dict("records"), reference.detail_columns_map, ruleset)
    producer = MetricValueProducer(
        formula_index=reference.formula_index,
        detail_sheet=reference.detail_sheet,
        resolver_by_sheet=reference.resolver_by_sheet,
        rows=rows,
        period=period,
        scope_records=len(frame),
        mode=mode,
    )
    expected = {row.source_metric_id: row.expected_value_type for row in manifest}
    ids = [row.source_metric_id for row in manifest]
    return producer.produce_run(ids, expected_types=expected)


# ---------------------------------------------------------------------------
# Equivalencia de valores de referencia + semántica del cero + mismatches
# ---------------------------------------------------------------------------


def build_reference_equivalence(
    manifest: list[ManifestRow],
    reference: LegacyReference,
    final_values: dict[tuple[str, str], object],
) -> list[dict]:
    rows: list[dict] = []
    for row in manifest:
        mid = row.source_metric_id
        # `source_reference_value` = valor legacy REEVALUADO (motor de agregación
        # sobre el detalle legacy). `source_reference_value_cached` = lo que el
        # workbook legacy dejó en caché (puede estar obsoleto: AF/edad, ver
        # docs/TECHNICAL_OVERVIEW.md §8).
        source_value = reference.reference_value.get(mid)
        cached_value = reference.cached_value.get(mid)
        target_value = final_values.get((row.target_sheet, row.target_cell))
        target_present = (row.target_sheet, row.target_cell) in final_values and target_value is not None
        status, zero_sem = compare_reference_values(
            source_value, target_value, target_present=target_present
        )
        reasons: list[str] = []
        if cached_value != source_value:
            reasons.append(f"SOURCE_CACHE_{reference.cache_status.get(mid) or 'DIFFERENCE'}")
        if (
            status == CMP_MISMATCH
            and target_present
            and cached_value == target_value
            and source_value != target_value
        ):
            reasons.append("LEGACY_CACHE_MATCHES_TARGET")
        rows.append(
            {
                "instruction_id": row.instruction_id,
                "source_metric_id": mid,
                "form": row.form,
                "target_sheet": row.target_sheet,
                "target_cell": row.target_cell,
                "source_reference_value": source_value,
                "source_reference_value_cached": cached_value,
                "target_reference_value": target_value,
                "comparison_status": status,
                "zero_semantics": zero_sem,
                "reason": "|".join(reasons),
            }
        )
    return rows


def build_zero_semantics(
    equivalence_rows: list[dict],
    reference: LegacyReference,
    manifest_by_id: dict[str, ManifestRow],
    final_values: dict[tuple[str, str], object],
    base_values: dict[tuple[str, str], object],
) -> list[dict]:
    rows: list[dict] = []
    for row in equivalence_rows:
        source_value = row["source_reference_value"]
        if not (isinstance(source_value, (int, float)) and source_value == 0):
            continue
        manifest_row = manifest_by_id[row["source_metric_id"]]
        key = (row["target_sheet"], row["target_cell"])
        final_value = final_values.get(key)
        base_value = base_values.get(key)
        rows.append(
            {
                "instruction_id": row["instruction_id"],
                "source_metric_id": row["source_metric_id"],
                "form": row["form"],
                "section": _section_of(manifest_row.source_semantic_signature),
                "target_sheet": row["target_sheet"],
                "target_cell": row["target_cell"],
                "final_target_value": final_value,
                "base_template_value": base_value,
                "zero_semantics": row["zero_semantics"],
            }
        )
    return rows


def derive_zero_policy(zero_rows: list[dict], *, min_cases: int, min_ratio: float) -> dict:
    """Deriva la política de escritura del cero a partir de la evidencia.

    ``UNRESOLVED`` si la evidencia es insuficiente o inconsistente. No se fuerza.
    """
    total = len(zero_rows)
    counts = Counter(r["zero_semantics"] for r in zero_rows)
    by_form: dict[str, Counter] = defaultdict(Counter)
    for r in zero_rows:
        by_form[r["form"]][r["zero_semantics"]] += 1
    target_zero = counts.get(ZERO_SOURCE_ZERO_TARGET_ZERO, 0)
    target_blank = counts.get(ZERO_SOURCE_ZERO_TARGET_BLANK, 0)
    target_other = counts.get(ZERO_SOURCE_ZERO_TARGET_OTHER, 0)
    decided = target_zero + target_blank
    resolution = ZERO_POLICY_UNRESOLVED
    if total >= min_cases and decided:
        if target_blank == 0 and target_zero / decided >= min_ratio:
            resolution = ZERO_POLICY_WRITE
        elif target_zero == 0 and target_blank / decided >= min_ratio:
            resolution = ZERO_POLICY_PRESERVE_BLANK
    return {
        "resolution": resolution,
        "observed": {
            "zero_source_cases": total,
            "source_zero_target_zero": target_zero,
            "source_zero_target_blank": target_blank,
            "source_zero_target_other": target_other,
            "by_form": {
                form: dict(sorted(sem.items())) for form, sem in sorted(by_form.items())
            },
        },
    }


def compare_zero_policy_config(config_doc: dict, derived: dict) -> dict:
    """Contrasta la política declarada en config con la derivada de la evidencia.

    **No** sobrescribe el archivo de config (hand-authored, revisado). Sólo
    informa si coinciden.
    """
    declared = str((config_doc or {}).get("resolution", ZERO_POLICY_UNRESOLVED))
    return {
        "config_resolution": declared,
        "evidence_resolution": derived["resolution"],
        "config_matches_evidence": declared == derived["resolution"],
    }


def value_equivalence_summary(equivalence_rows: list[dict]) -> list[dict]:
    buckets = ("MATCH", "ZERO_VS_BLANK", "MISMATCH", "UNAVAILABLE")
    per_form: dict[str, Counter] = defaultdict(Counter)
    for row in equivalence_rows:
        status = row["comparison_status"]
        if status == CMP_MATCH:
            bucket = "MATCH"
        elif status == CMP_ZERO_VS_BLANK:
            bucket = "ZERO_VS_BLANK"
        elif status == CMP_MISMATCH:
            bucket = "MISMATCH"
        else:
            bucket = "UNAVAILABLE"
        per_form[row["form"]][bucket] += 1
        per_form["ALL"][bucket] += 1
    return [
        {"form": form, **{b: per_form[form].get(b, 0) for b in buckets},
         "total": sum(per_form[form].values())}
        for form in sorted(per_form)
    ]


def _investigation_axis(row: dict, reference: LegacyReference) -> str:
    mid = row["source_metric_id"]
    if "LEGACY_CACHE_MATCHES_TARGET" in row.get("reason", ""):
        return "LEGACY_CACHE_MATCHES_TARGET"
    if row["zero_semantics"] == ZERO_SOURCE_ZERO_TARGET_OTHER:
        return "ZERO_SEMANTICS"
    if reference.cache_status.get(mid) == "CACHE_DIFFERENCE":
        return "LEGACY_CACHE_SUSPECT"
    if row["comparison_status"] == CMP_TARGET_UNAVAILABLE:
        return "TARGET_UNAVAILABLE"
    if row["comparison_status"] == CMP_SOURCE_UNAVAILABLE:
        return "SOURCE_UNAVAILABLE"
    return "SOURCE_POPULATION_OR_TARGET_ADJUSTMENT"


def build_mismatches(equivalence_rows: list[dict], reference: LegacyReference) -> list[dict]:
    rows: list[dict] = []
    for row in equivalence_rows:
        if row["comparison_status"] not in (CMP_MISMATCH, CMP_TARGET_UNAVAILABLE, CMP_SOURCE_UNAVAILABLE):
            continue
        mid = row["source_metric_id"]
        source_value = row["source_reference_value"]
        target_value = row["target_reference_value"]
        delta = None
        if isinstance(source_value, (int, float)) and isinstance(target_value, (int, float)):
            delta = target_value - source_value
        rows.append(
            {
                "instruction_id": row["instruction_id"],
                "source_metric_id": mid,
                "form": row["form"],
                "source_kind": reference.kind.get(mid, ""),
                "target_sheet": row["target_sheet"],
                "target_cell": row["target_cell"],
                "source_reference_value": source_value,
                "target_reference_value": target_value,
                "delta": delta,
                "comparison_status": row["comparison_status"],
                "zero_semantics": row["zero_semantics"],
                "investigation_axis": _investigation_axis(row, reference),
            }
        )
    return rows


def cluster_mismatches(mismatch_rows: list[dict]) -> list[dict]:
    clusters: dict[tuple, list[dict]] = defaultdict(list)
    for row in mismatch_rows:
        key = (
            row["form"],
            row["source_kind"],
            row["comparison_status"],
            row["zero_semantics"],
            row["investigation_axis"],
        )
        clusters[key].append(row)
    out: list[dict] = []
    for key, members in sorted(clusters.items(), key=lambda kv: (-len(kv[1]), kv[0])):
        deltas = [m["delta"] for m in members if isinstance(m["delta"], (int, float))]
        out.append(
            {
                "form": key[0],
                "source_kind": key[1],
                "comparison_status": key[2],
                "zero_semantics": key[3],
                "investigation_axis": key[4],
                "count": len(members),
                "delta_min": min(deltas) if deltas else None,
                "delta_max": max(deltas) if deltas else None,
                "example_target": f"{members[0]['target_sheet']}!{members[0]['target_cell']}",
            }
        )
    return out


# ---------------------------------------------------------------------------
# Salidas
# ---------------------------------------------------------------------------


def _write_csv(path: Path, fieldnames: list[str], rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=fieldnames)
        writer.writeheader()
        for row in rows:
            writer.writerow(row)


def _metric_value_rows(run) -> list[dict]:
    return [
        {
            "source_metric_id": mv.source_metric_id,
            "value": mv.value,
            "period": mv.period,
            "producer_version": mv.producer_version,
        }
        for mv in sorted(run.metric_values, key=lambda mv: mv.source_metric_id)
    ]


def _pending_rows(pending) -> list[dict]:
    return [
        {
            "instruction_id": pw.instruction_id,
            "source_metric_id": pw.source_metric_id,
            "target_sheet": pw.target_sheet,
            "target_cell": pw.target_cell,
            "value": pw.value,
            "expected_value_type": pw.expected_value_type,
            "period": pw.period,
        }
        for pw in sorted(pending, key=lambda pw: pw.instruction_id)
    ]


def write_outputs(output_dir: Path, payload: dict) -> list[str]:
    output_dir.mkdir(parents=True, exist_ok=True)
    files: list[str] = []

    def emit(name: str, fieldnames: list[str], rows: list[dict]) -> None:
        _write_csv(output_dir / name, fieldnames, rows)
        files.append(name)

    emit(
        "production_metric_values.csv",
        ["source_metric_id", "value", "period", "producer_version"],
        _metric_value_rows(payload["production_run"]),
    )
    diag_rows = (
        _metric_value_rows(payload["diagnostic_run"])
        if payload["diagnostic_run"] is not None
        else []
    )
    emit(
        "diagnostic_legacy_metric_values.csv",
        ["source_metric_id", "value", "period", "producer_version"],
        diag_rows,
    )
    emit(
        "reference_value_equivalence.csv",
        [
            "instruction_id", "source_metric_id", "form", "target_sheet", "target_cell",
            "source_reference_value", "source_reference_value_cached",
            "target_reference_value", "comparison_status", "zero_semantics", "reason",
        ],
        payload["equivalence_rows"],
    )
    emit(
        "value_equivalence_summary.csv",
        ["form", "MATCH", "ZERO_VS_BLANK", "MISMATCH", "UNAVAILABLE", "total"],
        payload["equivalence_summary"],
    )
    emit(
        "zero_semantics_analysis.csv",
        [
            "instruction_id", "source_metric_id", "form", "section", "target_sheet",
            "target_cell", "final_target_value", "base_template_value", "zero_semantics",
        ],
        payload["zero_rows"],
    )
    emit(
        "value_mismatches.csv",
        [
            "instruction_id", "source_metric_id", "form", "source_kind", "target_sheet",
            "target_cell", "source_reference_value", "target_reference_value", "delta",
            "comparison_status", "zero_semantics", "investigation_axis",
        ],
        payload["mismatch_rows"],
    )
    emit(
        "value_mismatch_clusters.csv",
        [
            "form", "source_kind", "comparison_status", "zero_semantics",
            "investigation_axis", "count", "delta_min", "delta_max", "example_target",
        ],
        payload["mismatch_clusters"],
    )
    emit(
        "pending_writes_reference.csv",
        [
            "instruction_id", "source_metric_id", "target_sheet", "target_cell",
            "value", "expected_value_type", "period",
        ],
        payload["pending_reference"],
    )
    emit(
        "pending_writes_production.csv",
        [
            "instruction_id", "source_metric_id", "target_sheet", "target_cell",
            "value", "expected_value_type", "period",
        ],
        payload["pending_production"],
    )

    (output_dir / "summary.json").write_text(
        json.dumps(payload["summary"], indent=2, ensure_ascii=False, default=str) + "\n",
        encoding="utf-8",
    )
    files.append("summary.json")
    (output_dir / "README.md").write_text(payload["readme"], encoding="utf-8")
    files.append("README.md")
    return files


def _readme(summary: dict) -> str:
    return f"""# Metric Value Producer — Sprint 3.7A

Productor real de `MetricValue` para MEDINET y equivalencia de valores contra la
referencia legacy y la plantilla final oficial. **Solo lectura**: no se escribe
Excel, no se usa COM, no se modifica ninguna fuente. Sin PII (conteos agregados).

## Fuentes (SHA256 verificado sin cambios: {summary['source_hash_verified_unchanged']})

- export directo (producción): `{summary['export_name']}`
- referencia legacy: `{summary['generator_name']}`
- plantilla final oficial: `{summary['template_name']}`
- plantilla base/incompleta: `{summary['base_template_name']}`

## Contrato de período

`processing_scope_records` = registros estructuralmente válidos ∩ mes/año
seleccionado. **No** se aplica el filtro por ESTADO (pendiente de confirmación
funcional: `estado_filter_status = {summary['estado_filter_status']}`).

## Modo producción (`PRODUCTION_PERIOD_SCOPE`)

- período: {summary['production_period']}
- registros de alcance: {summary['production_scope_records']}
- `MetricValue` producidos: {summary['production_metric_values']}
- valores WRITE_READY disponibles: {summary['production_write_ready_values']} /
  {summary['write_manifest_instruction_count']}
- WRITE_READY sin valor: {summary['production_missing_values']}
- conflictos de tipo de valor: {summary['production_value_type_conflicts']}
- completitud (missing/duplicate/orphan):
  {summary['production_completeness']['missing']} /
  {summary['production_completeness']['duplicate']} /
  {summary['production_completeness']['orphan']}

No se compara el modo producción contra la referencia legacy final como si
debieran coincidir: el filtro por ESTADO no está confirmado.

## Modo diagnóstico (`LEGACY_EQUIVALENCE_DIAGNOSTIC`)

{summary['diagnostic_note']}

## Equivalencia de valores de referencia (1122 WriteInstructions)

`source_reference_value` = valor legacy reevaluado sobre el detalle legacy
(`{summary['generator_name']}`, {summary['legacy_active_records']} filas activas).
`target_reference_value` = celda final en `{summary['template_name']}`.

- MATCH: {summary['reference_value_matches']}
- ZERO_VS_BLANK_EQUIVALENT: {summary['zero_vs_blank']}
- MISMATCH: {summary['mismatches']}
- UNAVAILABLE (target/source): {summary['unavailable']}

Por forma: ver `value_equivalence_summary.csv`.

Hallazgo: si en lugar del valor reevaluado se compara el **valor cacheado** por
el workbook legacy (`source_reference_value_cached`), coinciden con la celda
final {summary['reference_value_matches_vs_legacy_cache']}/1122. Es decir, la
plantilla final oficial se llenó con los valores que el Excel legacy tenía en
caché — incluida la caché obsoleta de la columna AF/edad
(`docs/TECHNICAL_OVERVIEW.md §8`). Los {summary['mismatches']} MISMATCH son ese
artefacto de caché, no un error del productor. No se autocorrige.

## Semántica del cero

Casos con `source_reference_value == 0`: {summary['zero_source_cases']}.
`SOURCE_ZERO_TARGET_ZERO` = {summary['zero_semantics']['source_zero_target_zero']},
`SOURCE_ZERO_TARGET_BLANK` = {summary['zero_semantics']['source_zero_target_blank']},
`SOURCE_ZERO_TARGET_OTHER` = {summary['zero_semantics']['source_zero_target_other']}.

`zero_write_policy` derivada: **{summary['zero_write_policy']}** (ver
`config/metric_value_producer_2026/zero_write_policy.yaml`).

## Mismatches

{summary['mismatches']} mismatches. No se autocorrigen. Agrupados en
`value_mismatch_clusters.csv` por forma / kind / estado / semántica del cero /
eje de investigación.

## Archivos

`production_metric_values.csv`, `diagnostic_legacy_metric_values.csv`,
`reference_value_equivalence.csv`, `value_equivalence_summary.csv`,
`zero_semantics_analysis.csv`, `value_mismatches.csv`,
`value_mismatch_clusters.csv`, `pending_writes_reference.csv`,
`pending_writes_production.csv`, `summary.json`.
"""


# ---------------------------------------------------------------------------
# Orquestación
# ---------------------------------------------------------------------------


def build_analysis(
    export_path: str,
    generator_path: str,
    template_path: str,
    base_template_path: str,
    manifest_path: str,
    period: Period,
    *,
    run_diagnostic: bool,
    zero_policy_config: str = ZERO_POLICY_CONFIG,
) -> dict:
    export = Path(export_path)
    generator = Path(generator_path)
    template = Path(template_path)
    base_template = Path(base_template_path)
    for candidate in (export, generator, template):
        if not candidate.is_file():
            raise RemasepError(f"no existe el archivo: {candidate}")

    sha_before = {
        p.name: _sha256(p) for p in (export, generator, template) if p.is_file()
    }
    if base_template.is_file():
        sha_before[base_template.name] = _sha256(base_template)

    manifest = load_manifest(Path(manifest_path))
    manifest_by_id = {row.source_metric_id: row for row in manifest}
    reference = build_legacy_reference(generator)

    unresolved_ids = [
        row.source_metric_id
        for row in manifest
        if row.source_metric_id not in reference.formula_index
    ]

    frame = processing_scope_frame(export, period)
    analysis_scope = MedinetAnalysisService().analyze(export, period).processing_scope_records
    if len(frame) != analysis_scope:
        raise RemasepError(
            f"desajuste de alcance: processing_scope_frame={len(frame)} != "
            f"MedinetAnalysisResult.processing_scope_records={analysis_scope}"
        )

    production_run = run_producer(frame, reference, manifest, period, mode=MODE_PRODUCTION)
    production_completeness = check_write_completeness(production_run.metric_values, manifest)
    pending_production = join_metric_values_with_manifest(production_run.metric_values, manifest)

    diagnostic_run = None
    diagnostic_equivalence = None
    if run_diagnostic:
        diag_frame = _diagnostic_frame(frame, LEGACY_KEPT_STATES)
        diagnostic_run = run_producer(
            diag_frame, reference, manifest, period, mode=MODE_DIAGNOSTIC
        )
        matches = mismatches = 0
        diff_examples: list[dict] = []
        for mv in diagnostic_run.metric_values:
            legacy = reference.reference_value.get(mv.source_metric_id)
            if legacy == mv.value:
                matches += 1
            else:
                mismatches += 1
                if len(diff_examples) < 25:
                    diff_examples.append(
                        {
                            "source_metric_id": mv.source_metric_id,
                            "legacy_value": legacy,
                            "diagnostic_value": mv.value,
                        }
                    )
        diagnostic_equivalence = {
            "diagnostic_scope_records": len(diag_frame),
            "expected_legacy_active_records": reference.active_records,
            "reproduces_scope_exactly": len(diag_frame) == reference.active_records,
            "metric_values": len(diagnostic_run.metric_values),
            "equal_to_legacy": matches,
            "different_from_legacy": mismatches,
            "status": "EXACT_EQUIVALENCE" if mismatches == 0 else "DIFFERENCES",
            "examples": diff_examples,
        }

    sheet_cells = {(row.target_sheet, row.target_cell) for row in manifest}
    final_values = read_target_values(template, sheet_cells)
    base_values = (
        read_target_values(base_template, sheet_cells) if base_template.is_file() else {}
    )

    equivalence_rows = build_reference_equivalence(manifest, reference, final_values)
    equivalence_summary = value_equivalence_summary(equivalence_rows)
    zero_rows = build_zero_semantics(
        equivalence_rows, reference, manifest_by_id, final_values, base_values
    )
    policy_doc = yaml.safe_load(Path(zero_policy_config).read_text(encoding="utf-8")) or {}
    ev_cfg = policy_doc.get("evidence") or {}
    derived_policy = derive_zero_policy(
        zero_rows,
        min_cases=int(ev_cfg.get("min_zero_cases", 30)),
        min_ratio=float(ev_cfg.get("min_consistency_ratio", 0.98)),
    )
    zero_policy_check = compare_zero_policy_config(policy_doc, derived_policy)

    mismatch_rows = build_mismatches(equivalence_rows, reference)
    mismatch_clusters = cluster_mismatches(mismatch_rows)

    reference_metric_values = {
        mid: reference.reference_value.get(mid)
        for mid in manifest_by_id
        if isinstance(reference.reference_value.get(mid), (int, float))
    }
    pending_reference = join_metric_values_with_manifest(
        [
            MetricValue(
                source_metric_id=mid,
                value=value,
                period=period.label,
                producer_version=f"legacy_reference::{reference.detail_sheet}",
            )
            for mid, value in reference_metric_values.items()
        ],
        manifest,
    )

    status_counts = Counter(row["comparison_status"] for row in equivalence_rows)
    cache_matches_target = sum(
        1
        for row in equivalence_rows
        if row["target_reference_value"] is not None
        and row["source_reference_value_cached"] == row["target_reference_value"]
    )
    zero_counts = derived_policy["observed"]

    sha_after = {
        p.name: _sha256(p)
        for p in (export, generator, template, base_template)
        if p.is_file()
    }
    hashes_unchanged = sha_before == sha_after

    summary = {
        "generated_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "producer_version": PRODUCER_VERSION,
        "export_name": export.name,
        "generator_name": generator.name,
        "template_name": template.name,
        "base_template_name": base_template.name if base_template.is_file() else "",
        "source_sha256_before": sha_before,
        "source_sha256_after": sha_after,
        "source_hash_verified_unchanged": hashes_unchanged,
        "estado_filter_status": ESTADO_FILTER_STATUS,
        "write_manifest_instruction_count": len(manifest),
        "unresolved_source_metric_ids": unresolved_ids,
        "legacy_active_records": reference.active_records,
        "reference_metric_values_available": len(reference_metric_values),
        "reference_value_matches": status_counts.get(CMP_MATCH, 0),
        "reference_value_matches_vs_legacy_cache": cache_matches_target,
        "zero_vs_blank": status_counts.get(CMP_ZERO_VS_BLANK, 0),
        "mismatches": status_counts.get(CMP_MISMATCH, 0),
        "unavailable": status_counts.get(CMP_TARGET_UNAVAILABLE, 0)
        + status_counts.get(CMP_SOURCE_UNAVAILABLE, 0),
        "production_period": period.label,
        "production_scope_records": production_run.scope_records,
        "production_input_scope": production_run.input_scope,
        "production_estado_filter_applied": production_run.estado_filter_applied,
        "production_metric_values": len(production_run.metric_values),
        "production_write_ready_values": production_run.write_ready_values(manifest),
        "production_missing_values": len(production_completeness.missing),
        "production_value_type_conflicts": len(production_run.value_type_conflicts),
        "production_unsupported": len(production_run.unsupported),
        "production_completeness": {
            "missing": len(production_completeness.missing),
            "duplicate": len(production_completeness.duplicate),
            "orphan": len(production_completeness.orphan),
            "ok": production_completeness.ok,
        },
        "diagnostic_scope_records": (
            diagnostic_equivalence["diagnostic_scope_records"] if diagnostic_equivalence else None
        ),
        "diagnostic_equivalence_status": (
            diagnostic_equivalence["status"] if diagnostic_equivalence else "NOT_RUN"
        ),
        "diagnostic_note": (
            f"scope {diagnostic_equivalence['diagnostic_scope_records']} "
            f"(esperado {diagnostic_equivalence['expected_legacy_active_records']}); "
            f"{diagnostic_equivalence['equal_to_legacy']}/"
            f"{diagnostic_equivalence['metric_values']} MetricValues iguales al valor "
            f"legacy reevaluado; estado {diagnostic_equivalence['status']}. "
            "La hipótesis ESTADO sigue PENDING_FUNCTIONAL_CONFIRMATION; no se "
            "convierte en regla productiva."
            if diagnostic_equivalence
            else "No ejecutado (usar --diagnostic)."
        ),
        "zero_source_cases": zero_counts["zero_source_cases"],
        "zero_semantics": {
            "source_zero_target_zero": zero_counts["source_zero_target_zero"],
            "source_zero_target_blank": zero_counts["source_zero_target_blank"],
            "source_zero_target_other": zero_counts["source_zero_target_other"],
        },
        "zero_write_policy": derived_policy["resolution"],
        "zero_write_policy_evidence": derived_policy["observed"],
        "zero_write_policy_config": zero_policy_check,
        "mismatch_cluster_count": len(mismatch_clusters),
        "pending_writes_production": len(pending_production),
        "pending_writes_reference": len(pending_reference),
    }
    if diagnostic_equivalence is not None:
        summary["diagnostic_detail"] = diagnostic_equivalence

    payload = {
        "production_run": production_run,
        "diagnostic_run": diagnostic_run,
        "equivalence_rows": equivalence_rows,
        "equivalence_summary": equivalence_summary,
        "zero_rows": zero_rows,
        "mismatch_rows": mismatch_rows,
        "mismatch_clusters": mismatch_clusters,
        "pending_production": _pending_rows(pending_production),
        "pending_reference": _pending_rows(pending_reference),
        "summary": summary,
    }
    payload["readme"] = _readme(summary)
    return payload


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="build_metric_values.py",
        description=(
            "Productor de MetricValue MEDINET y equivalencia de valores. "
            "Sólo lectura, sin COM ni macros, sin PII."
        ),
    )
    parser.add_argument("export", nargs="?", default=DEFAULT_EXPORT)
    parser.add_argument("generator", nargs="?", default=DEFAULT_GENERATOR)
    parser.add_argument("template", nargs="?", default=DEFAULT_TEMPLATE)
    parser.add_argument("--base-template", default=DEFAULT_BASE_TEMPLATE)
    parser.add_argument("--manifest", default=DEFAULT_MANIFEST)
    parser.add_argument("--period", default="2026-07", help="AAAA-MM")
    parser.add_argument(
        "--diagnostic",
        action="store_true",
        help="ejecuta también LEGACY_EQUIVALENCE_DIAGNOSTIC (fuera del flujo productivo)",
    )
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        period = _parse_period(args.period)
        payload = build_analysis(
            args.export,
            args.generator,
            args.template,
            args.base_template,
            args.manifest,
            period,
            run_diagnostic=args.diagnostic,
        )
    except RemasepError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    files = write_outputs(Path(args.output), payload)
    s = payload["summary"]
    print(f"export: {s['export_name']}   período: {s['production_period']}")
    print(f"  hash de fuentes verificado sin cambios: {s['source_hash_verified_unchanged']}")
    print(f"  producer_version: {s['producer_version']}")
    print(
        f"  PRODUCCIÓN: scope={s['production_scope_records']} "
        f"input_scope={s['production_input_scope']} "
        f"estado_filter_applied={s['production_estado_filter_applied']}"
    )
    print(
        f"    MetricValues={s['production_metric_values']} "
        f"write_ready={s['production_write_ready_values']}/{s['write_manifest_instruction_count']} "
        f"missing={s['production_missing_values']} "
        f"value_type_conflicts={s['production_value_type_conflicts']} "
        f"completitud_ok={s['production_completeness']['ok']}"
    )
    print(
        f"  EQUIVALENCIA REF: MATCH={s['reference_value_matches']} "
        f"ZERO_VS_BLANK={s['zero_vs_blank']} MISMATCH={s['mismatches']} "
        f"UNAVAILABLE={s['unavailable']}"
    )
    print(
        f"  CERO: casos={s['zero_source_cases']} zero/blank/other="
        f"{s['zero_semantics']['source_zero_target_zero']}/"
        f"{s['zero_semantics']['source_zero_target_blank']}/"
        f"{s['zero_semantics']['source_zero_target_other']}  "
        f"zero_write_policy={s['zero_write_policy']}"
    )
    print(
        f"    config zero_write_policy={s['zero_write_policy_config']['config_resolution']} "
        f"(coincide con evidencia: {s['zero_write_policy_config']['config_matches_evidence']})"
    )
    print(f"  mismatch clusters: {s['mismatch_cluster_count']}")
    print(
        f"  DIAGNÓSTICO: {s['diagnostic_equivalence_status']} "
        f"scope={s['diagnostic_scope_records']}"
    )
    print(f"  estado_filter_status: {s['estado_filter_status']}")
    print(f"Artefactos en {args.output}:")
    for name in files:
        print(f"  - {name}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
