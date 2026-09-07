"""Mapeo semántico de las métricas MEDINET (Sprint 3.2).

Parte del inventario de Sprint 3.1 (`scripts/inventory_semantic_metrics.py`) y lo
convierte en un modelo `SemanticMetric` con dimensiones **explícitas y
auditables** — sexo, edad, alcance de agregación, código de procedimiento — cada
una con `value`, `status` y `evidence`.

- **No inventa semántica**: sin evidencia suficiente → `LABEL_ONLY` /
  `FORMULA_ONLY` / `NOT_APPLICABLE` / `UNRESOLVED`.
- Vocabulario explícito versionado en `config/semantic_mapping_2026/`
  (`status: preliminary_pending_functional_validation`).
- `source = "MEDINET"` para toda métrica. El modelo queda preparado
  conceptualmente para `EGRESOS` / `RESOURCE_CALCULATION` (no implementadas).
- **No** mapea a la plantilla oficial MINSAL. No es validación funcional.

Uso:

    python scripts/map_semantic_metrics.py \\
        "data/local/GENERACION DATOS REMASEP.xlsx"
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from datetime import UTC, datetime
from pathlib import Path

import close_legacy_aggregations as clo
import inventory_semantic_metrics as si

from remasep.services import semantic_mapping as sm

_ROW_RE = re.compile(r"[A-Za-z]+(\d+)")

DEFAULT_OUTPUT = "artifacts/semantic_metric_mapping"
_SEX_STATES = (sm.CONFIRMED, sm.LABEL_ONLY, sm.FORMULA_ONLY, sm.NOT_APPLICABLE,
               sm.UNRESOLVED, sm.CONFLICT)
_AGE_STATES = _SEX_STATES
_PROC_STATES = (sm.PROC_EXPLICIT, sm.NOT_APPLICABLE, sm.UNRESOLVED)


def _metric_input(mc: si.MetricContext, layout) -> sm.MetricInput:
    formula_age = (
        (mc.age_lower_bound, mc.age_upper_bound)
        if mc.age_lower_bound is not None or mc.age_upper_bound is not None
        else None
    )
    label_age = (
        (mc.label_age_lower, mc.label_age_upper) if mc.column_age_label else None
    )
    own_row_is_header = False
    if layout is not None:
        match = _ROW_RE.match(mc.cell)
        if match:
            row = int(match.group(1))
            own_row_is_header = row in layout._section_rows or row in layout._group_rows
    return sm.MetricInput(
        metric_id=mc.metric_id,
        sheet=mc.sheet,
        cell=mc.cell,
        kind=mc.kind,
        value=mc.value,
        section_path=mc.section_labels_raw,
        row_path=mc.row_labels_raw,
        column_path=mc.column_labels_raw,
        text_criteria=mc.text_criteria,
        formula=mc.formula,
        formula_age=formula_age,
        label_age=label_age,
        label_age_text=mc.column_age_label,
        semantic_signature=mc.semantic_signature,
        context_status=mc.context_status,
        own_row_is_header=own_row_is_header,
    )


def build_mapping(path: str | Path):
    inventory = si.build_inventory(path)
    config = sm.load_mapping_config()
    metrics = [
        sm.map_metric(_metric_input(mc, inventory.layouts.get(mc.sheet)), config)
        for mc in inventory.metrics
    ]
    return inventory, config, metrics


# ---------------------------------------------------------------------------
# Artefactos
# ---------------------------------------------------------------------------


def _csv(path: Path, header: list[str], rows) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def _fmt(value: object) -> str:
    return "" if value is None else str(value)


def _write_semantic_metrics(metrics: list[sm.SemanticMetric], out: Path) -> Path:
    path = out / "semantic_metrics.csv"
    header = [
        "metric_id", "source", "form", "sheet", "cell", "kind", "value",
        "section_path_raw", "row_path_raw", "column_path_raw",
        "sex_value", "sex_status",
        "age_min_years", "age_max_years", "age_status",
        "aggregation_scope", "aggregation_scope_status",
        "procedure_code_raw", "procedure_label_raw", "procedure_status",
        "mapping_status", "semantic_signature",
    ]
    rows = (
        [
            m.metric_id, m.source, m.form, m.sheet, m.cell, m.kind, _fmt(m.value),
            " :: ".join(m.section_path_raw), " :: ".join(m.row_path_raw),
            " :: ".join(m.column_path_raw),
            m.sex_value, m.sex_status,
            _fmt(m.age_min_years), _fmt(m.age_max_years), m.age_status,
            m.aggregation_scope, m.aggregation_scope_status,
            m.procedure_code_raw, m.procedure_label_raw, m.procedure_status,
            m.mapping_status, m.semantic_signature,
        ]
        for m in metrics
    )
    _csv(path, header, rows)
    return path


def _write_dimension_evidence(metrics: list[sm.SemanticMetric], out: Path) -> Path:
    path = out / "dimension_evidence.csv"
    header = ["metric_id", "dimension", "evidence_type", "raw_value", "normalized_value",
              "evidence_status"]
    rows = (
        [m.metric_id, e.dimension, e.evidence_type, e.raw_value, e.normalized_value,
         e.evidence_status]
        for m in metrics
        for e in m.evidence
    )
    _csv(path, header, rows)
    return path


def _write_conflicts(metrics: list[sm.SemanticMetric], out: Path) -> Path:
    path = out / "mapping_conflicts.csv"
    header = ["metric_id", "dimension", "label_evidence", "formula_evidence", "conflict_type"]
    rows = (
        [m.metric_id, c.dimension, c.label_evidence, c.formula_evidence, c.conflict_type]
        for m in metrics
        for c in m.conflicts
    )
    _csv(path, header, rows)
    return path


def _write_procedure_codes(metrics: list[sm.SemanticMetric], out: Path) -> Path:
    path = out / "procedure_codes.csv"
    header = ["metric_id", "sheet", "cell", "procedure_code_raw", "procedure_label_raw",
              "procedure_status", "row_path_raw"]
    rows = (
        [m.metric_id, m.sheet, m.cell, m.procedure_code_raw, m.procedure_label_raw,
         m.procedure_status, " :: ".join(m.row_path_raw)]
        for m in metrics
        if m.procedure_code_raw
    )
    _csv(path, header, rows)
    return path


def _coverage_row(scope: str, metrics: list[sm.SemanticMetric]) -> list:
    ms = Counter(m.mapping_status for m in metrics)
    sx = Counter(m.sex_status for m in metrics)
    ag = Counter(m.age_status for m in metrics)
    pr = Counter(m.procedure_status for m in metrics)
    return [
        scope, len(metrics),
        ms.get(sm.MAPPING_CONFIRMED, 0), ms.get(sm.MAPPING_PARTIAL, 0),
        ms.get(sm.MAPPING_CONFLICT, 0),
        sx.get(sm.CONFIRMED, 0), sx.get(sm.LABEL_ONLY, 0), sx.get(sm.FORMULA_ONLY, 0),
        sx.get(sm.NOT_APPLICABLE, 0), sx.get(sm.UNRESOLVED, 0), sx.get(sm.CONFLICT, 0),
        ag.get(sm.CONFIRMED, 0), ag.get(sm.LABEL_ONLY, 0), ag.get(sm.FORMULA_ONLY, 0),
        ag.get(sm.NOT_APPLICABLE, 0), ag.get(sm.UNRESOLVED, 0), ag.get(sm.CONFLICT, 0),
        pr.get(sm.PROC_EXPLICIT, 0), pr.get(sm.NOT_APPLICABLE, 0), pr.get(sm.UNRESOLVED, 0),
    ]


def _write_coverage(metrics: list[sm.SemanticMetric], out: Path) -> Path:
    path = out / "mapping_coverage.csv"
    header = [
        "scope", "metric_count",
        "mapping_confirmed", "mapping_partial", "mapping_conflict",
        "sex_confirmed", "sex_label_only", "sex_formula_only", "sex_not_applicable",
        "sex_unresolved", "sex_conflict",
        "age_confirmed", "age_label_only", "age_formula_only", "age_not_applicable",
        "age_unresolved", "age_conflict",
        "procedure_explicit", "procedure_not_applicable", "procedure_unresolved",
    ]
    grouped: dict[tuple[str, str], list[sm.SemanticMetric]] = defaultdict(list)
    for m in metrics:
        grouped[(m.form, m.kind)].append(m)
    rows = [_coverage_row(f"{form} / {kind}", g) for (form, kind), g in sorted(grouped.items())]
    rows.append(_coverage_row("TOTAL", metrics))
    _csv(path, header, rows)
    return path


def _manual_review_sample(metrics: list[sm.SemanticMetric]) -> list[sm.SemanticMetric]:
    ordered = sorted(metrics, key=lambda m: m.metric_id)
    buckets = [
        lambda m: m.form == "REMASEP_01",
        lambda m: m.form == "REMASEP_OD",
        lambda m: m.form == "B2_ANEXO",
        lambda m: m.kind == "BASE_AGGREGATION",
        lambda m: m.kind == "DERIVED_AGGREGATION",
        lambda m: m.kind == "DOWNSTREAM_TOTAL",
        lambda m: m.sex_value == sm.SEX_MALE,
        lambda m: m.sex_value == sm.SEX_FEMALE,
        lambda m: m.sex_value == sm.SEX_BOTH,
        lambda m: bool(m.procedure_code_raw),
        lambda m: m.mapping_status == sm.MAPPING_PARTIAL,
        lambda m: m.mapping_status == sm.MAPPING_CONFLICT,
        lambda m: m.age_min_years == 0,
        lambda m: m.age_min_years == 20,
        lambda m: m.age_min_years is not None and m.age_max_years is None,
        lambda m: m.aggregation_scope == sm.SCOPE_TOTAL,
        lambda m: m.aggregation_scope == sm.SCOPE_SUBTOTAL,
    ]
    picked: dict[str, sm.SemanticMetric] = {}
    for predicate in buckets:
        hits = [m for m in ordered if predicate(m)]
        for m in hits[:3]:
            picked[m.metric_id] = m
    # todos los CONFLICT, sin límite
    for m in ordered:
        if m.mapping_status == sm.MAPPING_CONFLICT:
            picked[m.metric_id] = m
    return sorted(picked.values(), key=lambda m: m.metric_id)


def _write_manual_review_sample(metrics: list[sm.SemanticMetric], out: Path) -> Path:
    path = out / "manual_review_sample.csv"
    header = [
        "metric_id", "source", "form", "kind", "cell", "row_path_raw", "column_path_raw",
        "sex_value", "sex_status", "age_min_years", "age_max_years", "age_status",
        "procedure_code_raw", "procedure_label_raw", "aggregation_scope", "mapping_status",
    ]
    rows = (
        [
            m.metric_id, m.source, m.form, m.kind, m.cell, " :: ".join(m.row_path_raw),
            " :: ".join(m.column_path_raw), m.sex_value, m.sex_status,
            _fmt(m.age_min_years), _fmt(m.age_max_years), m.age_status,
            m.procedure_code_raw, m.procedure_label_raw, m.aggregation_scope,
            m.mapping_status,
        ]
        for m in _manual_review_sample(metrics)
    )
    _csv(path, header, rows)
    return path


_README = """# Semantic metric mapping (Sprint 3.2)

Salida de `scripts/map_semantic_metrics.py`. Convierte el *contexto* preliminar
de Sprint 3.1 en **dimensiones explícitas** de cada métrica MEDINET, cada una
con `value`, `status` y `evidence`.

> **No inventa semántica.** Sin evidencia suficiente la dimensión queda
> `LABEL_ONLY` / `FORMULA_ONLY` / `NOT_APPLICABLE` / `UNRESOLVED`, nunca un valor
> "adivinado". **No** es un mapping a la plantilla oficial MINSAL ni una
> validación funcional.

## Dimensiones

| Dimensión | Evidencia | Estados |
| --- | --- | --- |
| `form` | nombre de hoja (`config/semantic_mapping_2026/form_labels.yaml`) | — |
| `sex` | rótulo de columna (`Hombres`/`Mujeres`/`Ambos sexos`) + criterio de fórmula (`"*Hombre*"`) | `CONFIRMED` / `LABEL_ONLY` / `FORMULA_ONLY` / `NOT_APPLICABLE` / `UNRESOLVED` / `CONFLICT` |
| `age` (`age_min_years`, `age_max_years`) | rótulo de edad (Sprint 3.1) + cotas `$AF:$AF` de la fórmula | idem |
| `aggregation_scope` | rótulo `TOTAL` / fila-encabezado de grupo | `DETAIL` / `TOTAL` / `SUBTOTAL`, con `status` |
| `procedure_code_raw` / `procedure_label_raw` | código explícito en un rótulo de fila (`5010009 - VIDRIO IONÓMERO`, o `0601105` + nombre) | `EXPLICIT` / `NOT_APPLICABLE` / `UNRESOLVED` |

`age_max_years` vacío = extremo abierto (`75 y más años`).

## `mapping_status` (≠ `context_status` de Sprint 3.1)

- `CONFLICT` — alguna dimensión tiene `CONFLICT`.
- `CONFIRMED` — `sex` y `age` están `CONFIRMED` o justificadamente
  `NOT_APPLICABLE`, **y** `aggregation_scope` está `CONFIRMED`.
- `PARTIAL` — alguna dimensión queda `LABEL_ONLY` / `FORMULA_ONLY` /
  `UNRESOLVED`. Es información válida, no un fallo.

Nunca se usa `COMPLETE` (para no confundir con `context_status`).

## Archivos

- `semantic_metrics.csv` — una fila por métrica con todas las dimensiones.
- `dimension_evidence.csv` — una fila por evidencia (`COLUMN_LABEL` /
  `FORMULA_CRITERION` / `FORMULA_BOUNDS` / `ROW_LABEL` / `SHEET`).
- `mapping_coverage.csv` — cobertura por `form` × `kind`.
- `mapping_conflicts.csv` — conflictos rótulo ↔ fórmula (sin resolver).
- `procedure_codes.csv` — códigos de procedimiento explícitos encontrados.
- `manual_review_sample.csv` — muestra determinista y estratificada para
  revisión humana (`metric_id`, `source`, `form`, `kind`, …). Su tamaño exacto
  está en `summary.json` → `manual_review_sample_size`. No se ajustan mappings
  automáticamente en función de ella.
- `summary.json` — totales.

## Provenance

`source = "MEDINET"` en todas las métricas. Preparado para `EGRESOS` y
`RESOURCE_CALCULATION` (sin código en este sprint).

## Privacidad

Sólo rótulos del formulario, coordenadas y valores agregados. Ninguna fila
Medinet individual.
"""


def _summary(inventory, config: sm.MappingConfig, metrics: list[sm.SemanticMetric]) -> dict:
    ms = Counter(m.mapping_status for m in metrics)
    sx = Counter(m.sex_status for m in metrics)
    ag = Counter(m.age_status for m in metrics)
    pr = Counter(m.procedure_status for m in metrics)
    scope = Counter(f"{m.aggregation_scope}:{m.aggregation_scope_status}" for m in metrics)
    b2 = [m for m in metrics if m.form == "B2_ANEXO"]
    proc_codes = sorted({m.procedure_code_raw for m in metrics if m.procedure_code_raw})
    sample = _manual_review_sample(metrics)
    return {
        "source": inventory.source_name,
        "source_sha256": inventory.source_sha256,
        "generated_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "mapping_config_version": config.version,
        "mapping_status": "PRELIMINARY_NOT_VALIDATED",
        "total_semantic_metrics": len(metrics),
        "by_form": dict(sorted(Counter(m.form for m in metrics).items())),
        "by_kind": dict(sorted(Counter(m.kind for m in metrics).items())),
        "mapping_confirmed": ms.get(sm.MAPPING_CONFIRMED, 0),
        "mapping_partial": ms.get(sm.MAPPING_PARTIAL, 0),
        "mapping_conflict": ms.get(sm.MAPPING_CONFLICT, 0),
        "sex_status": {s: sx.get(s, 0) for s in _SEX_STATES},
        "age_status": {s: ag.get(s, 0) for s in _AGE_STATES},
        "procedure_status": {s: pr.get(s, 0) for s in _PROC_STATES},
        "aggregation_scope": dict(sorted(scope.items())),
        "distinct_procedure_codes": len(proc_codes),
        "b2_anexo_metrics": len(b2),
        "b2_anexo_mapping_status": dict(sorted(Counter(m.mapping_status for m in b2).items())),
        "conflicts_total": sum(len(m.conflicts) for m in metrics),
        "manual_review_sample_size": len(sample),
        "manual_review_sample_by_status": dict(
            sorted(Counter(m.mapping_status for m in sample).items())
        ),
    }


def write_outputs(inventory, config: sm.MappingConfig, metrics: list[sm.SemanticMetric],
                  output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    files = [
        _write_semantic_metrics(metrics, output_dir),
        _write_dimension_evidence(metrics, output_dir),
        _write_coverage(metrics, output_dir),
        _write_conflicts(metrics, output_dir),
        _write_procedure_codes(metrics, output_dir),
        _write_manual_review_sample(metrics, output_dir),
    ]
    (output_dir / "summary.json").write_text(
        json.dumps(_summary(inventory, config, metrics), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (output_dir / "README.md").write_text(_README, encoding="utf-8")
    return [*files, output_dir / "summary.json", output_dir / "README.md"]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="map_semantic_metrics.py",
        description=(
            "Mapeo semántico de las métricas MEDINET: dimensiones explícitas "
            "(sexo, edad, alcance, código de procedimiento) con evidencia. Sólo lectura."
        ),
    )
    parser.add_argument("workbook", help="Ruta al workbook legacy (.xlsx/.xlsm).")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"(def: {DEFAULT_OUTPUT})")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        inventory, config, metrics = build_mapping(args.workbook)
    except (clo.ClosureError, sm.SemanticMappingError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    files = write_outputs(inventory, config, metrics, Path(args.output))
    summary = _summary(inventory, config, metrics)

    print(f"Workbook: {summary['source']}")
    print(f"  SemanticMetric: {summary['total_semantic_metrics']}  por hoja {summary['by_form']}")
    print(
        f"  mapping: {summary['mapping_confirmed']} CONFIRMED / "
        f"{summary['mapping_partial']} PARTIAL / {summary['mapping_conflict']} CONFLICT"
    )
    print(f"  sex_status:  {summary['sex_status']}")
    print(f"  age_status:  {summary['age_status']}")
    print(f"  aggregation_scope: {summary['aggregation_scope']}")
    print(
        f"  procedure: {summary['procedure_status']} "
        f"({summary['distinct_procedure_codes']} códigos distintos)"
    )
    print(
        f"  B2 ANEXO: {summary['b2_anexo_metrics']} métricas "
        f"({summary['b2_anexo_mapping_status']})"
    )
    print(f"  conflictos: {summary['conflicts_total']}")
    print(
        f"  manual_review_sample: {summary['manual_review_sample_size']} filas "
        f"{summary['manual_review_sample_by_status']}"
    )
    print(f"Artefactos en {args.output}:")
    for path in files:
        print(f"  - {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
