"""Inventario semántico preliminar de las métricas legacy (Sprint 3.1).

Empieza la **capa semántica**. Sprint 0–2.5 cerraron la reproducción matemática
del subgrafo Medinet del workbook legacy: 2043/2043 celdas evaluables. Ahora, sin
ampliar el parser de Excel, se descubre **qué significa** cada métrica a partir
del *layout visual* del workbook:

    metric_id técnico  +  rótulos de fila  +  rótulos de columna  +  secciones
                       +  evidencia de fórmula (edad / criterios de texto)

- **No** asigna códigos clínicos (sexo, grupo de edad, actividad, especialidad):
  eso es Sprint 3.2.
- **No** mapea a la plantilla oficial.
- Sólo lectura; no ejecuta Excel, COM ni macros; no modifica el workbook.
- `source = "MEDINET"` para toda métrica de este inventario. El modelo futuro
  tendrá también `EGRESOS` y `RESOURCE_CALCULATION` (no implementadas aquí).

Uso:

    python scripts/inventory_semantic_metrics.py \\
        "data/local/GENERACION DATOS REMASEP.xlsx"
"""

from __future__ import annotations

import argparse
import csv
import json
import re
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path

import close_legacy_aggregations as clo
import evaluate_legacy_downstream as ev
import openpyxl

from remasep.services.semantic_layout import (
    CONTEXT_AMBIGUOUS,
    CONTEXT_COMPLETE,
    CONTEXT_NO_CONTEXT,
    CONTEXT_PARTIAL,
    SheetLayout,
    age_bounds_from_formula,
    compare_label_and_formula_age,
    normalize_semantic_label,
    parse_age_label,
    text_criteria_from_formula,
)

DEFAULT_OUTPUT = "artifacts/semantic_metric_inventory"
SOURCE = "MEDINET"
METRIC_KINDS = (clo.KIND_BASE, clo.KIND_DERIVED, clo.KIND_TOTAL)
_CONTEXT_ORDER = (CONTEXT_COMPLETE, CONTEXT_PARTIAL, CONTEXT_AMBIGUOUS, CONTEXT_NO_CONTEXT)
_MANUAL_FIELD_RE = re.compile(r"\b(AN|AO)\d+\b")
_MANUAL_FIELD_NAME = {"AN": "EMBARAZADAS", "AO": "MIGRANTES"}


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------


@dataclass
class MetricContext:
    metric_id: str
    source: str
    sheet: str
    cell: str
    kind: str
    dependency_depth: int
    value: object
    depends_on_age: bool
    section_labels_raw: tuple[str, ...]
    row_labels_raw: tuple[str, ...]
    column_labels_raw: tuple[str, ...]
    semantic_signature: str
    context_status: str
    notes: tuple[str, ...] = ()
    formula: str = ""
    # evidencia de fórmula
    referenced_detail_columns: tuple[str, ...] = ()
    text_criteria: tuple[str, ...] = ()
    age_lower_bound: int | None = None
    age_upper_bound: int | None = None
    formula_evidence_status: str = "NONE"
    # consistencia rótulo/fórmula
    column_age_label: str = ""
    label_age_lower: int | None = None
    label_age_upper: int | None = None
    consistency_status: str = "NO_COMPARABLE"

    @property
    def age_comparable(self) -> bool:
        """¿Tiene evidencia de edad en al menos un lado (rótulo o fórmula)?"""
        return bool(self.column_age_label) or self.age_lower_bound is not None or (
            self.age_upper_bound is not None
        )


@dataclass
class InventoryResult:
    source_name: str
    source_sha256: str
    generated_at: str
    formula_support_status: str
    metrics: list[MetricContext]
    validations: list[dict]
    layouts: dict[str, SheetLayout] = field(default_factory=dict)

    def by_sheet_kind(self) -> dict[tuple[str, str], list[MetricContext]]:
        grouped: dict[tuple[str, str], list[MetricContext]] = defaultdict(list)
        for m in self.metrics:
            grouped[(m.sheet, m.kind)].append(m)
        return grouped


# ---------------------------------------------------------------------------
# Construcción
# ---------------------------------------------------------------------------


def _norm_tuple(values: tuple[str, ...]) -> tuple[str, ...]:
    return tuple(normalize_semantic_label(v) for v in values)


def _formula_evidence(node: clo.Node, ctx) -> dict:
    formula = node.formula if isinstance(node.formula, str) else ""
    age = age_bounds_from_formula(formula)
    text_criteria = text_criteria_from_formula(formula)
    lower, upper = (age or (None, None))

    if node.kind == clo.KIND_TOTAL:
        status = "AGGREGATE_ONLY"
    elif age and text_criteria:
        status = "AGE_AND_TEXT"
    elif age:
        status = "AGE_BOUNDS"
    elif text_criteria:
        status = "TEXT_ONLY"
    else:
        status = "NONE"

    # comparación objetiva rótulo de columna vs. fórmula
    column_age_label = ""
    label_bounds = None
    for label in ctx.column_labels_raw:
        parsed = parse_age_label(label)
        if parsed is not None:
            column_age_label = label
            label_bounds = parsed
            break
    consistency = compare_label_and_formula_age(label_bounds, age)

    return {
        "referenced_detail_columns": tuple(node.detail_columns),
        "text_criteria": text_criteria,
        "age_lower_bound": lower,
        "age_upper_bound": upper,
        "formula_evidence_status": status,
        "column_age_label": column_age_label,
        "label_age_lower": label_bounds[0] if label_bounds else None,
        "label_age_upper": label_bounds[1] if label_bounds else None,
        "consistency_status": consistency,
    }


def build_inventory(path: str | Path) -> InventoryResult:
    closure = ev.evaluate_full_closure(path)
    source = Path(path)

    wb = openpyxl.load_workbook(source, data_only=True, read_only=False)
    try:
        present = [s for s in clo.TARGET_SHEETS if s in wb.sheetnames]
        layouts = {name: SheetLayout(wb[name]) for name in present}
    finally:
        wb.close()

    metrics: list[MetricContext] = []
    validations: list[dict] = []

    for node in sorted(closure.medinet_nodes(), key=lambda n: clo._key_sort(n.key)):
        layout = layouts.get(node.sheet)
        if layout is None:
            continue

        if node.kind == clo.KIND_VALIDATION:
            ctx = layout.cell_context(node.cell)
            manual = sorted(
                {_MANUAL_FIELD_NAME[m] for m in _MANUAL_FIELD_RE.findall(node.formula or "")}
            )
            validations.append(
                {
                    "validation_id": node.metric_id,
                    "sheet": node.sheet,
                    "cell": node.cell,
                    "row_context": " :: ".join(ctx.row_labels_raw),
                    "column_context": " :: ".join(ctx.column_labels_raw),
                    "section_context": " :: ".join(ctx.section_labels_raw),
                    "references_manual_field": "|".join(manual),
                    "context_status": ctx.context_status,
                    "formula": node.formula,
                    "python_value": node.python_value,
                }
            )
            continue

        if node.kind not in METRIC_KINDS:
            continue

        ctx = layout.cell_context(node.cell)
        evidence = _formula_evidence(node, ctx)
        metrics.append(
            MetricContext(
                metric_id=node.metric_id,
                source=SOURCE,
                sheet=node.sheet,
                cell=node.cell,
                kind=node.kind,
                dependency_depth=node.depth,
                value=node.python_value,
                depends_on_age=node.depends_on_age,
                section_labels_raw=ctx.section_labels_raw,
                row_labels_raw=ctx.row_labels_raw,
                column_labels_raw=ctx.column_labels_raw,
                semantic_signature=ctx.semantic_signature(),
                context_status=ctx.context_status,
                notes=ctx.notes,
                formula=node.formula if isinstance(node.formula, str) else "",
                **evidence,
            )
        )

    return InventoryResult(
        source_name=source.name,
        source_sha256=closure.closure.source_sha256,
        generated_at=datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        formula_support_status=closure.closure.formula_support_status,
        metrics=metrics,
        validations=validations,
        layouts=layouts,
    )


# ---------------------------------------------------------------------------
# Artefactos
# ---------------------------------------------------------------------------


def _csv(path: Path, header: list[str], rows) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def _leveled(values: tuple[str, ...], count: int) -> list[str]:
    padded = list(values[:count]) + [""] * (count - len(values[:count]))
    out: list[str] = []
    for raw in padded:
        out.append(raw)
        out.append(normalize_semantic_label(raw))
    return out


def _fmt(value: object) -> str:
    return "" if value is None else str(value)


def _write_candidates(result: InventoryResult, out: Path) -> Path:
    path = out / "semantic_metric_candidates.csv"
    header = [
        "metric_id", "source", "sheet", "cell", "kind", "dependency_depth", "value",
        "depends_on_AF",
        "section_label_1_raw", "section_label_1_normalized",
        "section_label_2_raw", "section_label_2_normalized",
        "row_label_1_raw", "row_label_1_normalized",
        "row_label_2_raw", "row_label_2_normalized",
        "row_label_3_raw", "row_label_3_normalized",
        "column_label_1_raw", "column_label_1_normalized",
        "column_label_2_raw", "column_label_2_normalized",
        "column_label_3_raw", "column_label_3_normalized",
        "semantic_signature", "context_status", "context_notes",
    ]
    rows = (
        [
            m.metric_id, m.source, m.sheet, m.cell, m.kind, m.dependency_depth, _fmt(m.value),
            "True" if m.depends_on_age else "False",
            *_leveled(m.section_labels_raw, 2),
            *_leveled(m.row_labels_raw, 3),
            *_leveled(m.column_labels_raw, 3),
            m.semantic_signature, m.context_status, " | ".join(m.notes),
        ]
        for m in result.metrics
    )
    _csv(path, header, rows)
    return path


def _write_validations(result: InventoryResult, out: Path) -> Path:
    path = out / "validation_context.csv"
    header = [
        "validation_id", "sheet", "cell", "row_context", "column_context",
        "section_context", "references_manual_field", "context_status", "formula",
        "python_value",
    ]
    rows = (
        [
            v["validation_id"], v["sheet"], v["cell"], v["row_context"], v["column_context"],
            v["section_context"], v["references_manual_field"], v["context_status"],
            v["formula"], _fmt(v["python_value"]),
        ]
        for v in result.validations
    )
    _csv(path, header, rows)
    return path


def _write_formula_evidence(result: InventoryResult, out: Path) -> Path:
    path = out / "formula_evidence.csv"
    header = [
        "metric_id", "sheet", "cell", "kind", "referenced_detail_columns", "text_criteria",
        "age_lower_bound", "age_upper_bound", "criteria_count", "formula_evidence_status",
    ]
    rows = (
        [
            m.metric_id, m.sheet, m.cell, m.kind, "|".join(m.referenced_detail_columns),
            "|".join(m.text_criteria), _fmt(m.age_lower_bound), _fmt(m.age_upper_bound),
            len(m.text_criteria), m.formula_evidence_status,
        ]
        for m in result.metrics
    )
    _csv(path, header, rows)
    return path


def _write_consistency(result: InventoryResult, out: Path) -> Path:
    path = out / "label_formula_consistency.csv"
    header = [
        "metric_id", "sheet", "cell", "column_age_label_raw", "label_age_lower",
        "label_age_upper", "formula_age_lower", "formula_age_upper", "consistency_status",
    ]
    rows = (
        [
            m.metric_id, m.sheet, m.cell, m.column_age_label, _fmt(m.label_age_lower),
            _fmt(m.label_age_upper), _fmt(m.age_lower_bound), _fmt(m.age_upper_bound),
            m.consistency_status,
        ]
        for m in result.metrics
        if m.age_comparable
    )
    _csv(path, header, rows)
    return path


def _write_duplicates(result: InventoryResult, out: Path) -> Path:
    path = out / "duplicate_semantic_signatures.csv"
    groups: dict[str, list[MetricContext]] = defaultdict(list)
    for m in result.metrics:
        groups[m.semantic_signature].append(m)
    header = ["semantic_signature", "metric_count", "metric_ids", "cells", "status"]
    rows = (
        [
            sig, len(members),
            "|".join(m.metric_id for m in members),
            "|".join(f"{m.sheet}!{m.cell}" for m in members),
            "UNRESOLVED",
        ]
        for sig, members in sorted(groups.items())
        if len(members) > 1
    )
    _csv(path, header, rows)
    return path


def _write_context_coverage(result: InventoryResult, out: Path) -> Path:
    path = out / "context_coverage.csv"
    header = ["scope", "total", *_CONTEXT_ORDER]

    def _row(label: str, metrics: list[MetricContext]) -> list:
        counts = Counter(m.context_status for m in metrics)
        return [label, len(metrics), *(counts.get(s, 0) for s in _CONTEXT_ORDER)]

    rows = [_row(f"{sheet} / {kind}", members) for (sheet, kind), members in
            sorted(result.by_sheet_kind().items())]
    rows.append(_row("TOTAL", result.metrics))
    _csv(path, header, rows)
    return path


def _write_layouts(result: InventoryResult, out: Path) -> list[Path]:
    paths: list[Path] = []
    for name, layout in result.layouts.items():
        slug = name.replace(" ", "_")
        path = out / f"layout_{slug}.csv"
        header = [
            "row", "col", "cell", "resolved_text_raw", "resolved_text_normalized",
            "is_bold", "is_section_row", "is_group_row",
        ]
        rows = (
            [
                row, col, f"{_col_letter(col)}{row}", text, normalize_semantic_label(text),
                "True" if layout.is_bold(row, col) else "False",
                "True" if row in layout._section_rows else "False",
                "True" if row in layout._group_rows else "False",
            ]
            for row, col, text in layout.text_cells()
        )
        _csv(path, header, rows)
        paths.append(path)
    return paths


def _col_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def _summary_dict(result: InventoryResult) -> dict:
    metrics = result.metrics
    by_sheet = Counter(m.sheet for m in metrics)
    by_kind = Counter(m.kind for m in metrics)
    ctx = Counter(m.context_status for m in metrics)

    sig_groups: dict[str, list[MetricContext]] = defaultdict(list)
    for m in metrics:
        sig_groups[m.semantic_signature].append(m)
    dup_groups = {s: g for s, g in sig_groups.items() if len(g) > 1}

    comparable = [m for m in metrics if m.age_comparable]
    consistency = Counter(m.consistency_status for m in comparable)
    return {
        "source": result.source_name,
        "source_sha256": result.source_sha256,
        "generated_at": result.generated_at,
        "formula_support_status": result.formula_support_status,
        "semantic_layer_status": "PRELIMINARY_NOT_VALIDATED",
        "total_metric_candidates": len(metrics),
        "metrics_by_sheet": dict(sorted(by_sheet.items())),
        "metrics_by_kind": dict(sorted(by_kind.items())),
        "context_status": {s: ctx.get(s, 0) for s in _CONTEXT_ORDER},
        "duplicate_semantic_signatures": len(dup_groups),
        "metrics_in_duplicate_groups": sum(len(g) for g in dup_groups.values()),
        "max_row_hierarchy_depth": max((len(m.row_labels_raw) for m in metrics), default=0),
        "max_column_hierarchy_depth": max((len(m.column_labels_raw) for m in metrics), default=0),
        "max_section_hierarchy_depth": max((len(m.section_labels_raw) for m in metrics), default=0),
        "metrics_with_age_bounds": sum(
            1 for m in metrics if m.age_lower_bound is not None or m.age_upper_bound is not None
        ),
        "metrics_with_text_criteria": sum(1 for m in metrics if m.text_criteria),
        "label_formula_evaluated": len(comparable),
        "label_formula_consistent": consistency.get("CONSISTENT", 0),
        "label_formula_conflicts": consistency.get("CONFLICT", 0),
        "label_formula_no_comparable": consistency.get("NO_COMPARABLE", 0),
        "label_formula_not_applicable": len(metrics) - len(comparable),
        "validation_total": len(result.validations),
        "validation_context_status": {
            s: sum(1 for v in result.validations if v["context_status"] == s)
            for s in _CONTEXT_ORDER
        },
    }


_README = """# Semantic metric inventory (Sprint 3.1)

Salida de `scripts/inventory_semantic_metrics.py`. **Inicio de la capa
semántica**: recupera los rótulos visibles alrededor de cada métrica legacy
(`REMASEP 01`, `B2 ANEXO`, `REMASEP_OD`) para empezar a razonar por significado
en lugar de por coordenada.

> **Esto NO es un semantic mapping validado.** Son *candidatos* extraídos del
> layout del workbook. No hay códigos de sexo, grupo de edad, actividad ni
> especialidad — eso es Sprint 3.2.

## Archivos

- `semantic_metric_candidates.csv` — una fila por métrica (`BASE_AGGREGATION` /
  `DERIVED_AGGREGATION` / `DOWNSTREAM_TOTAL`): `metric_id` técnico, `source`,
  rótulos de sección / fila / columna (raw + normalizado), `semantic_signature`
  candidata y `context_status`.
- `validation_context.csv` — las celdas `VALIDATION` (`IF`) **por separado**: su
  contexto visible, qué campo manual referencian (`EMBARAZADAS`/`MIGRANTES`), la
  fórmula y su valor Python. No se convierten en métricas.
- `formula_evidence.csv` — evidencia **técnica** por métrica: columnas de la hoja
  de detalle, criterios de texto, cotas de edad (`AF >= / <=`).
- `label_formula_consistency.csv` — comparación **objetiva** rótulo de columna de
  edad vs. cotas de la fórmula (`CONSISTENT` / `CONFLICT` / `NO_COMPARABLE`).
- `duplicate_semantic_signatures.csv` — `metric_id` distintos con la misma firma
  (sin resolver).
- `context_coverage.csv` — cobertura de contexto por hoja y tipo.
- `layout_<HOJA>.csv` — volcado del texto resuelto (merges incluidos) de cada
  hoja, para auditar las heurísticas.
- `summary.json` — totales.

## `context_status`

| Estado | Criterio objetivo |
| --- | --- |
| `COMPLETE` | hay ≥ 1 rótulo de fila **y** ≥ 1 rótulo de columna, sin ambigüedad |
| `PARTIAL` | hay contexto de un solo eje (fila **o** columna, y/o sección) |
| `AMBIGUOUS` | la extracción detectó una ambigüedad (p.ej. la banda de encabezados sigue en una columna vecina pero la de la métrica está vacía) |
| `NO_CONTEXT` | no se pudo recuperar ningún rótulo visible |

## `semantic_signature`

`hoja :: sección… :: fila… :: columna…` (normalizado). Es **candidata**: el
`metric_id` técnico (`LEGACY::<hoja>::<celda>`) sigue siendo el identificador.

## `source` / provenance

`source = "MEDINET"` para todo este inventario. El modelo semántico futuro
tendrá además `EGRESOS` y `RESOURCE_CALCULATION` (no implementadas).

## Privacidad

Sólo rótulos del formulario, coordenadas y valores agregados. Ninguna fila
Medinet individual.
"""


def write_outputs(result: InventoryResult, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    files = [
        _write_candidates(result, output_dir),
        _write_validations(result, output_dir),
        _write_formula_evidence(result, output_dir),
        _write_consistency(result, output_dir),
        _write_duplicates(result, output_dir),
        _write_context_coverage(result, output_dir),
        *_write_layouts(result, output_dir),
    ]
    summary_path = output_dir / "summary.json"
    summary_path.write_text(
        json.dumps(_summary_dict(result), indent=2, ensure_ascii=False), encoding="utf-8"
    )
    readme = output_dir / "README.md"
    readme.write_text(_README, encoding="utf-8")
    return [*files, summary_path, readme]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="inventory_semantic_metrics.py",
        description=(
            "Inventario semántico preliminar de las métricas legacy: rótulos de "
            "fila/columna/sección + evidencia de fórmula. Sólo lectura."
        ),
    )
    parser.add_argument("workbook", help="Ruta al workbook legacy (.xlsx/.xlsm).")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"(def: {DEFAULT_OUTPUT})")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = build_inventory(args.workbook)
    except clo.ClosureError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    files = write_outputs(result, Path(args.output))
    summary = _summary_dict(result)

    print(f"Workbook: {result.source_name}")
    print(f"  métricas candidatas: {summary['total_metric_candidates']}")
    print(f"  por hoja: {summary['metrics_by_sheet']}")
    print(f"  por tipo: {summary['metrics_by_kind']}")
    print(f"  context_status: {summary['context_status']}")
    print(
        f"  jerarquía máx: filas {summary['max_row_hierarchy_depth']} / "
        f"columnas {summary['max_column_hierarchy_depth']} / "
        f"secciones {summary['max_section_hierarchy_depth']}"
    )
    print(
        f"  firmas semánticas duplicadas: {summary['duplicate_semantic_signatures']} "
        f"({summary['metrics_in_duplicate_groups']} métricas)"
    )
    print(
        f"  evidencia: {summary['metrics_with_age_bounds']} con cotas de edad, "
        f"{summary['metrics_with_text_criteria']} con criterios de texto"
    )
    print(
        f"  rótulo vs. fórmula: {summary['label_formula_consistent']} CONSISTENT / "
        f"{summary['label_formula_conflicts']} CONFLICT / "
        f"{summary['label_formula_no_comparable']} NO_COMPARABLE"
    )
    print(f"  validaciones: {summary['validation_total']} (contexto: {summary['validation_context_status']})")
    print(f"Artefactos en {args.output}:")
    for path in files:
        print(f"  - {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
