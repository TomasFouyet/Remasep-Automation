"""Validación / readiness del mapeo semántico (Sprint 3.3).

Parte del mapeo de Sprint 3.2 (`scripts/map_semantic_metrics.py`) y le añade una
capa técnica de **readiness**:

    AUTO_READY · REVIEW_REQUIRED · BLOCKED_CONFLICT · NOT_APPLICABLE

- **No cambia** las inferencias semánticas de Sprint 3.2.
- Reglas transparentes y versionadas en `config/semantic_validation_2026/`.
- Los 5 conflictos reales del workbook quedan `BLOCKED_CONFLICT` con
  `recommended_action = HUMAN_REVIEW` — **nunca** se propone corregirlos.
- `AUTO_READY` = "suficiente evidencia técnica según la policy". **No** es
  "validado MINSAL".

Uso:

    python scripts/validate_semantic_mapping.py \\
        "data/local/GENERACION DATOS REMASEP.xlsx"
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import Counter, defaultdict
from datetime import UTC, datetime
from pathlib import Path

import close_legacy_aggregations as clo
import map_semantic_metrics as mp

from remasep.services import semantic_mapping as sm
from remasep.services import semantic_validation as sv

DEFAULT_OUTPUT = "artifacts/semantic_mapping_validation"
_READINESS_ORDER = (sv.AUTO_READY, sv.REVIEW_REQUIRED, sv.BLOCKED_CONFLICT, sv.NOT_APPLICABLE)
_SEVERITY_ORDER = (sv.SEV_BLOCKING, sv.SEV_REVIEW, sv.SEV_INFO)


def build_validation(path):
    inventory, mapping_config, metrics = mp.build_mapping(path)
    policy = sv.load_readiness_policy()
    overrides = sv.load_overrides()
    result = sv.assess_all(metrics, policy, overrides)
    return inventory, mapping_config, metrics, policy, overrides, result


# ---------------------------------------------------------------------------
# Artefactos
# ---------------------------------------------------------------------------


def _csv(path: Path, header: list[str], rows) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def _all_issues(result: sv.ValidationResult) -> list[sv.SemanticReviewIssue]:
    return [issue for r in result.readiness for issue in r.issues]


def _write_readiness(result: sv.ValidationResult, out: Path) -> Path:
    path = out / "semantic_metric_readiness.csv"
    header = [
        "metric_id", "source", "form", "sheet", "cell", "kind", "mapping_status",
        "readiness_status", "review_reason_count", "review_reasons",
        "sex_status", "age_status", "aggregation_scope_status", "procedure_status",
    ]
    rows = (
        [
            r.metric_id, r.source, r.form, r.sheet, r.cell, r.kind, r.mapping_status,
            r.readiness_status, len(r.review_reasons), "|".join(r.review_reasons),
            r.sex_status, r.age_status, r.aggregation_scope_status, r.procedure_status,
        ]
        for r in result.readiness
    )
    _csv(path, header, rows)
    return path


def _write_review_queue(result: sv.ValidationResult, out: Path) -> Path:
    path = out / "review_queue.csv"
    header = [
        "issue_id", "metric_id", "form", "sheet", "cell", "dimension", "issue_type",
        "severity", "evidence_summary", "recommended_action",
    ]
    rows = (
        [
            i.issue_id, i.metric_id, i.form, i.sheet, i.cell, i.dimension, i.issue_type,
            i.severity, i.evidence_summary(), i.recommended_action,
        ]
        for i in sorted(_all_issues(result), key=lambda x: (x.severity != sv.SEV_BLOCKING, x.issue_id))
        if i.severity in (sv.SEV_BLOCKING, sv.SEV_REVIEW)
    )
    _csv(path, header, rows)
    return path


def _write_conflicts(result: sv.ValidationResult, out: Path) -> Path:
    path = out / "conflicts.csv"
    header = [
        "issue_id", "metric_id", "form", "cell", "dimension", "issue_type",
        "label_evidence", "formula_evidence", "recommended_action", "status",
    ]
    rows = (
        [
            i.issue_id, i.metric_id, i.form, i.cell, i.dimension, i.issue_type,
            i.label_evidence, i.formula_evidence, i.recommended_action, i.status,
        ]
        for i in sorted(_all_issues(result), key=lambda x: x.issue_id)
        if i.severity == sv.SEV_BLOCKING
    )
    _csv(path, header, rows)
    return path


def _write_review_reason_summary(result: sv.ValidationResult, out: Path) -> Path:
    path = out / "review_reason_summary.csv"
    counter: Counter[tuple[str, str, str]] = Counter()
    for issue in _all_issues(result):
        counter[(issue.issue_type, issue.dimension, issue.severity)] += 1
    header = ["issue_type", "dimension", "severity", "metric_count"]
    rows = [
        [issue_type, dimension, severity, count]
        for (issue_type, dimension, severity), count in sorted(
            counter.items(), key=lambda kv: (-kv[1], kv[0])
        )
    ]
    _csv(path, header, rows)
    return path


def _write_review_clusters(result: sv.ValidationResult, out: Path) -> Path:
    path = out / "review_clusters.csv"
    header = [
        "cluster_id", "form", "dimension", "issue_type", "severity", "metric_count",
        "example_metric_ids",
    ]
    rows = (
        [c.cluster_id, c.form, c.dimension, c.issue_type, c.severity, c.metric_count,
         "|".join(c.example_metric_ids)]
        for c in sv.build_review_clusters(result)
    )
    _csv(path, header, rows)
    return path


def _write_readiness_coverage(result: sv.ValidationResult, out: Path) -> Path:
    path = out / "readiness_coverage.csv"
    grouped: dict[tuple[str, str], list[sv.SemanticReadiness]] = defaultdict(list)
    for r in result.readiness:
        grouped[(r.form, r.kind)].append(r)
    header = ["scope", "metric_count", *_READINESS_ORDER]

    def _row(label: str, items: list[sv.SemanticReadiness]) -> list:
        counts = Counter(x.readiness_status for x in items)
        return [label, len(items), *(counts.get(s, 0) for s in _READINESS_ORDER)]

    rows = [_row(f"{form} / {kind}", g) for (form, kind), g in sorted(grouped.items())]
    rows.append(_row("TOTAL", result.readiness))
    _csv(path, header, rows)
    return path


def _manual_review_sample(result: sv.ValidationResult) -> list[sv.SemanticReadiness]:
    ordered = sorted(result.readiness, key=lambda r: r.metric_id)
    buckets = [
        lambda r: r.readiness_status == sv.AUTO_READY,
        lambda r: r.readiness_status == sv.REVIEW_REQUIRED,
        lambda r: r.readiness_status == sv.NOT_APPLICABLE,
        lambda r: r.form == "REMASEP_01",
        lambda r: r.form == "REMASEP_OD",
        lambda r: r.form == "B2_ANEXO",
        lambda r: r.kind == "BASE_AGGREGATION",
        lambda r: r.kind == "DERIVED_AGGREGATION",
        lambda r: r.kind == "DOWNSTREAM_TOTAL",
        lambda r: "SEX_LABEL_ONLY" in r.review_reasons,
        lambda r: "AGE_LABEL_ONLY" in r.review_reasons,
        lambda r: "AGGREGATION_SCOPE_UNRESOLVED" in r.review_reasons,
        lambda r: r.readiness_status == sv.AUTO_READY and r.mapping_status == "PARTIAL",
    ]
    picked: dict[str, sv.SemanticReadiness] = {}
    for predicate in buckets:
        for r in [x for x in ordered if predicate(x)][:4]:
            picked[r.metric_id] = r
    for r in ordered:  # todos los BLOCKED_CONFLICT
        if r.readiness_status == sv.BLOCKED_CONFLICT:
            picked[r.metric_id] = r
    return sorted(picked.values(), key=lambda r: r.metric_id)


def _write_manual_review_sample(
    result: sv.ValidationResult, metrics_by_id: dict[str, sm.SemanticMetric], out: Path
) -> Path:
    path = out / "readiness_manual_review_sample.csv"
    header = [
        "metric_id", "form", "kind", "cell", "row_path", "column_path",
        "mapping_status", "readiness_status", "review_reasons",
    ]
    rows = []
    for r in _manual_review_sample(result):
        m = metrics_by_id[r.metric_id]
        rows.append(
            [
                r.metric_id, r.form, r.kind, r.cell, " :: ".join(m.row_path_raw),
                " :: ".join(m.column_path_raw), r.mapping_status, r.readiness_status,
                "|".join(r.review_reasons),
            ]
        )
    _csv(path, header, rows)
    return path


_README = """# Semantic mapping validation & readiness (Sprint 3.3)

Salida de `scripts/validate_semantic_mapping.py`. Capa técnica sobre el
`SemanticMetric` de Sprint 3.2: responde, por métrica, **si puede usarse
automáticamente, si necesita revisión o si debe bloquearse**.

> `AUTO_READY` significa **únicamente** "suficiente evidencia técnica según la
> policy actual" (`config/semantic_validation_2026/policy.yaml`). **No** significa
> "validado MINSAL" ni "aprobado por el cliente".

## `readiness_status` (≠ `mapping_status`)

| Estado | Cuándo |
| --- | --- |
| `BLOCKED_CONFLICT` | alguna dimensión tiene `CONFLICT` (rótulo ↔ fórmula) — **no se resuelve** |
| `NOT_APPLICABLE` | la métrica es un roll-up del workbook (`aggregation_scope` ∈ `TOTAL`/`SUBTOTAL`): no es un punto de ingreso desde Medinet |
| `REVIEW_REQUIRED` | una **dimensión requerida** para esa hoja/alcance no está resuelta (`LABEL_ONLY` / `FORMULA_ONLY` / `UNRESOLVED`) |
| `AUTO_READY` | sin conflicto y con todas las dimensiones requeridas en `CONFIRMED` o `NOT_APPLICABLE` |

Precedencia: `BLOCKED_CONFLICT` > `NOT_APPLICABLE` > `REVIEW_REQUIRED` >
`AUTO_READY`.

## Severidad de issue

| Severidad | Origen |
| --- | --- |
| `BLOCKING` | un `CONFLICT` de dimensión |
| `REVIEW` | una **dimensión requerida** no resuelta |
| `INFO` | una dimensión **no requerida** no resuelta (p.ej. código de procedimiento en `REMASEP_OD`) |

## Archivos

- `semantic_metric_readiness.csv` — una fila por métrica: `readiness_status`,
  `review_reasons`, estados de dimensión.
- `review_queue.csv` — una fila por issue accionable (`BLOCKING` + `REVIEW`),
  con `recommended_action = HUMAN_REVIEW`.
- `conflicts.csv` — los issues `BLOCKING` (los 5 conflictos reales del workbook).
  `label_evidence` / `formula_evidence` se registran; **no** se propone un valor.
- `review_reason_summary.csv` — conteo por `issue_type` × `dimension` × severidad.
- `review_clusters.csv` — problemas equivalentes agrupados (`form` × `dimension`
  × `issue_type`), con ejemplos.
- `readiness_coverage.csv` — `readiness_status` por `form` × `kind`.
- `readiness_manual_review_sample.csv` — muestra determinista y estratificada.
- `summary.json` — totales; `validation_status: PRELIMINARY_NOT_VALIDATED`.

## Overrides (arquitectura, sin resolución en 3.3)

`config/semantic_validation_2026/overrides.yaml` está **vacío**. En el futuro una
decisión humana podrá resolver una excepción de forma auditable, identificada
por `metric_signature` (independiente de coordenadas; `cell` sólo como
trazabilidad). En 3.3 sólo se carga y se registra la coincidencia por firma.

## Privacidad

Sólo `metric_id`, coordenadas, rótulos del formulario y estados agregados.
Ninguna fila Medinet individual.
"""


def _summary(inventory, result: sv.ValidationResult, policy: sv.ReadinessPolicy,
             overrides: dict) -> dict:
    readiness = result.readiness
    by_status = Counter(r.readiness_status for r in readiness)
    issues = _all_issues(result)
    by_sev = Counter(i.severity for i in issues)
    partial = [r for r in readiness if r.mapping_status == sm.MAPPING_PARTIAL]
    partial_block = [r for r in partial if r.readiness_status == sv.REVIEW_REQUIRED]
    partial_ok = [
        r for r in partial if r.readiness_status in (sv.AUTO_READY, sv.NOT_APPLICABLE)
    ]
    top_reasons = Counter(
        i.issue_type for i in issues if i.severity in (sv.SEV_BLOCKING, sv.SEV_REVIEW)
    )
    sample = _manual_review_sample(result)
    return {
        "source": inventory.source_name,
        "source_sha256": inventory.source_sha256,
        "generated_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "validation_policy_version": policy.version,
        "validation_status": "PRELIMINARY_NOT_VALIDATED",
        "total_metrics": len(readiness),
        "auto_ready": by_status.get(sv.AUTO_READY, 0),
        "review_required": by_status.get(sv.REVIEW_REQUIRED, 0),
        "blocked_conflict": by_status.get(sv.BLOCKED_CONFLICT, 0),
        "not_applicable": by_status.get(sv.NOT_APPLICABLE, 0),
        "readiness_by_form": {
            form: dict(sorted(Counter(
                r.readiness_status for r in readiness if r.form == form
            ).items()))
            for form in sorted({r.form for r in readiness})
        },
        "readiness_by_kind": {
            kind: dict(sorted(Counter(
                r.readiness_status for r in readiness if r.kind == kind
            ).items()))
            for kind in sorted({r.kind for r in readiness})
        },
        "review_issue_count": len(issues),
        "blocking_issue_count": by_sev.get(sv.SEV_BLOCKING, 0),
        "review_severity_issue_count": by_sev.get(sv.SEV_REVIEW, 0),
        "info_issue_count": by_sev.get(sv.SEV_INFO, 0),
        "review_cluster_count": len(sv.build_review_clusters(result)),
        "top_review_reasons": dict(top_reasons.most_common(10)),
        "mapping_confirmed": sum(1 for r in readiness if r.mapping_status == sm.MAPPING_CONFIRMED),
        "mapping_partial": len(partial),
        "mapping_conflict": sum(1 for r in readiness if r.mapping_status == sm.MAPPING_CONFLICT),
        "partial_that_block_readiness": len(partial_block),
        "partial_acceptable": len(partial_ok),
        "metrics_with_no_review_reasons": sum(1 for r in readiness if not r.review_reasons),
        "override_count": len(overrides),
        "readiness_manual_review_sample_size": len(sample),
        "readiness_manual_review_sample_by_status": dict(
            sorted(Counter(r.readiness_status for r in sample).items())
        ),
    }


def write_outputs(inventory, result: sv.ValidationResult, policy: sv.ReadinessPolicy,
                  overrides: dict, metrics: list[sm.SemanticMetric], output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    by_id = {m.metric_id: m for m in metrics}
    files = [
        _write_readiness(result, output_dir),
        _write_review_queue(result, output_dir),
        _write_conflicts(result, output_dir),
        _write_review_reason_summary(result, output_dir),
        _write_review_clusters(result, output_dir),
        _write_readiness_coverage(result, output_dir),
        _write_manual_review_sample(result, by_id, output_dir),
    ]
    (output_dir / "summary.json").write_text(
        json.dumps(_summary(inventory, result, policy, overrides), indent=2, ensure_ascii=False),
        encoding="utf-8",
    )
    (output_dir / "README.md").write_text(_README, encoding="utf-8")
    return [*files, output_dir / "summary.json", output_dir / "README.md"]


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="validate_semantic_mapping.py",
        description=(
            "Readiness técnica del mapeo semántico: AUTO_READY / REVIEW_REQUIRED / "
            "BLOCKED_CONFLICT / NOT_APPLICABLE. Sólo lectura, no cambia el mapping."
        ),
    )
    parser.add_argument("workbook", help="Ruta al workbook legacy (.xlsx/.xlsm).")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"(def: {DEFAULT_OUTPUT})")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        inventory, _mapcfg, metrics, policy, overrides, result = build_validation(args.workbook)
    except (clo.ClosureError, sm.SemanticMappingError, sv.SemanticValidationError) as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    files = write_outputs(inventory, result, policy, overrides, metrics, Path(args.output))
    summary = _summary(inventory, result, policy, overrides)

    print(f"Workbook: {summary['source']}  (policy {summary['validation_policy_version']})")
    print(
        f"  readiness: {summary['auto_ready']} AUTO_READY / "
        f"{summary['review_required']} REVIEW_REQUIRED / "
        f"{summary['blocked_conflict']} BLOCKED_CONFLICT / "
        f"{summary['not_applicable']} NOT_APPLICABLE  (de {summary['total_metrics']})"
    )
    print(f"  por hoja: {summary['readiness_by_form']}")
    print(f"  por tipo: {summary['readiness_by_kind']}")
    print(
        f"  issues: {summary['blocking_issue_count']} BLOCKING / "
        f"{summary['review_severity_issue_count']} REVIEW / {summary['info_issue_count']} INFO  "
        f"({summary['review_cluster_count']} clusters)"
    )
    print(f"  top review reasons: {summary['top_review_reasons']}")
    print(
        f"  PARTIAL: {summary['partial_that_block_readiness']} bloquean readiness / "
        f"{summary['partial_acceptable']} aceptables (N/A o no-requerida)"
    )
    print(f"  mapping: {summary['mapping_confirmed']} CONFIRMED / {summary['mapping_partial']} "
          f"PARTIAL / {summary['mapping_conflict']} CONFLICT")
    print(f"  overrides cargados: {summary['override_count']}")
    print(
        f"  muestra: {summary['readiness_manual_review_sample_size']} filas "
        f"{summary['readiness_manual_review_sample_by_status']}"
    )
    print(f"Artefactos en {args.output}:")
    for path in files:
        print(f"  - {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
