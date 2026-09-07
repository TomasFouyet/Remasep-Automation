"""Sprint 3.4 — Delta del REMASEP final + *provenance* de fuente.

Compara la versión **incompleta** del REMASEP julio 2026 (sin cirugías) con la
versión **terminada según el cliente** y responde, por celda cambiada:

- ¿es un **input directo** o la **propagación** de una fórmula?
- ¿de qué **fuente/proceso** proviene? (``EGRESOS`` · ``RESOURCE_CALCULATION`` ·
  ``SURGICAL_TABLE`` · ``CONTROL_METADATA`` · ``MEDINET`` · ``UNKNOWN_PENDING`` ·
  ``MIXED_DERIVED``)

Sólo lectura: no modifica ni guarda ningún workbook, no usa COM/LibreOffice/
macros, no recalcula. Verifica el SHA256 de ambos archivos antes y después.

Uso:

    python scripts/analyze_final_remasep_delta.py \\
        "data/local/2026-7 REMASEP_V1.4.xlsm" \\
        "data/local/REMASEP_V1.4 Julio 2026.xlsm"
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from collections import Counter, defaultdict
from datetime import UTC, datetime
from pathlib import Path

from remasep.services import source_provenance as sp
from remasep.services import workbook_delta as wd

DEFAULT_OUTPUT = "artifacts/final_remasep_delta_provenance"
DEFAULT_BEFORE = "data/local/2026-7 REMASEP_V1.4.xlsm"
DEFAULT_AFTER = "data/local/REMASEP_V1.4 Julio 2026.xlsm"

REFERENCE_BEFORE_STATUS = "INCOMPLETE_PRE_SURGERY"
REFERENCE_AFTER_STATUS = "FINAL_PER_CLIENT"
APPROVAL_STATUS = "UNCONFIRMED"
CANDIDATE_GOLDEN_STATUS = "CANDIDATE_PENDING_SOURCE_COMPLETENESS"


# ---------------------------------------------------------------------------
# Provenance
# ---------------------------------------------------------------------------


def _control_context(pair: wd.WorkbookPair, delta: wd.DeltaCell) -> tuple[str, str]:
    """(hoja_sin_datos, causal) para un cambio directo de la hoja CONTROL."""
    ws = pair.after_value["CONTROL"]
    reason = ws[delta.cell].value
    sheet_without_data = ws.cell(row=delta.row, column=4).value  # columna D
    return (
        str(sheet_without_data or "").strip(),
        str(reason or "").strip(),
    )


def _delta_context(pair: wd.WorkbookPair, delta: wd.DeltaCell) -> sp.DeltaContext:
    procedure_code = ""
    control_reason = ""
    control_sheet = ""
    if delta.sheet in {"B2 ANEXO", "B2_ANEXO"}:
        raw = pair.after_value[delta.sheet].cell(row=delta.row, column=1).value
        if isinstance(raw, (int, str)) and str(raw).strip().isdigit():
            procedure_code = str(raw).strip()
    if delta.sheet == "CONTROL":
        control_sheet, control_reason = _control_context(pair, delta)
    return sp.DeltaContext(
        sheet=delta.sheet,
        section_path=delta.section_path,
        row_path=delta.row_path,
        column_path=delta.column_path,
        procedure_code=procedure_code,
        control_reason=control_reason,
        control_sheet_without_data=control_sheet,
    )


def assign_provenance(
    pair: wd.WorkbookPair,
    deltas: list[wd.DeltaCell],
    edges: list[wd.DeltaEdge],
) -> dict[tuple[str, str], sp.SourceProvenance]:
    """Clasifica los directos y propaga a los resultados de fórmula (punto fijo)."""
    prov: dict[tuple[str, str], sp.SourceProvenance] = {}
    incoming: dict[tuple[str, str], list[tuple[str, str]]] = defaultdict(list)
    for edge in edges:
        incoming[(edge.target_sheet, edge.target_cell)].append(
            (edge.source_sheet, edge.source_cell)
        )

    for delta in deltas:
        if delta.change_kind == wd.DIRECT_INPUT_CHANGE:
            prov[delta.key] = sp.classify_direct(_delta_context(pair, delta))
        elif delta.change_kind == wd.FORMULA_EXPRESSION_CHANGE:
            # la expresión se reescribió: nunca input ni propagación de resultado.
            prov[delta.key] = sp.unknown_expression_change(
                delta.before_formula, delta.formula
            )

    formula_keys = [d.key for d in deltas if d.change_kind == wd.FORMULA_RESULT_CHANGE]
    for _ in range(len(formula_keys) + 1):
        progressed = False
        for key in formula_keys:
            if key in prov:
                continue
            sources = incoming.get(key, [])
            if sources and all(s in prov for s in sources):
                prov[key] = sp.combine_derived([prov[s] for s in sources])
                progressed = True
        if not progressed:
            break

    # cualquier fórmula sin resolver (ciclo o upstream no cambiado): mejor esfuerzo
    for key in formula_keys:
        if key not in prov:
            resolved = [prov[s] for s in incoming.get(key, []) if s in prov]
            prov[key] = sp.combine_derived(resolved)
    return prov


# ---------------------------------------------------------------------------
# 14. Consistency checks
# ---------------------------------------------------------------------------


def _norm(value: object) -> str:
    return sp._norm(value)


def _b1_section_a_axes(ws) -> tuple[list[int], list[int], int | None]:
    """(columnas de edad, columnas de sexo, fila de etiquetas) en REMASEP B1 Sec. A."""
    header_row = age_start = sex_start = None
    for r in range(1, 25):
        for c in range(1, 40):
            txt = _norm(ws.cell(row=r, column=c).value)
            if "GRUPO DE EDAD" in txt:
                header_row, age_start = r, c
            if "POR SEXO" in txt and header_row == r:
                sex_start = c
    if header_row is None or age_start is None or sex_start is None:
        return [], [], None
    label_row = header_row + 1
    age_cols, sex_cols = [], []
    for c in range(age_start, sex_start):
        label = _norm(ws.cell(row=label_row, column=c).value)
        if label and label != "TOTAL":
            age_cols.append(c)
    for c in range(sex_start, sex_start + 12):
        label = _norm(ws.cell(row=label_row, column=c).value)
        if label and label != "TOTAL":
            sex_cols.append(c)
    return age_cols, sex_cols, label_row


def _num(value: object) -> float:
    return float(value) if isinstance(value, (int, float)) else 0.0


def build_consistency_checks(
    pair: wd.WorkbookPair, deltas: list[wd.DeltaCell]
) -> list[dict]:
    checks: list[dict] = []
    if "REMASEP B1" not in pair.after_value.sheetnames:
        return checks
    ws = pair.after_value["REMASEP B1"]

    age_cols, sex_cols, _label_row = _b1_section_a_axes(ws)
    changed_rows = sorted({
        d.row for d in deltas
        if d.sheet == "REMASEP B1" and d.change_kind == wd.DIRECT_INPUT_CHANGE and d.row < 25
    })
    age_sum = sum(_num(ws.cell(row=r, column=c).value) for r in changed_rows for c in age_cols)
    sex_sum = sum(_num(ws.cell(row=r, column=c).value) for r in changed_rows for c in sex_cols)
    checks.append({
        "check_id": "B1_SECCION_A_AGE_TOTAL_VS_SEX_TOTAL",
        "left_metric": "REMASEP B1 Sección A · Σ columnas por grupo de edad (filas con input)",
        "left_value": int(age_sum),
        "right_metric": "REMASEP B1 Sección A · Σ columnas por sexo (filas con input)",
        "right_value": int(sex_sum),
        "difference": int(age_sum - sex_sum),
        "status": "OK" if age_sum == sex_sum else "OPEN_FUNCTIONAL_QUESTION",
        "note": (
            "El desglose por edad y por sexo de la grilla de cirugías (EGRESOS) "
            "debe cuadrar entre sí." if age_sum == sex_sum else
            "Descuadre entre el total por edad y por sexo — revisar con el cliente."
        ),
    })

    # total de intervenciones codificadas (Sección B) = fila 'TOTAL DE INTERVENCIONES...'
    coded_total = None
    for r in range(1, ws.max_row + 1):
        if "TOTAL DE INTERVENCIONES QUIRURGICAS" in _norm(ws.cell(row=r, column=1).value):
            coded_total = _num(ws.cell(row=r, column=3).value)
            break
    if coded_total is not None:
        checks.append({
            "check_id": "B1_GRID_TOTAL_VS_CODED_INTERVENTIONS",
            "left_metric": "REMASEP B1 Sección A · Σ intervenciones por grupo de edad",
            "left_value": int(age_sum),
            "right_metric": "REMASEP B1 Sección B · TOTAL de intervenciones quirúrgicas codificadas",
            "right_value": int(coded_total),
            "difference": int(age_sum - coded_total),
            "status": "OK" if age_sum == coded_total else "OPEN_FUNCTIONAL_QUESTION",
            "note": (
                "32 vs 31: la grilla por edad/sexo tiene una intervención más que la "
                "suma de códigos de B2 ANEXO. No se asume error; pendiente de "
                "reconciliación funcional con el cliente (¿intervención sin código? "
                "¿múltiples procedimientos en un egreso?)."
            ) if age_sum != coded_total else "Cuadran.",
        })

    # cadena de propagación B2 -> B1 Sección B
    b1_secB_changed = sum(
        _num(d.after_value) for d in deltas
        if d.sheet == "REMASEP B1" and d.change_kind == wd.FORMULA_RESULT_CHANGE
        and 26 < d.row < 46
    )
    if coded_total is not None:
        checks.append({
            "check_id": "B1_SECCION_B_PULLS_VS_TOTAL",
            "left_metric": "REMASEP B1 Sección B · Σ filas de categoría que cambiaron (pull de B2 ANEXO)",
            "left_value": int(b1_secB_changed),
            "right_metric": "REMASEP B1 Sección B · TOTAL de intervenciones quirúrgicas codificadas",
            "right_value": int(coded_total),
            "difference": int(b1_secB_changed - coded_total),
            "status": "OK" if b1_secB_changed == coded_total else "OPEN_FUNCTIONAL_QUESTION",
            "note": "Valida la propagación EGRESOS → B2 subtotales → B1 Sección B.",
        })
    return checks


# ---------------------------------------------------------------------------
# 17/18. CONTROL
# ---------------------------------------------------------------------------


def build_control_summary(pair: wd.WorkbookPair) -> dict:
    if "CONTROL" not in pair.after_value.sheetnames:
        return {"total_errores": 0, "control_status": "PASS_INTERNAL_VALIDATION",
                "approval_status": APPROVAL_STATUS, "sheets": [], "causales": []}
    ws = pair.after_value["CONTROL"]
    rows = []
    total_errors = 0
    for r in range(1, ws.max_row + 1):
        name = ws.cell(row=r, column=4).value  # D
        status = ws.cell(row=r, column=6).value  # F
        if not isinstance(name, str) or not name.strip():
            continue
        if not isinstance(status, str) or status.strip() not in {"OK", "SIN DATOS"}:
            continue
        errors = ws.cell(row=r, column=5).value  # E
        reason = ws.cell(row=r, column=7).value  # G
        n_err = int(errors) if isinstance(errors, (int, float)) else 0
        total_errors += n_err
        rows.append({
            "sheet": name.strip(),
            "n_errores": n_err,
            "estado": status.strip(),
            "causal": str(reason).strip() if isinstance(reason, str) and reason.strip() else "",
        })
    return {
        "total_errores": total_errors,
        "control_status": "PASS_INTERNAL_VALIDATION" if total_errors == 0 else "HAS_ERRORS",
        "approval_status": APPROVAL_STATUS,
        "sheets": rows,
        "causales": [r for r in rows if r["causal"]],
    }


# ---------------------------------------------------------------------------
# Artefactos
# ---------------------------------------------------------------------------


def _csv(path: Path, header: list[str], rows) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def _ctx(parts: tuple[str, ...]) -> str:
    return " :: ".join(parts)


def _direct_or_derived(change_kind: str) -> str:
    if change_kind == wd.DIRECT_INPUT_CHANGE:
        return "DIRECT"
    if change_kind == wd.FORMULA_EXPRESSION_CHANGE:
        return "EXPRESSION_CHANGE"
    return "DERIVED"


def write_outputs(
    *,
    output_dir: Path,
    pair: wd.WorkbookPair,
    structure: wd.StructureComparison,
    formula_changes: list[wd.FormulaChange],
    deltas: list[wd.DeltaCell],
    edges: list[wd.DeltaEdge],
    prov: dict[tuple[str, str], sp.SourceProvenance],
    checks: list[dict],
    control: dict,
    summary: dict,
) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    incoming: dict[tuple[str, str], list[str]] = defaultdict(list)
    for edge in edges:
        incoming[(edge.target_sheet, edge.target_cell)].append(
            f"{edge.source_sheet}!{edge.source_cell}"
        )

    files: list[Path] = []

    p = output_dir / "workbook_structure_comparison.json"
    p.write_text(json.dumps(structure.to_dict(), indent=2, ensure_ascii=False), encoding="utf-8")
    files.append(p)

    p = output_dir / "formula_expression_changes.csv"
    _csv(p, ["sheet", "cell", "before_formula", "after_formula", "change_type"],
         ([f.sheet, f.cell, f.before_formula, f.after_formula, f.change_type]
          for f in formula_changes))
    files.append(p)

    p = output_dir / "delta_cells.csv"
    _csv(
        p,
        ["sheet", "cell", "change_kind", "before_value", "after_value",
         "before_formula", "formula", "formula_change_type",
         "section_path", "row_path", "column_path", "source", "provenance_status",
         "provenance_evidence"],
        (
            [
                d.sheet, d.cell, d.change_kind, d.before_value, d.after_value,
                d.before_formula, d.formula, d.formula_change_type,
                _ctx(d.section_path), _ctx(d.row_path), _ctx(d.column_path),
                prov[d.key].source, prov[d.key].status, prov[d.key].evidence_summary(),
            ]
            for d in deltas
        ),
    )
    files.append(p)

    p = output_dir / "direct_input_changes.csv"
    _csv(
        p,
        ["sheet", "cell", "before_value", "after_value", "section_path", "row_path",
         "column_path", "source", "provenance_status", "provenance_evidence"],
        (
            [d.sheet, d.cell, d.before_value, d.after_value, _ctx(d.section_path),
             _ctx(d.row_path), _ctx(d.column_path), prov[d.key].source,
             prov[d.key].status, prov[d.key].evidence_summary()]
            for d in deltas if d.change_kind == wd.DIRECT_INPUT_CHANGE
        ),
    )
    files.append(p)

    p = output_dir / "formula_result_changes.csv"
    _csv(
        p,
        ["sheet", "cell", "before_value", "after_value", "formula",
         "upstream_changed_cells", "source", "provenance_status", "provenance_evidence"],
        (
            [d.sheet, d.cell, d.before_value, d.after_value, d.formula,
             "|".join(incoming.get(d.key, [])), prov[d.key].source,
             prov[d.key].status, prov[d.key].evidence_summary()]
            for d in deltas if d.change_kind == wd.FORMULA_RESULT_CHANGE
        ),
    )
    files.append(p)

    p = output_dir / "delta_dependency_edges.csv"
    _csv(p, ["source_sheet", "source_cell", "target_sheet", "target_cell", "dependency_type"],
         ([e.source_sheet, e.source_cell, e.target_sheet, e.target_cell, e.dependency_type]
          for e in edges))
    files.append(p)

    p = output_dir / "provenance_map.csv"
    _csv(
        p,
        ["sheet", "cell", "semantic_context", "change_origin", "source", "status",
         "direct_or_derived", "upstream_changed_cells", "evidence_summary",
         "needs_human_review"],
        (
            [
                d.sheet, d.cell,
                _ctx(d.section_path + d.row_path + d.column_path),
                prov[d.key].change_origin, prov[d.key].source, prov[d.key].status,
                _direct_or_derived(d.change_kind),
                "|".join(incoming.get(d.key, [])),
                prov[d.key].evidence_summary(),
                "yes" if prov[d.key].needs_human_review else "no",
            ]
            for d in deltas
        ),
    )
    files.append(p)

    p = output_dir / "source_summary.csv"
    direct_ct: Counter[str] = Counter()
    derived_ct: Counter[str] = Counter()
    expr_ct: Counter[str] = Counter()
    review_ct: Counter[str] = Counter()
    for d in deltas:
        pr = prov[d.key]
        if d.change_kind == wd.DIRECT_INPUT_CHANGE:
            direct_ct[pr.source] += 1
        elif d.change_kind == wd.FORMULA_EXPRESSION_CHANGE:
            expr_ct[pr.source] += 1
        else:
            derived_ct[pr.source] += 1
        if pr.needs_human_review:
            review_ct[pr.source] += 1
    seen_sources = sorted(set(direct_ct) | set(derived_ct) | set(expr_ct))
    _csv(
        p,
        ["source", "direct_change_count", "derived_change_count",
         "expression_change_count", "total_change_count", "human_review_count"],
        (
            [s, direct_ct.get(s, 0), derived_ct.get(s, 0), expr_ct.get(s, 0),
             direct_ct.get(s, 0) + derived_ct.get(s, 0) + expr_ct.get(s, 0),
             review_ct.get(s, 0)]
            for s in seen_sources
        ),
    )
    files.append(p)

    p = output_dir / "source_consistency_checks.csv"
    _csv(
        p,
        ["check_id", "left_metric", "left_value", "right_metric", "right_value",
         "difference", "status", "note"],
        ([c["check_id"], c["left_metric"], c["left_value"], c["right_metric"],
          c["right_value"], c["difference"], c["status"], c["note"]] for c in checks),
    )
    files.append(p)

    p = output_dir / "control_summary.csv"
    _csv(p, ["sheet", "n_errores", "estado", "causal"],
         ([r["sheet"], r["n_errores"], r["estado"], r["causal"]] for r in control["sheets"]))
    files.append(p)

    p = output_dir / "summary.json"
    p.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    files.append(p)

    p = output_dir / "README.md"
    p.write_text(_README, encoding="utf-8")
    files.append(p)
    return files


_README = """# Final REMASEP delta & source provenance (Sprint 3.4)

Salida de `scripts/analyze_final_remasep_delta.py`.

Compara dos versiones del REMASEP oficial de julio 2026:

| | archivo | estado |
| --- | --- | --- |
| **before** | `2026-7 REMASEP_V1.4.xlsm` | `INCOMPLETE_PRE_SURGERY` (sin cirugías, no final) |
| **after** | `REMASEP_V1.4 Julio 2026.xlsm` | `FINAL_PER_CLIENT` · `approval_status = UNCONFIRMED` |

Sólo lectura: no modifica ni guarda ningún workbook, no usa COM/LibreOffice/
macros, no recalcula. Verifica el SHA256 de ambos archivos antes y después.

## `change_kind`

Se decide con la fórmula **antes y después** (no sólo la de la versión final):

| valor | criterio |
| --- | --- |
| `DIRECT_INPUT_CHANGE` | sin fórmula ni antes ni después: es un dato ingresado |
| `FORMULA_RESULT_CHANGE` | **misma** expresión de fórmula; cambió sólo el valor cacheado |
| `FORMULA_EXPRESSION_CHANGE` | la **expresión** de la fórmula cambió (añadida / eliminada / modificada); ver `formula_change_type` |

Una *formula cache difference* **no** es un input manual. Un cambio de expresión
**no** se trata ni como input ni como propagación: su `source` queda
`UNKNOWN_PENDING` / `needs_human_review` salvo evidencia explícita.

## `source` (provenance)

| source | de dónde | cómo se afirma |
| --- | --- | --- |
| `EGRESOS` | edad/sexo de cirugía → B1; códigos Qx → B2 ANEXO | `CLIENT_CONFIRMED` |
| `RESOURCE_CALCULATION` | capacidad instalada (días hábiles × 7–8 h) — REMASEP 01 Sección D | `STRUCTURALLY_INFERRED` |
| `SURGICAL_TABLE` | horas ocupadas / programadas de tabla quirúrgica — REMASEP 01 Sección D | `STRUCTURALLY_INFERRED` |
| `CONTROL_METADATA` | causales de hojas sin datos en `CONTROL` | `STRUCTURALLY_INFERRED` |
| `MEDINET` | subgrafo ambulatorio (Sprint 2/3) | — |
| `UNKNOWN_PENDING` | no asignable con la evidencia actual (p.ej. Sección E: causas de suspensión) | `UNKNOWN` |
| `MIXED_DERIVED` | fórmula alimentada por varias fuentes distintas | `DERIVED_FROM_DEPENDENCIES` |

**No se infiere más de lo que la evidencia permite.** Sólo se marca
`CLIENT_CONFIRMED` donde el cliente lo confirmó explícitamente.

## Archivos

- `workbook_structure_comparison.json` — hojas, orden, dimensiones, recuento de
  fórmulas, merges, VBA, SHA256; `significant_changes`.
- `formula_expression_changes.csv` — cambios de **expresión** de fórmula (0
  esperados; descubierto, no *hardcodeado*).
- `delta_cells.csv` — todas las celdas cambiadas, con contexto semántico y
  provenance.
- `direct_input_changes.csv` / `formula_result_changes.csv` — vistas filtradas.
- `delta_dependency_edges.csv` — dependencia estructural entre celdas cambiadas
  (referencias A1 de la fórmula *después*).
- `provenance_map.csv` — una fila por delta: origen, fuente, estado,
  directo/derivado, upstream, `needs_human_review`.
- `source_summary.csv` — recuento por fuente (directo / derivado / total /
  revisión humana).
- `source_consistency_checks.csv` — chequeos cruzados; el 32-vs-31 queda como
  `OPEN_FUNCTIONAL_QUESTION`, nunca `FAIL`.
- `control_summary.csv` — estado por hoja de `CONTROL` y causales.
- `summary.json` — totales, `candidate_golden_status`.

## Golden candidate

`candidate_golden_status = CANDIDATE_PENDING_SOURCE_COMPLETENESS`: falta el
Medinet original de julio, los egresos reales de julio, la tabla quirúrgica y la
confirmación de envío/aprobación a MINSAL. `CONTROL = 0` significa
`PASS_INTERNAL_VALIDATION`, **no** "MINSAL approved".

## Privacidad

Sólo coordenadas, etiquetas del formulario, fórmulas y valores agregados. Sin
datos individuales.
"""


# ---------------------------------------------------------------------------
# Orquestación
# ---------------------------------------------------------------------------


def _egresos_findings(pair: wd.WorkbookPair, deltas: list[wd.DeltaCell],
                      prov: dict) -> dict:
    b2_codes = []
    for d in deltas:
        if d.sheet == "B2 ANEXO" and d.change_kind == wd.DIRECT_INPUT_CHANGE:
            ws = pair.after_value["B2 ANEXO"]
            b2_codes.append({
                "cell": d.cell,
                "procedure_code": str(ws.cell(row=d.row, column=1).value or "").strip(),
                "procedure_label": str(ws.cell(row=d.row, column=2).value or "").strip(),
                "count": d.after_value,
            })
    b1_direct = [
        {"cell": d.cell, "column_path": _ctx(d.column_path), "after_value": d.after_value}
        for d in deltas
        if d.sheet == "REMASEP B1" and d.change_kind == wd.DIRECT_INPUT_CHANGE
    ]
    return {"b2_procedure_inputs": b2_codes, "b1_direct_inputs": b1_direct}


def build_summary(
    *,
    pair: wd.WorkbookPair,
    structure: wd.StructureComparison,
    formula_changes: list[wd.FormulaChange],
    deltas: list[wd.DeltaCell],
    prov: dict[tuple[str, str], sp.SourceProvenance],
    checks: list[dict],
    control: dict,
) -> dict:
    by_sheet = Counter(d.sheet for d in deltas)
    by_kind = Counter(d.change_kind for d in deltas)
    by_source_direct = Counter(
        prov[d.key].source for d in deltas if d.change_kind == wd.DIRECT_INPUT_CHANGE
    )
    by_source_derived = Counter(
        prov[d.key].source for d in deltas if d.change_kind == wd.FORMULA_RESULT_CHANGE
    )
    by_source_expression = Counter(
        prov[d.key].source for d in deltas if d.change_kind == wd.FORMULA_EXPRESSION_CHANGE
    )
    needs_review = sum(1 for d in deltas if prov[d.key].needs_human_review)
    return {
        "generated_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "reference_before": {
            "filename": pair.before_path.name,
            "sha256": pair.before_sha256,
            "status": REFERENCE_BEFORE_STATUS,
        },
        "reference_after": {
            "filename": pair.after_path.name,
            "sha256": pair.after_sha256,
            "status": REFERENCE_AFTER_STATUS,
            "approval_status": APPROVAL_STATUS,
        },
        "structure_significant_changes": structure.significant_changes,
        "formula_expression_changes": len(formula_changes),
        "total_changed_cells": len(deltas),
        "direct_input_changes": by_kind.get(wd.DIRECT_INPUT_CHANGE, 0),
        "formula_result_changes": by_kind.get(wd.FORMULA_RESULT_CHANGE, 0),
        "formula_expression_value_changes": by_kind.get(wd.FORMULA_EXPRESSION_CHANGE, 0),
        "changed_cells_by_sheet": dict(sorted(by_sheet.items())),
        "provenance_direct_by_source": dict(sorted(by_source_direct.items())),
        "provenance_derived_by_source": dict(sorted(by_source_derived.items())),
        "provenance_expression_change_by_source": dict(sorted(by_source_expression.items())),
        "needs_human_review": needs_review,
        "provenance_coverage": {
            "classified": sum(1 for d in deltas if prov[d.key].source != sp.SRC_UNKNOWN_PENDING),
            "unknown_pending": sum(
                1 for d in deltas if prov[d.key].source == sp.SRC_UNKNOWN_PENDING
            ),
        },
        "consistency_checks": checks,
        "control_total_errors": control["total_errores"],
        "control_status": control["control_status"],
        "candidate_golden_status": CANDIDATE_GOLDEN_STATUS,
        "candidate_golden_pending": [
            "Medinet original de julio 2026",
            "egresos hospitalarios reales de julio 2026",
            "tabla quirúrgica de julio 2026",
            "confirmación de envío / aprobación MINSAL",
        ],
    }


def run(before_path: str, after_path: str, output_dir: str) -> dict:
    before, after = Path(before_path), Path(after_path)
    for p in (before, after):
        if not p.is_file():
            raise wd.WorkbookDeltaError(f"no es un archivo: {p}")
    sha_before_pre = wd.compute_sha256(before)
    sha_after_pre = wd.compute_sha256(after)

    pair = wd.open_pair(before, after)
    try:
        structure = wd.compare_structure(pair)
        formula_changes = wd.diff_formulas(pair)
        deltas = wd.diff_values(pair)
        wd.attach_semantic_context(pair, deltas, skip_sheets={"CONTROL"})
        for delta in deltas:
            if delta.sheet == "CONTROL":
                sheet_without_data, reason = _control_context(pair, delta)
                delta.section_path = ("CONTROL: causales de hojas sin datos",)
                delta.row_path = (sheet_without_data,) if sheet_without_data else ()
                delta.column_path = (reason,) if reason else ()
        edges = wd.build_delta_edges(deltas)
        prov = assign_provenance(pair, deltas, edges)
        checks = build_consistency_checks(pair, deltas)
        control = build_control_summary(pair)
        summary = build_summary(
            pair=pair, structure=structure, formula_changes=formula_changes,
            deltas=deltas, prov=prov, checks=checks, control=control,
        )
        summary["egresos_findings"] = _egresos_findings(pair, deltas, prov)
        files = write_outputs(
            output_dir=Path(output_dir), pair=pair, structure=structure,
            formula_changes=formula_changes, deltas=deltas, edges=edges, prov=prov,
            checks=checks, control=control, summary=summary,
        )
    finally:
        pair.close()

    sha_before_post = wd.compute_sha256(before)
    sha_after_post = wd.compute_sha256(after)
    if (sha_before_pre, sha_after_pre) != (sha_before_post, sha_after_post):
        raise wd.WorkbookDeltaError(
            "¡el SHA256 de un workbook fuente cambió durante el análisis!"
        )
    summary["source_hash_verified_unchanged"] = True
    return {
        "summary": summary, "files": files, "deltas": deltas, "prov": prov,
        "edges": edges, "checks": checks, "control": control,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="analyze_final_remasep_delta.py",
        description=(
            "Delta entre el REMASEP julio incompleto y el final, con provenance de "
            "fuente por celda. Sólo lectura, sin COM ni macros."
        ),
    )
    parser.add_argument("before", nargs="?", default=DEFAULT_BEFORE)
    parser.add_argument("after", nargs="?", default=DEFAULT_AFTER)
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = run(args.before, args.after, args.output)
    except wd.WorkbookDeltaError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    s = result["summary"]
    prov = result["prov"]
    deltas = result["deltas"]
    print(f"before: {s['reference_before']['filename']}  ({s['reference_before']['status']})")
    print(f"after : {s['reference_after']['filename']}  ({s['reference_after']['status']}, "
          f"approval={s['reference_after']['approval_status']})")
    print(f"  hash de ambos fuentes verificado sin cambios: {s['source_hash_verified_unchanged']}")
    print(f"  cambios estructurales significativos: {s['structure_significant_changes'] or 'ninguno'}")
    print(f"  formula_expression_changes: {s['formula_expression_changes']}")
    print(f"  total_changed_cells: {s['total_changed_cells']}  "
          f"(directos {s['direct_input_changes']} / derivados {s['formula_result_changes']} / "
          f"expresión reescrita {s['formula_expression_value_changes']})")
    print(f"  por hoja: {s['changed_cells_by_sheet']}")
    print(f"  provenance directo: {s['provenance_direct_by_source']}")
    print(f"  provenance derivado: {s['provenance_derived_by_source']}")
    print(f"  necesita revisión humana: {s['needs_human_review']} / {s['total_changed_cells']}")

    eg = s["egresos_findings"]
    print(f"  EGRESOS · B2 ANEXO inputs directos: {len(eg['b2_procedure_inputs'])}")
    for row in eg["b2_procedure_inputs"]:
        print(f"     {row['cell']}  {row['procedure_code']}  {row['procedure_label'][:48]!r}  = {row['count']}")
    print(f"  EGRESOS · B1 inputs directos (edad/sexo): {len(eg['b1_direct_inputs'])}")

    print("  consistency checks:")
    for c in result["checks"]:
        print(f"     {c['check_id']}: {c['left_value']} vs {c['right_value']} "
              f"(dif {c['difference']}) -> {c['status']}")

    d_section = [d for d in deltas if d.sheet == "REMASEP 01"
                 and "SECCION D" in sp._norm(_ctx(d.section_path))]
    e_section = [d for d in deltas if d.sheet == "REMASEP 01"
                 and ("SECCION E" in sp._norm(_ctx(d.section_path))
                      or "CAUSAS DE SUSPENSION" in sp._norm(_ctx(d.row_path)))]
    print(f"  REMASEP 01 Sección D (capacidad/quirófanos): {len(d_section)} celdas — "
          + str(dict(Counter(prov[d.key].source for d in d_section))))
    print(f"  REMASEP 01 Sección E (suspensiones) UNKNOWN: {len(e_section)} celdas — "
          + str(dict(Counter(prov[d.key].source for d in e_section))))

    print(f"  CONTROL: total_errores={result['control']['total_errores']} -> "
          f"{result['control']['control_status']} (approval={APPROVAL_STATUS})")
    print(f"     causales de hoja sin datos: {len(result['control']['causales'])}")
    print(f"  unresolved provenance (UNKNOWN_PENDING): "
          f"{s['provenance_coverage']['unknown_pending']}")
    print(f"  candidate_golden_status: {s['candidate_golden_status']}")
    print(f"Artefactos en {args.output}:")
    for path in result["files"]:
        print(f"  - {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
