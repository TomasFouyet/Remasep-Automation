"""Cierre completo del universo Medinet-dependent legacy (Sprint 2.5).

Sprint 2.4 evaluó las 1427 celdas ``BASE_AGGREGATION``/``DERIVED_AGGREGATION`` y
descubrió 616 celdas transitivas más: 344 ``DOWNSTREAM_TOTAL`` (``SUM``) y 272
``VALIDATION`` (``IF``). Este script evalúa esas dos últimas usando el **mismo
grafo de dependencias** ya construido (orden topológico, `depends_on_age`
propagado, ciclos/dependencias faltantes ya detectados) — sin volver a leer el
workbook ni recalcular el cierre transitivo.

    BASE (Sprint 2.3) -> DERIVED (Sprint 2.4) -> DOWNSTREAM_TOTAL -> VALIDATION
                                (orden topológico REAL, no por fases)

- Solo lectura. No ejecuta Excel, COM ni macros. No escribe el workbook.
- Los ``SUM``/``IF`` se parsean con un AST explícito
  (``remasep.services.legacy_aggregation``); sin ``eval()``.
- Una dependencia que no se pudo evaluar propaga ``DEPENDENCY_UNAVAILABLE``:
  nunca se inventa un valor.
- Las 272 ``VALIDATION`` siguen siendo comprobaciones legacy, no métricas
  clínicas: se evalúan y reportan, no se interpretan.

Uso:

    python scripts/evaluate_legacy_downstream.py \\
        "data/local/GENERACION DATOS REMASEP.xlsx"
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from dataclasses import dataclass
from pathlib import Path

import close_legacy_aggregations as clo

from remasep.services.legacy_aggregation import (
    UnsupportedFormulaError,
    arithmetic_constructs,
    formula_pattern,
    functions_used,
    parse_if_formula,
    parse_sum_formula,
)

DEFAULT_OUTPUT = "artifacts/legacy_downstream_equivalence"

EVAL_SUPPORTED = "SUPPORTED"
EVAL_DEPENDENCY_UNAVAILABLE = "DEPENDENCY_UNAVAILABLE"
EVAL_UNSUPPORTED_SUM = "UNSUPPORTED_SUM"
EVAL_UNSUPPORTED_IF = "UNSUPPORTED_IF"
EVAL_UNSUPPORTED_FORMULA = "UNSUPPORTED_FORMULA"
EVAL_CYCLE = "CYCLE"

_UNSUPPORTED_STATUSES = frozenset(
    {EVAL_UNSUPPORTED_SUM, EVAL_UNSUPPORTED_IF, EVAL_UNSUPPORTED_FORMULA}
)
_ALL_KINDS = (clo.KIND_BASE, clo.KIND_DERIVED, clo.KIND_TOTAL, clo.KIND_VALIDATION)


class DownstreamEvaluationError(clo.ClosureError):
    """Error de entrada del cierre completo, con mensaje claro."""


@dataclass
class FullClosureResult:
    closure: clo.ClosureResult
    unsupported_reasons: dict[tuple[str, str], str]

    @property
    def nodes(self) -> dict[tuple[str, str], clo.Node]:
        return self.closure.nodes

    def medinet_nodes(self) -> list[clo.Node]:
        return self.closure.medinet_nodes()

    def by_kind(self, kind: str) -> list[clo.Node]:
        return self.closure.by_kind(kind)


# ---------------------------------------------------------------------------
# Evaluación de SUM / IF sobre el DAG ya construido
# ---------------------------------------------------------------------------


def _build_value_lookup(
    node: clo.Node,
    needed_coords: tuple[str, ...],
    nodes: dict[tuple[str, str], clo.Node],
    cached: dict[str, dict[str, object]],
) -> dict[str, object] | None:
    """Valores para las celdas que necesita `node`, usando nodos ya evaluados.

    Una celda con fórmula que NO se pudo evaluar es una dependencia
    ``DEPENDENCY_UNAVAILABLE`` (se devuelve ``None``: nunca se sustituye por su
    cache). Una celda SIN fórmula (entrada manual del formulario, p.ej. un
    campo "digite EMBARAZADAS") no es una métrica Medinet: se lee su valor tal
    cual (igual que una columna raw de Atenciones), no es un atajo sobre una
    fórmula que estemos intentando demostrar.
    """
    lookup: dict[str, object] = {}
    sheet_cache = cached.get(node.sheet, {})
    for coord in needed_coords:
        dep = nodes.get((node.sheet, coord))
        if dep is not None:
            if not dep.evaluation_supported or dep.python_value is None:
                return None
            lookup[coord] = dep.python_value
        else:
            lookup[coord] = sheet_cache.get(coord)
    return lookup


def evaluate_full_closure(path: str | Path) -> FullClosureResult:
    result = clo.close_legacy_aggregations(path)
    nodes = result.nodes

    # BASE/DERIVED ya fueron evaluadas por Sprint 2.4: solo se traduce su
    # resultado a `evaluation_status` (vocabulario común con SUM/IF).
    for node in result.medinet_nodes():
        if node.kind not in (clo.KIND_BASE, clo.KIND_DERIVED):
            continue
        if node.status == clo.STATUS_CYCLE:
            node.evaluation_status = EVAL_CYCLE
        elif node.evaluation_supported:
            node.evaluation_status = EVAL_SUPPORTED
        else:
            node.evaluation_status = EVAL_DEPENDENCY_UNAVAILABLE

    for node in result.medinet_nodes():
        if node.kind == clo.KIND_OTHER:
            node.evaluation_status = EVAL_UNSUPPORTED_FORMULA

    # Los nodos en ciclo NUNCA aparecen en `topo_order` (Kahn los deja fuera):
    # se marcan aquí explícitamente para que no queden con `evaluation_status`
    # vacío, sea cual sea su `kind`.
    for key in result.cycle_nodes:
        nodes[key].evaluation_status = EVAL_CYCLE

    unsupported_reasons: dict[tuple[str, str], str] = {}

    # Único pase, en el orden topológico REAL de todo el grafo Medinet (no por
    # fases): BASE/DERIVED ya están resueltas, así que sus valores están
    # disponibles apenas les toca el turno a sus dependientes. Los nodos en
    # ciclo no están en `topo_order`, así que este bucle no los vuelve a tocar.
    for key in result.topo_order:
        node = nodes[key]
        if node.kind not in (clo.KIND_TOTAL, clo.KIND_VALIDATION):
            continue

        try:
            parsed = parse_sum_formula(node.formula) if node.kind == clo.KIND_TOTAL else (
                parse_if_formula(node.formula)
            )
        except UnsupportedFormulaError as exc:
            node.evaluation_status = (
                EVAL_UNSUPPORTED_SUM if node.kind == clo.KIND_TOTAL else EVAL_UNSUPPORTED_IF
            )
            unsupported_reasons[key] = str(exc)
            continue

        lookup = _build_value_lookup(node, parsed.value_ref_coords, nodes, result.cached)
        if lookup is None:
            node.evaluation_status = EVAL_DEPENDENCY_UNAVAILABLE
            continue
        try:
            value = parsed.evaluate(lookup)
        except UnsupportedFormulaError:
            node.evaluation_status = EVAL_DEPENDENCY_UNAVAILABLE
            continue

        node.python_value = value
        node.evaluation_supported = True
        node.evaluation_status = EVAL_SUPPORTED
        # depends_on_age ya se propagó sobre TODO el grafo Medinet (Sprint 2.4);
        # no se recalcula aquí.

    # --- comparación de cache para lo recién evaluado (TOTAL/VALIDATION) --
    for node in nodes.values():
        if node.kind not in (clo.KIND_TOTAL, clo.KIND_VALIDATION):
            continue
        if node.evaluation_status != EVAL_SUPPORTED:
            continue
        cached_value = result.cached.get(node.sheet, {}).get(node.cell)
        node.cached_value = cached_value
        node.cache_status, note = clo._compare_cache(
            node.python_value, cached_value, node.depends_on_age
        )
        if note:
            node.notes = note

    return FullClosureResult(closure=result, unsupported_reasons=unsupported_reasons)


# ---------------------------------------------------------------------------
# Inventario Parte A
# ---------------------------------------------------------------------------


def _same_sheet_dependencies(node: clo.Node) -> tuple[str, ...]:
    return tuple(sorted(coord for sheet, coord in node.graph_deps if sheet == node.sheet))


def downstream_formula_inventory(result: FullClosureResult) -> list[list]:
    transitive = sorted(
        (n for n in result.medinet_nodes() if not n.direct), key=lambda n: clo._key_sort(n.key)
    )
    rows = []
    for node in transitive:
        rows.append(
            [
                node.sheet,
                node.cell,
                node.kind,
                node.formula,
                "|".join(functions_used(node.formula)),
                "|".join(_same_sheet_dependencies(node)),
                node.depth,
                formula_pattern(node.formula),
                _b(node.depends_on_age),
            ]
        )
    return rows


# ---------------------------------------------------------------------------
# Artefactos
# ---------------------------------------------------------------------------

_README = """# Legacy downstream equivalence (Sprint 2.5)

Salida de `scripts/evaluate_legacy_downstream.py`. Cierra la evaluación del
universo Medinet-dependent completo descubierto en Sprint 2.4: además de
`BASE_AGGREGATION`/`DERIVED_AGGREGATION` (Sprint 2.3/2.4), evalúa
`DOWNSTREAM_TOTAL` (`SUM`) y `VALIDATION` (`IF`) usando el mismo grafo de
dependencias, en orden topológico real.

## Archivos

- `downstream_formula_inventory.csv` — las 616 celdas transitivas (no
  referencian `Atenciones` directamente): fórmula, dependencias de la misma
  hoja, profundidad, patrón, `depends_on_AF`.
- `full_equivalence.csv` — las ~2043 celdas Medinet-dependientes conocidas:
  `evaluation_status` (`SUPPORTED` / `DEPENDENCY_UNAVAILABLE` /
  `UNSUPPORTED_SUM` / `UNSUPPORTED_IF` / `UNSUPPORTED_FORMULA` / `CYCLE`) y
  `cache_comparison_status` por **separado**.
- `sheet_summary.csv` — conteos por hoja y tipo.
- `kind_summary.csv` — conteos por `BASE_AGGREGATION`/`DERIVED_AGGREGATION`/
  `DOWNSTREAM_TOTAL`/`VALIDATION`.
- `validation_summary.csv` — las celdas `VALIDATION`: valor Python, cache,
  `depends_on_AF`. Son comprobaciones legacy, **no** métricas clínicas.
- `summary.json` — totales; `formula_support_status` y
  `cache_consistency_status` **separados**.

## SUM downstream

Cada `SUM(...)` observado es, o bien un rango rectangular de la misma hoja
(`SUM(Q86:Q92)`), o bien una cadena `+` de celdas individuales
(`SUM(F40+H40+...+AL40)`) — la forma dominante en el workbook real. Se resuelve
sumando los **valores ya evaluados** de esas celdas en el DAG (no se relee la
cache como atajo de cálculo).

## IF validation

Las 272 fórmulas `IF` observadas tienen 4 formas: condición numérica
(`CELDA < CELDA`), condición `CELDA <> 0`, y un `IF` anidado que compara una
celda de entrada manual del formulario contra `""` (verifica si el operador
llenó un campo). Se evalúan con las mismas semánticas de comparación legacy
(`normalize_legacy_text` para `=`/`<>` de texto; numérica para el resto) — sin
convertirlas en reglas clínicas: siguen clasificadas `VALIDATION`.

## Cache y `depends_on_AF`

Igual que Sprint 2.3/2.4: se separan `evaluation_status` (¿se pudo calcular?) y
`cache_comparison_status` (¿coincide con el valor cacheado por Excel?). La señal
`depends_on_age` se propaga transitivamente por todo el grafo. Una
`CACHE_DIFFERENCE` en una fórmula que depende de `AF` **nunca** se convierte en
`MATCH` (cache de `AF` conocida como obsoleta, Sprint 2.2).

## Aviso

Reproducir `GENERACION DATOS REMASEP.xlsx` **no** es validación oficial MINSAL.
"""


def _b(value: object) -> str:
    return "True" if value else "False"


def _fmt(value: object) -> str:
    return "" if value is None else str(value)


def _csv(path: Path, header: list[str], rows) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def write_outputs(result: FullClosureResult, output_dir: Path) -> list[Path]:
    output_dir.mkdir(parents=True, exist_ok=True)
    med = sorted(result.medinet_nodes(), key=lambda n: clo._key_sort(n.key))

    inventory = output_dir / "downstream_formula_inventory.csv"
    _csv(
        inventory,
        ["sheet", "cell", "classification", "formula", "functions",
         "same_sheet_dependencies", "dependency_depth", "formula_pattern", "depends_on_AF"],
        downstream_formula_inventory(result),
    )

    full = output_dir / "full_equivalence.csv"
    _csv(
        full,
        ["metric_id", "sheet", "cell", "kind", "dependency_depth", "python_value",
         "cached_excel_value", "evaluation_status", "cache_status", "depends_on_AF"],
        (
            [
                n.metric_id, n.sheet, n.cell, n.kind, n.depth, _fmt(n.python_value),
                _fmt(n.cached_value), n.evaluation_status,
                n.cache_status or ("" if n.evaluation_status != EVAL_SUPPORTED else clo.CACHE_UNAVAILABLE),
                _b(n.depends_on_age),
            ]
            for n in med
        ),
    )

    sheet_summary = output_dir / "sheet_summary.csv"
    _csv(
        sheet_summary,
        ["sheet", "base_total", "base_supported", "derived_total", "derived_supported",
         "downstream_total", "downstream_supported", "validation_total", "validation_supported",
         "cache_match", "cache_difference", "cache_unavailable"],
        (_sheet_row(result, sheet) for sheet in _present_sheets(result)),
    )

    kind_summary = output_dir / "kind_summary.csv"
    _csv(
        kind_summary,
        ["kind", "total", "supported", "cache_match", "cache_difference", "cache_unavailable"],
        (_kind_row(result, kind) for kind in _ALL_KINDS),
    )

    validation = output_dir / "validation_summary.csv"
    _csv(
        validation,
        ["sheet", "cell", "python_value", "cached_value", "cache_status", "depends_on_AF",
         "formula_pattern"],
        (
            [n.sheet, n.cell, _fmt(n.python_value), _fmt(n.cached_value),
             n.cache_status or "", _b(n.depends_on_age), formula_pattern(n.formula)]
            for n in result.by_kind(clo.KIND_VALIDATION)
        ),
    )

    summary_path = output_dir / "summary.json"
    summary_path.write_text(
        json.dumps(_summary_dict(result), indent=2, ensure_ascii=False), encoding="utf-8"
    )

    readme = output_dir / "README.md"
    readme.write_text(_README, encoding="utf-8")

    return [inventory, full, sheet_summary, kind_summary, validation, summary_path, readme]


def _present_sheets(result: FullClosureResult) -> list[str]:
    return [s for s in clo.TARGET_SHEETS if any(n.sheet == s for n in result.medinet_nodes())]


def _sheet_row(result: FullClosureResult, sheet: str) -> list:
    cells = [n for n in result.medinet_nodes() if n.sheet == sheet]

    def _count(kind: str, supported: bool | None = None) -> int:
        subset = [n for n in cells if n.kind == kind]
        if supported is None:
            return len(subset)
        return len([n for n in subset if (n.evaluation_status == EVAL_SUPPORTED) == supported])

    evaluated = [n for n in cells if n.evaluation_status == EVAL_SUPPORTED]
    return [
        sheet,
        _count(clo.KIND_BASE), _count(clo.KIND_BASE, True),
        _count(clo.KIND_DERIVED), _count(clo.KIND_DERIVED, True),
        _count(clo.KIND_TOTAL), _count(clo.KIND_TOTAL, True),
        _count(clo.KIND_VALIDATION), _count(clo.KIND_VALIDATION, True),
        len([n for n in evaluated if n.cache_status == clo.CACHE_MATCH]),
        len([n for n in evaluated if n.cache_status == clo.CACHE_DIFFERENCE]),
        len([n for n in evaluated if n.cache_status == clo.CACHE_UNAVAILABLE]),
    ]


def _kind_row(result: FullClosureResult, kind: str) -> list:
    cells = result.by_kind(kind)
    supported = [n for n in cells if n.evaluation_status == EVAL_SUPPORTED]
    return [
        kind,
        len(cells),
        len(supported),
        len([n for n in supported if n.cache_status == clo.CACHE_MATCH]),
        len([n for n in supported if n.cache_status == clo.CACHE_DIFFERENCE]),
        len([n for n in supported if n.cache_status == clo.CACHE_UNAVAILABLE]),
    ]


def _summary_dict(result: FullClosureResult) -> dict:
    med = result.medinet_nodes()
    closure = result.closure
    base = result.by_kind(clo.KIND_BASE)
    derived = result.by_kind(clo.KIND_DERIVED)
    total = result.by_kind(clo.KIND_TOTAL)
    validation = result.by_kind(clo.KIND_VALIDATION)
    supported = [n for n in med if n.evaluation_status == EVAL_SUPPORTED]
    af_cells = [n for n in med if n.depends_on_age]
    af_diff = [n for n in supported if n.cache_status == clo.CACHE_DIFFERENCE and n.depends_on_age]
    unsupported_cells = [n for n in med if n.evaluation_status in _UNSUPPORTED_STATUSES]
    dependency_unavailable = [n for n in med if n.evaluation_status == EVAL_DEPENDENCY_UNAVAILABLE]
    active_validations = [n for n in validation if n.evaluation_status == EVAL_SUPPORTED and n.python_value not in (0, "", None)]

    base_supported = len([n for n in base if n.evaluation_status == EVAL_SUPPORTED])
    if base and base_supported < len(base):
        formula_support_status = "FAIL"
    elif len(supported) == len(med):
        formula_support_status = "PASS"
    elif supported:
        formula_support_status = "PARTIAL"
    else:
        formula_support_status = "FAIL"

    if any(n.cache_status == clo.CACHE_DIFFERENCE for n in supported):
        cache_consistency_status = "DIFFERENCES"
    elif any(n.cache_status == clo.CACHE_UNAVAILABLE for n in supported):
        cache_consistency_status = "UNAVAILABLE"
    else:
        cache_consistency_status = "PASS"

    return {
        "source": closure.source_name,
        "source_sha256": closure.source_sha256,
        "detail_sheet": closure.detail_sheet,
        "generated_at": closure.generated_at,
        "physical_rows": closure.physical_rows,
        "structural_empty_rows": closure.structural_empty_rows,
        "active_records": closure.active_records,
        "total_medinet_dependent_cells": len(med),
        "base_total": len(base),
        "base_supported": base_supported,
        "derived_total": len(derived),
        "derived_supported": len([n for n in derived if n.evaluation_status == EVAL_SUPPORTED]),
        "downstream_total": len(total),
        "downstream_supported": len([n for n in total if n.evaluation_status == EVAL_SUPPORTED]),
        "validation_total": len(validation),
        "validation_supported": len([n for n in validation if n.evaluation_status == EVAL_SUPPORTED]),
        "fully_evaluated_cells": len(supported),
        "unsupported_cells": len(unsupported_cells),
        "dependency_unavailable_cells": len(dependency_unavailable),
        "max_dependency_depth": max((n.depth for n in med), default=0),
        "cycles": len(closure.cycle_nodes),
        "missing_dependencies": len(closure.missing_dependencies),
        "cache_matches": len([n for n in supported if n.cache_status == clo.CACHE_MATCH]),
        "cache_differences": len([n for n in supported if n.cache_status == clo.CACHE_DIFFERENCE]),
        "cache_unavailable": len([n for n in supported if n.cache_status == clo.CACHE_UNAVAILABLE]),
        "cache_differences_depending_on_AF": len(af_diff),
        "depends_on_AF_cells": len(af_cells),
        "active_validations_nonzero": len(active_validations),
        "formula_support_status": formula_support_status,
        "cache_consistency_status": cache_consistency_status,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="evaluate_legacy_downstream.py",
        description=(
            "Cierra la evaluación del universo Medinet-dependent legacy: BASE + "
            "DERIVED (Sprint 2.3/2.4) + DOWNSTREAM_TOTAL (SUM) + VALIDATION (IF). "
            "Solo lectura."
        ),
    )
    parser.add_argument("workbook", help="Ruta al workbook legacy (.xlsx/.xlsm).")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"(def: {DEFAULT_OUTPUT})")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    try:
        result = evaluate_full_closure(args.workbook)
    except clo.ClosureError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    files = write_outputs(result, Path(args.output))
    summary = _summary_dict(result)
    inventory = downstream_formula_inventory(result)
    sum_patterns = {r[7] for r in inventory if r[2] == clo.KIND_TOTAL}
    if_patterns = {r[7] for r in inventory if r[2] == clo.KIND_VALIDATION}
    secondary = set()
    for n in result.medinet_nodes():
        if not n.direct:
            secondary.update(arithmetic_constructs(n.formula))

    print(f"Workbook: {result.closure.source_name}")
    print(f"  celdas Medinet-dependent: {summary['total_medinet_dependent_cells']}")
    print(
        f"  base {summary['base_total']}/{summary['base_supported']} | "
        f"derived {summary['derived_total']}/{summary['derived_supported']} | "
        f"downstream_total {summary['downstream_total']}/{summary['downstream_supported']} | "
        f"validation {summary['validation_total']}/{summary['validation_supported']}"
    )
    print(f"  patrones SUM: {len(sum_patterns)} | patrones IF: {len(if_patterns)}")
    print(f"  constructs secundarios en celdas transitivas: {sorted(secondary) or 'ninguno'}")
    print(
        f"  totalmente evaluadas: {summary['fully_evaluated_cells']}/"
        f"{summary['total_medinet_dependent_cells']} | "
        f"unsupported: {summary['unsupported_cells']} | "
        f"dependency_unavailable: {summary['dependency_unavailable_cells']}"
    )
    print(
        f"  DAG: profundidad máx {summary['max_dependency_depth']} | "
        f"ciclos {summary['cycles']} | dependencias faltantes {summary['missing_dependencies']}"
    )
    print(
        f"  cache: {summary['cache_matches']} MATCH / {summary['cache_differences']} DIFF / "
        f"{summary['cache_unavailable']} UNAVAIL (dependientes de AF: "
        f"{summary['cache_differences_depending_on_AF']})"
    )
    print(f"  validaciones con valor Python distinto de cero/vacío: {summary['active_validations_nonzero']}")
    print(f"  formula_support_status: {summary['formula_support_status']}")
    print(f"  cache_consistency_status: {summary['cache_consistency_status']}")
    print(f"Artefactos en {args.output}:")
    for path in files:
        print(f"  - {path}")
    return 0 if summary["formula_support_status"] != "FAIL" else 1


if __name__ == "__main__":
    raise SystemExit(main())
