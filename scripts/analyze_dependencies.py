"""Inventario de dependencias técnicas de fórmulas del workbook REMASEP.

Sprint 1.2. Solo lectura, Linux/openpyxl, sin COM. NO infiere semántica REMASEP
ni clínica: solo describe qué celda/columna/hoja depende técnicamente de cuál,
usando SIEMPRE el texto original de la fórmula.

No se exportan valores de celdas: los CSV solo contienen coordenadas, letras de
columna, nombres de hoja y el texto de las fórmulas.

Uso:

    python scripts/analyze_dependencies.py \\
        "data/local/GENERACION DATOS REMASEP.xlsx" \\
        --output artifacts/workbook_inventory \\
        --flow-doc docs/CURRENT_EXCEL_FLOW.md
"""

from __future__ import annotations

import argparse
import csv
import sys
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path

from formula_refs import Reference, col_to_index, index_to_col, parse_references
from inventory_workbook import (
    InventoryError,
    cell_formula_text,
    compute_sha256,
    extract_function_names,
    load_workbook_safely,
    normalize_formula,
)

DETAIL_SHEET_DEFAULT = "Atenciones - Detalles de citas"
DEFAULT_OUTPUT = "artifacts/workbook_inventory"
DEFAULT_FLOW_DOC = "docs/CURRENT_EXCEL_FLOW.md"

# Tope defensivo al expandir un rango de columnas a letras individuales.
_MAX_COLUMN_SPAN = 256

LOCAL_SHEET = "(local)"


@dataclass
class FormulaCell:
    sheet: str
    coord: str
    col: str
    col_index: int
    row: int
    formula: str
    functions: tuple[str, ...]
    normalized: str
    references: tuple[Reference, ...]
    unsupported: tuple[str, ...]


# ---------------------------------------------------------------------------
# Recolección
# ---------------------------------------------------------------------------


def collect_formula_cells(wb) -> list[FormulaCell]:
    cells: list[FormulaCell] = []
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type != "f":
                    continue
                formula = cell_formula_text(cell.value)
                if formula is None:
                    formula = str(cell.value)
                parsed = parse_references(formula)
                cells.append(
                    FormulaCell(
                        sheet=ws.title,
                        coord=cell.coordinate,
                        col=cell.column_letter,
                        col_index=cell.column,
                        row=cell.row,
                        formula=formula,
                        functions=tuple(extract_function_names(formula)),
                        normalized=normalize_formula(formula),
                        references=tuple(parsed.references),
                        unsupported=tuple(parsed.unsupported),
                    )
                )
    return cells


def _expand_columns(ref: Reference, wide_spans: list[tuple[str, str]]) -> list[int]:
    """Índices de columna cubiertos por una referencia (vacío si es whole_row)."""
    if ref.col_start is None or ref.col_end is None:
        return []
    start, end = col_to_index(ref.col_start), col_to_index(ref.col_end)
    if end - start + 1 > _MAX_COLUMN_SPAN:
        wide_spans.append((ref.col_start, ref.col_end))
        return [start, end]
    return list(range(start, end + 1))


# ---------------------------------------------------------------------------
# Parte B — Derived columns
# ---------------------------------------------------------------------------

DERIVED_COLUMNS_HEADER = [
    "target_column",
    "formula_count",
    "first_formula_cell",
    "example_formula",
    "source_columns",
    "source_sheets",
    "function_names",
    "pattern_count",
]


def build_derived_columns(
    cells: list[FormulaCell], detail_sheet: str, wide_spans: list[tuple[str, str]]
) -> list[list]:
    by_column: dict[str, list[FormulaCell]] = defaultdict(list)
    for cell in cells:
        if cell.sheet == detail_sheet:
            by_column[cell.col].append(cell)

    rows: list[list] = []
    for column in sorted(by_column, key=col_to_index):
        members = sorted(by_column[column], key=lambda c: (c.row, c.col_index))
        first = members[0]

        source_columns: set[int] = set()
        source_sheets: set[str] = set()
        functions: set[str] = set()
        patterns: set[str] = set()
        for member in members:
            functions.update(member.functions)
            patterns.add(member.normalized)
            for ref in member.references:
                source_sheets.add(ref.sheet if ref.sheet is not None else LOCAL_SHEET)
                source_columns.update(_expand_columns(ref, wide_spans))

        rows.append(
            [
                column,
                len(members),
                first.coord,
                first.formula,
                "|".join(index_to_col(i) for i in sorted(source_columns)),
                "|".join(sorted(source_sheets)),
                "|".join(sorted(functions)),
                len(patterns),
            ]
        )
    return rows


# ---------------------------------------------------------------------------
# Parte C — Formula dependencies
# ---------------------------------------------------------------------------

FORMULA_DEPENDENCIES_HEADER = [
    "target_sheet",
    "target_cell",
    "source_sheet",
    "source_reference",
    "source_reference_type",
    "source_column_start",
    "source_column_end",
    "is_cross_sheet",
]


def build_formula_dependencies(
    cells: list[FormulaCell], sheet_order: dict[str, int], known_sheets: set[str]
) -> tuple[list[list], list[tuple[str, str, str]]]:
    ordered = sorted(cells, key=lambda c: (sheet_order.get(c.sheet, 0), c.row, c.col_index))
    rows: list[list] = []
    unknown: list[tuple[str, str, str]] = []
    for cell in ordered:
        for ref in cell.references:
            if ref.sheet is None:
                source_sheet = cell.sheet
                is_cross = False
            else:
                source_sheet = ref.sheet
                is_cross = source_sheet != cell.sheet
                if source_sheet not in known_sheets:
                    unknown.append((cell.sheet, cell.coord, source_sheet))
            rows.append(
                [
                    cell.sheet,
                    cell.coord,
                    source_sheet,
                    ref.raw,
                    ref.ref_type,
                    ref.col_start or "",
                    ref.col_end or "",
                    is_cross,
                ]
            )
    return rows, unknown


# ---------------------------------------------------------------------------
# Parte D — Sheet dependency graph
# ---------------------------------------------------------------------------

SHEET_DEPENDENCIES_HEADER = ["source_sheet", "target_sheet", "reference_count"]


def build_sheet_dependencies(dependency_rows: list[list]) -> list[list]:
    counter: Counter[tuple[str, str]] = Counter()
    for target_sheet, _cell, source_sheet, *_rest in dependency_rows:
        counter[(source_sheet, target_sheet)] += 1
    rows = [[source, target, count] for (source, target), count in counter.items()]
    rows.sort(key=lambda r: (-r[2], r[0], r[1]))
    return rows


# ---------------------------------------------------------------------------
# Parte E — Output column dependencies
# ---------------------------------------------------------------------------

OUTPUT_COLUMN_DEPENDENCIES_HEADER = [
    "target_sheet",
    "source_column",
    "formulas_using_column",
    "example_target_cell",
    "example_formula",
]


def build_output_column_dependencies(
    cells: list[FormulaCell],
    detail_sheet: str,
    output_sheets: list[str],
    wide_spans: list[tuple[str, str]],
) -> list[list]:
    output_set = set(output_sheets)
    accumulator: dict[tuple[str, str], list[FormulaCell]] = defaultdict(list)
    for cell in cells:
        if cell.sheet not in output_set:
            continue
        columns_hit: set[int] = set()
        for ref in cell.references:
            if ref.sheet == detail_sheet:
                columns_hit.update(_expand_columns(ref, wide_spans))
        for index in sorted(columns_hit):
            accumulator[(cell.sheet, index_to_col(index))].append(cell)

    rows: list[list] = []
    for (target_sheet, source_column), members in accumulator.items():
        example = min(members, key=lambda c: (c.row, c.col_index))
        rows.append(
            [target_sheet, source_column, len(members), example.coord, example.formula]
        )
    rows.sort(key=lambda r: (r[0], col_to_index(r[1])))
    return rows


# ---------------------------------------------------------------------------
# Unsupported constructs
# ---------------------------------------------------------------------------

UNSUPPORTED_HEADER = ["sheet", "construct_type", "occurrences", "example_cell", "example_formula"]


def build_unsupported(cells: list[FormulaCell]) -> list[list]:
    aggregate: dict[tuple[str, str], list] = {}
    for cell in cells:
        for construct in cell.unsupported:
            key = (cell.sheet, construct)
            if key not in aggregate:
                aggregate[key] = [0, cell.coord, cell.formula]
            aggregate[key][0] += 1
    rows = [
        [sheet, construct, count, coord, formula]
        for (sheet, construct), (count, coord, formula) in aggregate.items()
    ]
    rows.sort(key=lambda r: (r[0], r[1]))
    return rows


# ---------------------------------------------------------------------------
# docs/CURRENT_EXCEL_FLOW.md
# ---------------------------------------------------------------------------


def _mermaid_id(name: str, ids: dict[str, str]) -> str:
    if name not in ids:
        ids[name] = f"s{len(ids)}"
    return ids[name]


def render_flow_doc(
    *,
    sheet_dependency_rows: list[list],
    unsupported_rows: list[list],
    sheet_names: list[str],
    detail_sheet: str,
    output_sheets: list[str],
    file_name: str,
    sha256: str,
) -> str:
    cross = [(s, t, n) for s, t, n in sheet_dependency_rows if s != t]
    self_refs = [(s, n) for s, t, n in sheet_dependency_rows if s == t]

    ids: dict[str, str] = {}
    lines: list[str] = []
    lines.append("# Flujo actual del Excel (dependencias observadas)")
    lines.append("")
    lines.append(
        "> Generado por `scripts/analyze_dependencies.py`. Refleja **solo** "
        "dependencias técnicas entre celdas según el texto de las fórmulas: qué "
        "hoja lee celdas de qué otra hoja. **No** representa semántica REMASEP ni "
        "clínica y no infiere conexiones que no aparezcan en una fórmula."
    )
    lines.append(">")
    lines.append(f"> Archivo analizado: `{file_name}`  ")
    lines.append(f"> SHA256: `{sha256}`  ")
    lines.append(f"> Generado: {datetime.now(UTC).strftime('%Y-%m-%dT%H:%M:%SZ')}  ")
    lines.append(f"> Hoja de detalle: `{detail_sheet}`  ")
    lines.append(f"> Hojas output analizadas: {', '.join(f'`{s}`' for s in output_sheets) or '—'}")
    lines.append("")
    lines.append("## Grafo de dependencias entre hojas")
    lines.append("")
    lines.append("Aristas `origen --(nº referencias)--> destino`: las fórmulas de *destino*")
    lines.append("leen celdas de *origen*. Auto-referencias omitidas del grafo (ver tabla).")
    lines.append("")
    lines.append("```mermaid")
    lines.append("graph LR")
    for name in sheet_names:
        lines.append(f'    {_mermaid_id(name, ids)}["{name}"]')
    if not cross:
        lines.append("    %% no se observaron dependencias entre hojas distintas")
    for source, target, count in sorted(cross, key=lambda r: (-r[2], r[0], r[1])):
        src = _mermaid_id(source, ids)
        dst = _mermaid_id(target, ids)
        lines.append(f"    {src} -->|{count}| {dst}")
    lines.append("```")
    lines.append("")

    lines.append("## Aristas observadas (cross-sheet)")
    lines.append("")
    lines.append("| source_sheet | target_sheet | reference_count |")
    lines.append("| --- | --- | --- |")
    for source, target, count in sorted(cross, key=lambda r: (-r[2], r[0], r[1])):
        lines.append(f"| {source} | {target} | {count} |")
    if not cross:
        lines.append("| — | — | 0 |")
    lines.append("")

    lines.append("## Auto-referencias (fórmulas que leen su propia hoja)")
    lines.append("")
    lines.append("| sheet | reference_count |")
    lines.append("| --- | --- |")
    for sheet, count in sorted(self_refs, key=lambda r: (-r[1], r[0])):
        lines.append(f"| {sheet} | {count} |")
    if not self_refs:
        lines.append("| — | 0 |")
    lines.append("")

    lines.append("## Constructs no soportados")
    lines.append("")
    if unsupported_rows:
        lines.append("| sheet | construct_type | occurrences |")
        lines.append("| --- | --- | --- |")
        for sheet, construct, count, _cell, _formula in unsupported_rows:
            lines.append(f"| {sheet} | {construct} | {count} |")
    else:
        lines.append("No se encontraron `INDIRECT`, `OFFSET`, referencias externas,")
        lines.append("*structured references*, referencias 3D ni errores `#REF!`.")
    lines.append("")
    return "\n".join(lines)


# ---------------------------------------------------------------------------
# README del output
# ---------------------------------------------------------------------------

_DEPENDENCIES_README = """# Dependency inventory (Sprint 1.2)

Salida de `scripts/analyze_dependencies.py`. Análisis estático de solo lectura de
las **dependencias técnicas** entre fórmulas del workbook. No hay semántica
REMASEP ni valores de pacientes: solo coordenadas, letras de columna, nombres de
hoja y texto de fórmulas. Todas las referencias se derivan del texto ORIGINAL de
la fórmula (no de `formula_patterns.csv`).

## `derived_columns.csv`
Una fila por columna de la hoja de detalle que contiene fórmulas.

- `target_column`: letra de la columna derivada.
- `formula_count`: nº de celdas con fórmula en esa columna.
- `first_formula_cell` / `example_formula`: primera celda (por fila) y su fórmula.
- `source_columns`: letras de columna referenciadas por esas fórmulas (rangos
  expandidos a columnas; `whole_row` no aporta columnas). Son nombres técnicos,
  no valores.
- `source_sheets`: hojas referenciadas; `(local)` = misma hoja, sin cualificar.
- `function_names`: funciones usadas (unión, `_xlfn.` limpiado).
- `pattern_count`: nº de fórmulas normalizadas distintas en la columna (indica si
  todas las filas comparten la misma fórmula o hay variantes).

## `formula_dependencies.csv`
Una fila por dependencia única detectada en cada celda con fórmula.

- `target_sheet` / `target_cell`: la celda que contiene la fórmula.
- `source_sheet`: hoja referenciada (para referencias locales = `target_sheet`).
- `source_reference`: la referencia tal cual aparece (simbólica; `A:A` no se
  expande).
- `source_reference_type`: `cell`, `range`, `whole_column` o `whole_row`.
- `source_column_start` / `source_column_end`: letras de columna del rango
  (vacío para `whole_row`).
- `is_cross_sheet`: `True` si la hoja fuente difiere de la hoja destino.

## `sheet_dependencies.csv`
Agregado de `formula_dependencies.csv` por par de hojas.

- `source_sheet` -> `target_sheet`: las fórmulas de `target_sheet` leen celdas de
  `source_sheet`.
- `reference_count`: nº de filas de dependencia para ese par (incluye
  auto-referencias cuando `source_sheet == target_sheet`).

## `output_column_dependencies.csv`
Qué columnas de la hoja de detalle usa cada hoja "output" (todas las hojas con
fórmulas salvo la de detalle, o las indicadas con `--output-sheets`).

- `target_sheet`: hoja output.
- `source_column`: letra de columna de la hoja de detalle.
- `formulas_using_column`: nº de celdas de `target_sheet` que la referencian.
- `example_target_cell` / `example_formula`: una de esas celdas y su fórmula.

## `unsupported_constructs.csv`
Constructs que el analizador NO resuelve, reportados explícitamente (nunca se
infiere su resultado): `indirect`, `offset`, `external_reference`,
`structured_reference`, `three_d_reference`, `ref_error`, `name_error`.
Agregado por `(sheet, construct_type)` con un ejemplo.

## Limitaciones
- Analizador de referencias por regex, no un parser de Excel completo. Un named
  range con forma de celda (p.ej. `Q1`) sería indistinguible de una referencia;
  el workbook de referencia no define named ranges.
- Rangos con más de {max_span} columnas no se expanden (se listan `start`/`end`).
- `INDIRECT`/`OFFSET`: se listan sus argumentos literales como referencias, pero
  el rango efectivo es dinámico y queda como no soportado.
""".replace("{max_span}", str(_MAX_COLUMN_SPAN))


# ---------------------------------------------------------------------------
# Orquestación
# ---------------------------------------------------------------------------


def _write_csv(path: Path, header: list[str], rows: list[list]) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


@dataclass
class AnalysisResult:
    output_dir: Path
    flow_doc: Path
    files: list[Path]
    detail_sheet: str
    detail_sheet_present: bool
    output_sheets: list[str]
    derived_column_count: int
    dependency_count: int
    sheet_edges: list[list]
    output_column_rows: list[list]
    unsupported_rows: list[list]
    unknown_sheet_refs: list[tuple[str, str, str]]
    wide_spans: list[tuple[str, str]]


def analyze_dependencies(
    input_path: Path,
    output_dir: Path,
    *,
    detail_sheet: str = DETAIL_SHEET_DEFAULT,
    flow_doc: Path | None = None,
    output_sheets: list[str] | None = None,
) -> AnalysisResult:
    flow_doc = flow_doc or Path(DEFAULT_FLOW_DOC)
    sha256 = compute_sha256(input_path) if input_path.is_file() else ""

    wb = load_workbook_safely(input_path)
    try:
        sheet_names = list(wb.sheetnames)
        cells = collect_formula_cells(wb)
    finally:
        wb.close()

    sheet_order = {name: index for index, name in enumerate(sheet_names)}
    known_sheets = set(sheet_names)
    detail_present = detail_sheet in known_sheets

    if output_sheets is None:
        sheets_with_formulas = {c.sheet for c in cells}
        output_sheets = [
            name for name in sheet_names if name != detail_sheet and name in sheets_with_formulas
        ]

    wide_spans: list[tuple[str, str]] = []

    derived_rows = build_derived_columns(cells, detail_sheet, wide_spans)
    dependency_rows, unknown_refs = build_formula_dependencies(cells, sheet_order, known_sheets)
    sheet_rows = build_sheet_dependencies(dependency_rows)
    output_column_rows = build_output_column_dependencies(
        cells, detail_sheet, output_sheets, wide_spans
    )
    unsupported_rows = build_unsupported(cells)

    output_dir.mkdir(parents=True, exist_ok=True)
    derived_path = output_dir / "derived_columns.csv"
    dependencies_path = output_dir / "formula_dependencies.csv"
    sheet_path = output_dir / "sheet_dependencies.csv"
    output_column_path = output_dir / "output_column_dependencies.csv"
    unsupported_path = output_dir / "unsupported_constructs.csv"
    readme_path = output_dir / "dependencies_README.md"

    _write_csv(derived_path, DERIVED_COLUMNS_HEADER, derived_rows)
    _write_csv(dependencies_path, FORMULA_DEPENDENCIES_HEADER, dependency_rows)
    _write_csv(sheet_path, SHEET_DEPENDENCIES_HEADER, sheet_rows)
    _write_csv(output_column_path, OUTPUT_COLUMN_DEPENDENCIES_HEADER, output_column_rows)
    _write_csv(unsupported_path, UNSUPPORTED_HEADER, unsupported_rows)
    readme_path.write_text(_DEPENDENCIES_README, encoding="utf-8")

    flow_doc.parent.mkdir(parents=True, exist_ok=True)
    flow_doc.write_text(
        render_flow_doc(
            sheet_dependency_rows=sheet_rows,
            unsupported_rows=unsupported_rows,
            sheet_names=sheet_names,
            detail_sheet=detail_sheet,
            output_sheets=output_sheets,
            file_name=input_path.name,
            sha256=sha256,
        )
        + "\n",
        encoding="utf-8",
    )

    return AnalysisResult(
        output_dir=output_dir,
        flow_doc=flow_doc,
        files=[
            derived_path,
            dependencies_path,
            sheet_path,
            output_column_path,
            unsupported_path,
            readme_path,
            flow_doc,
        ],
        detail_sheet=detail_sheet,
        detail_sheet_present=detail_present,
        output_sheets=output_sheets,
        derived_column_count=len(derived_rows),
        dependency_count=len(dependency_rows),
        sheet_edges=sheet_rows,
        output_column_rows=output_column_rows,
        unsupported_rows=unsupported_rows,
        unknown_sheet_refs=unknown_refs,
        wide_spans=wide_spans,
    )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="analyze_dependencies.py",
        description=(
            "Inventario de dependencias técnicas de fórmulas (columnas derivadas, "
            "dependencias celda a celda, grafo de hojas y uso de columnas de la "
            "hoja de detalle desde las hojas output). Solo lectura."
        ),
    )
    parser.add_argument("workbook", help="Ruta al archivo .xlsx/.xlsm a analizar.")
    parser.add_argument(
        "--output",
        "-o",
        default=DEFAULT_OUTPUT,
        help=f"Directorio de salida (por defecto: {DEFAULT_OUTPUT}).",
    )
    parser.add_argument(
        "--detail-sheet",
        default=DETAIL_SHEET_DEFAULT,
        help=f"Hoja de detalle por atención (por defecto: {DETAIL_SHEET_DEFAULT!r}).",
    )
    parser.add_argument(
        "--flow-doc",
        default=DEFAULT_FLOW_DOC,
        help=f"Ruta del documento Mermaid a generar (por defecto: {DEFAULT_FLOW_DOC}).",
    )
    parser.add_argument(
        "--output-sheets",
        default=None,
        help=(
            "Lista separada por comas de hojas 'output' para la Parte E. "
            "Por defecto: todas las hojas con fórmulas salvo la de detalle."
        ),
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    input_path = Path(args.workbook)
    output_dir = Path(args.output)
    flow_doc = Path(args.flow_doc)
    output_sheets = (
        [s.strip() for s in args.output_sheets.split(",") if s.strip()]
        if args.output_sheets
        else None
    )

    try:
        result = analyze_dependencies(
            input_path,
            output_dir,
            detail_sheet=args.detail_sheet,
            flow_doc=flow_doc,
            output_sheets=output_sheets,
        )
    except InventoryError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print(f"Workbook analizado: {input_path}")
    if not result.detail_sheet_present:
        print(
            f"AVISO: la hoja de detalle {result.detail_sheet!r} no existe; "
            "derived_columns.csv y output_column_dependencies.csv quedan vacíos.",
            file=sys.stderr,
        )
    print(f"  columnas derivadas detectadas: {result.derived_column_count}")
    print(f"  dependencias (filas): {result.dependency_count}")
    print(f"  hojas output: {', '.join(result.output_sheets) or '—'}")
    print("  grafo de hojas (source -> target : refs):")
    for source, target, count in result.sheet_edges:
        marker = " (auto)" if source == target else ""
        print(f"    {source} -> {target} : {count}{marker}")
    if result.unsupported_rows:
        print("  constructs no soportados:")
        for sheet, construct, count, _cell, _formula in result.unsupported_rows:
            print(f"    {sheet}: {construct} x{count}")
    else:
        print("  constructs no soportados: ninguno")
    if result.unknown_sheet_refs:
        print(f"  referencias a hojas desconocidas: {len(result.unknown_sheet_refs)}")
    if result.wide_spans:
        print(f"  rangos de columnas muy anchos (no expandidos): {len(result.wide_spans)}")
    print(f"Artefactos escritos en: {output_dir}")
    for path in result.files:
        print(f"  - {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
