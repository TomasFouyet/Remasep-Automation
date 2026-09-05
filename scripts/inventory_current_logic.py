"""Inventario de la lógica ACTUAL de "Atenciones - Detalles de citas".

Sprint 1.3. Solo lectura, Linux/openpyxl, sin COM. Documenta *exactamente* lo que
hacen hoy las fórmulas del workbook: NO traduce a reglas REMASEP oficiales, NO
juzga corrección clínica y NO infiere lógica que no esté escrita en una fórmula.

Privacidad: solo se leen/exportan encabezados de la hoja de detalle, fórmulas,
coordenadas y metadatos. Nunca valores de las filas de pacientes.

Uso:

    python scripts/inventory_current_logic.py \\
        "data/local/GENERACION DATOS REMASEP.xlsx" \\
        --output artifacts/current_logic \\
        --doc docs/CURRENT_LOGIC.md
"""

from __future__ import annotations

import argparse
import csv
import re
import sys
import unicodedata
from collections import defaultdict
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path

from analyze_dependencies import DETAIL_SHEET_DEFAULT, collect_formula_cells
from formula_refs import col_to_index, index_to_col
from inventory_workbook import (
    InventoryError,
    compute_sha256,
    load_workbook_safely,
)

DEFAULT_OUTPUT = "artifacts/current_logic"
DEFAULT_DOC = "docs/CURRENT_LOGIC.md"
_MAX_COLUMN_SPAN = 256

# Tipos técnicos de transformación permitidos (no clínicos).
CONCAT = "CONCAT"
AGE_DATEDIF = "AGE_DATEDIF"
EXACT_MAP = "EXACT_MAP"
PATTERN_FLAG = "PATTERN_FLAG"
UNKNOWN = "UNKNOWN"

# Columnas cuya transformación implementa una clasificación observable.
_RULE_TYPES = {EXACT_MAP, PATTERN_FLAG}

_STRING_RE = re.compile(r'"((?:[^"]|"")*)"')
_A1_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")
_IFS_COND_RE = re.compile(
    r'^\s*(\$?[A-Za-z]{1,3}\$?[0-9]+)\s*(=|<>)\s*("(?:[^"]|"")*")\s*$'
)


# ---------------------------------------------------------------------------
# Utilidades de texto / fórmula
# ---------------------------------------------------------------------------


def _norm_header(text: str) -> str:
    text = unicodedata.normalize("NFD", str(text).strip().upper())
    return "".join(ch for ch in text if unicodedata.category(ch) != "Mn")


def _string_literals(formula: str) -> list[str]:
    return [m.group(1).replace('""', '"') for m in _STRING_RE.finditer(formula)]


def _unquote(token: str) -> str:
    token = token.strip()
    if len(token) >= 2 and token[0] == '"' and token[-1] == '"':
        return token[1:-1].replace('""', '"')
    return token


def _split_top_level(text: str, sep: str = ",") -> list[str]:
    parts: list[str] = []
    buf: list[str] = []
    depth = 0
    in_str = False
    i = 0
    while i < len(text):
        ch = text[i]
        if in_str:
            if ch == '"':
                if i + 1 < len(text) and text[i + 1] == '"':
                    buf.append('""')
                    i += 2
                    continue
                in_str = False
            buf.append(ch)
        elif ch == '"':
            in_str = True
            buf.append(ch)
        elif ch in "([{":
            depth += 1
            buf.append(ch)
        elif ch in ")]}":
            depth -= 1
            buf.append(ch)
        elif ch == sep and depth == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    parts.append("".join(buf))
    return parts


def _call_contents(formula: str, name: str) -> list[str]:
    """Contenido (texto entre paréntesis) de cada llamada ``name(...)`` de la fórmula."""
    pattern = re.compile(rf"(?<![A-Za-z0-9_])(?:_xl\w+\.)?{re.escape(name)}\s*\(", re.IGNORECASE)
    contents: list[str] = []
    for match in pattern.finditer(formula):
        depth = 0
        in_str = False
        start = match.end() - 1  # posición del '('
        j = start
        while j < len(formula):
            ch = formula[j]
            if in_str:
                if ch == '"':
                    if j + 1 < len(formula) and formula[j + 1] == '"':
                        j += 2
                        continue
                    in_str = False
            elif ch == '"':
                in_str = True
            elif ch == "(":
                depth += 1
            elif ch == ")":
                depth -= 1
                if depth == 0:
                    contents.append(formula[start + 1 : j])
                    break
            j += 1
    return contents


def _column_of_ref(token: str) -> str | None:
    match = _A1_RE.match(token.strip())
    return match.group(1).upper() if match else None


def _has_wildcard_literal(formula: str) -> bool:
    return any("*" in lit or "?" in lit for lit in _string_literals(formula))


def _has_equality_literal(formula: str) -> bool:
    return re.search(r'[A-Za-z]{1,3}\$?[0-9]+\s*=\s*"', formula) is not None


def classify_transformation(formula: str, functions: set[str]) -> str:
    if "DATEDIF" in functions:
        return AGE_DATEDIF
    if ("COUNTIF" in functions or "COUNTIFS" in functions) and _has_wildcard_literal(formula):
        return PATTERN_FLAG
    if "IFS" in functions or ("IF" in functions and _has_equality_literal(formula)):
        return EXACT_MAP
    if not functions and "&" in formula:
        return CONCAT
    return UNKNOWN


# ---------------------------------------------------------------------------
# Modelo del análisis
# ---------------------------------------------------------------------------


@dataclass
class ColumnInfo:
    index: int
    letter: str
    header: str
    kind: str  # raw | derived
    formula_count: int
    local_source_indices: set[int] = field(default_factory=set)
    directly_used: bool = False
    transitively_used: bool = False
    gap_rows: list[int] = field(default_factory=list)
    pattern_count: int = 0
    example_formula: str = ""
    functions: set[str] = field(default_factory=set)
    normalized_forms: set[str] = field(default_factory=set)


@dataclass
class Transformation:
    target_column: str
    target_header: str
    transformation_type: str
    source_columns: list[str]
    source_headers: list[str]
    example_formula: str
    formula_count: int
    pattern_count: int


@dataclass
class RuleCandidate:
    target_column: str
    target_header: str
    rule_index: int
    source_column: str
    source_header: str
    operator: str
    criterion: str
    output_literal: str
    appended_columns: str
    original_formula: str


@dataclass
class InvariantWarning:
    check: str
    severity: str
    target_column: str
    target_header: str
    detail: str


@dataclass
class CurrentLogicResult:
    output_dir: Path
    doc_path: Path
    files: list[Path]
    detail_sheet: str
    detail_sheet_present: bool
    output_sheets: list[str]
    columns: list[ColumnInfo]
    transformations: list[Transformation]
    rule_candidates: list[RuleCandidate]
    transitive_rows: list[list]
    transitive_raw_columns: list[tuple[str, str]]
    invariant_warnings: list[InvariantWarning]


# ---------------------------------------------------------------------------
# Núcleo
# ---------------------------------------------------------------------------


def _expand_columns(col_start: str | None, col_end: str | None) -> list[int]:
    if not col_start or not col_end:
        return []
    start, end = col_to_index(col_start), col_to_index(col_end)
    if end - start + 1 > _MAX_COLUMN_SPAN:
        return [start, end]
    return list(range(start, end + 1))


def _extract_rules(
    column_letter: str,
    header: str,
    transformation_type: str,
    example_formula: str,
    appended_columns: list[str],
    headers: dict[int, str],
) -> list[RuleCandidate]:
    rules: list[RuleCandidate] = []
    appended = "|".join(appended_columns)

    if transformation_type == EXACT_MAP:
        contents = _call_contents(example_formula, "IFS")
        if not contents:
            contents = _call_contents(example_formula, "IF")
        if contents:
            args = [a.strip() for a in _split_top_level(contents[0])]
            for pair_start in range(0, len(args) - 1, 2):
                condition = args[pair_start]
                value = args[pair_start + 1]
                match = _IFS_COND_RE.match(condition)
                if match:
                    source_column = _column_of_ref(match.group(1)) or ""
                    operator = "equals" if match.group(2) == "=" else "not_equals"
                    criterion = _unquote(match.group(3))
                else:
                    source_column = ""
                    operator = "unparsed"
                    criterion = condition
                rules.append(
                    RuleCandidate(
                        target_column=column_letter,
                        target_header=header,
                        rule_index=len(rules) + 1,
                        source_column=source_column,
                        source_header=headers.get(col_to_index(source_column), "")
                        if source_column
                        else "",
                        operator=operator,
                        criterion=criterion,
                        output_literal=_unquote(value),
                        appended_columns=appended,
                        original_formula=example_formula,
                    )
                )

    elif transformation_type == PATTERN_FLAG:
        for content in _call_contents(example_formula, "COUNTIF"):
            args = [a.strip() for a in _split_top_level(content)]
            if len(args) != 2:
                continue
            source_column = _column_of_ref(args[0]) or ""
            literal = _unquote(args[1])
            if len(literal) >= 2 and literal.startswith("*") and literal.endswith("*"):
                operator, criterion = "contains", literal[1:-1]
            elif "*" in literal or "?" in literal:
                operator, criterion = "wildcard_match", literal
            else:
                operator, criterion = "equals", literal
            rules.append(
                RuleCandidate(
                    target_column=column_letter,
                    target_header=header,
                    rule_index=len(rules) + 1,
                    source_column=source_column,
                    source_header=headers.get(col_to_index(source_column), "")
                    if source_column
                    else "",
                    operator=operator,
                    criterion=criterion,
                    output_literal="",
                    appended_columns=appended,
                    original_formula=example_formula,
                )
            )

    return rules


def analyze_current_logic(
    input_path: Path,
    output_dir: Path,
    *,
    detail_sheet: str = DETAIL_SHEET_DEFAULT,
    doc_path: Path | None = None,
    output_sheets: list[str] | None = None,
) -> CurrentLogicResult:
    doc_path = doc_path or Path(DEFAULT_DOC)
    sha256 = compute_sha256(input_path) if input_path.is_file() else ""

    wb = load_workbook_safely(input_path)
    try:
        sheet_names = list(wb.sheetnames)
        cells = collect_formula_cells(wb)
        detail_present = detail_sheet in sheet_names
        headers: dict[int, str] = {}
        max_column = 0
        data_rows = 0
        detail_values_present: dict[tuple[int, int], bool] = {}
        if detail_present:
            ws = wb[detail_sheet]
            max_column = ws.max_column or 0
            data_rows = max((ws.max_row or 1) - 1, 0)
            for col_index in range(1, max_column + 1):
                raw = ws.cell(row=1, column=col_index).value
                headers[col_index] = "" if raw is None else str(raw).strip()
            # Presencia (booleana) de datos por celda: NO se guardan valores.
            for row in ws.iter_rows(min_row=2):
                for cell in row:
                    if cell.value is not None:
                        detail_values_present[(cell.row, cell.column)] = True
    finally:
        wb.close()

    if output_sheets is None:
        sheets_with_formulas = {c.sheet for c in cells}
        output_sheets = [
            name
            for name in sheet_names
            if name != detail_sheet and name in sheets_with_formulas
        ]
    output_set = set(output_sheets)

    # --- fórmulas de la hoja de detalle agrupadas por columna ---------------
    detail_cells_by_col: dict[int, list] = defaultdict(list)
    for cell in cells:
        if cell.sheet == detail_sheet:
            detail_cells_by_col[cell.col_index].append(cell)

    # --- columnas referenciadas directamente por outputs -------------------
    outputs_direct: set[int] = set()
    for cell in cells:
        if cell.sheet not in output_set:
            continue
        for ref in cell.references:
            if ref.sheet == detail_sheet:
                outputs_direct.update(_expand_columns(ref.col_start, ref.col_end))

    # --- fuentes locales de cada columna derivada -------------------------
    derived_sources: dict[int, set[int]] = {}
    for col_index, members in detail_cells_by_col.items():
        sources: set[int] = set()
        for member in members:
            for ref in member.references:
                if ref.sheet is None:
                    sources.update(_expand_columns(ref.col_start, ref.col_end))
        sources.discard(col_index)
        derived_sources[col_index] = sources

    def is_derived(col_index: int) -> bool:
        return col_index in detail_cells_by_col

    # --- cierre transitivo hacia atrás desde lo que usan los outputs -------
    used: set[int] = set(outputs_direct)
    stack = list(outputs_direct)
    while stack:
        current = stack.pop()
        for source in derived_sources.get(current, ()):
            if source not in used:
                used.add(source)
                stack.append(source)

    # --- construir ColumnInfo por cada columna A..max ---------------------
    total_columns = max(max_column, *detail_cells_by_col) if detail_cells_by_col else max_column
    columns: list[ColumnInfo] = []
    for col_index in range(1, total_columns + 1):
        members = detail_cells_by_col.get(col_index, [])
        kind = "derived" if members else "raw"
        gap_rows: list[int] = []
        pattern_count = 0
        example_formula = ""
        functions: set[str] = set()
        normalized_forms: set[str] = set()
        if members:
            ordered = sorted(members, key=lambda c: c.row)
            example_formula = ordered[0].formula
            functions = {fn for m in members for fn in m.functions}
            normalized_forms = {m.normalized for m in members}
            pattern_count = len(normalized_forms)
            present_rows = {m.row for m in members}
            for row_index in range(2, data_rows + 2):
                if row_index in present_rows:
                    continue
                row_has_data = any(
                    detail_values_present.get((row_index, other))
                    for other in range(1, total_columns + 1)
                    if other != col_index
                )
                if row_has_data:
                    gap_rows.append(row_index)
        columns.append(
            ColumnInfo(
                index=col_index,
                letter=index_to_col(col_index),
                header=headers.get(col_index, ""),
                kind=kind,
                formula_count=len(members),
                local_source_indices=derived_sources.get(col_index, set()),
                directly_used=col_index in outputs_direct,
                transitively_used=col_index in used,
                gap_rows=gap_rows,
                pattern_count=pattern_count,
                example_formula=example_formula,
                functions=functions,
                normalized_forms=normalized_forms,
            )
        )

    by_index = {c.index: c for c in columns}

    # --- Parte B: transformaciones ---------------------------------------
    transformations: list[Transformation] = []
    for column in columns:
        if column.kind != "derived":
            continue
        source_indices = sorted(column.local_source_indices)
        transformation_type = classify_transformation(column.example_formula, column.functions)
        transformations.append(
            Transformation(
                target_column=column.letter,
                target_header=column.header,
                transformation_type=transformation_type,
                source_columns=[index_to_col(i) for i in source_indices],
                source_headers=[by_index[i].header if i in by_index else "" for i in source_indices],
                example_formula=column.example_formula,
                formula_count=column.formula_count,
                pattern_count=column.pattern_count,
            )
        )
    transformation_type_by_col = {t.target_column: t.transformation_type for t in transformations}

    # --- Parte C: reglas candidatas ------------------------------------
    rule_candidates: list[RuleCandidate] = []
    for transformation in transformations:
        if transformation.transformation_type not in _RULE_TYPES:
            continue
        column = by_index[col_to_index(transformation.target_column)]
        local_refs = set(column.local_source_indices)
        rules = _extract_rules(
            transformation.target_column,
            transformation.target_header,
            transformation.transformation_type,
            transformation.example_formula,
            appended_columns=[],  # se completa abajo
            headers=headers,
        )
        rule_source_cols = {
            col_to_index(rule.source_column) for rule in rules if rule.source_column
        }
        appended = sorted(
            index_to_col(i) for i in local_refs - rule_source_cols
        )
        for rule in rules:
            rule.appended_columns = "|".join(appended)
        rule_candidates.extend(rules)

    # --- Parte D: dependencias transitivas hasta columnas raw -------------
    def paths_to_raw(start: int) -> dict[int, set[str]]:
        results: dict[int, set[str]] = defaultdict(set)

        def walk(node: int, trail: list[str]) -> None:
            if not is_derived(node):
                results[node].add("->".join(trail))
                return
            for source in sorted(derived_sources.get(node, ())):
                letter = index_to_col(source)
                if letter in trail:
                    continue
                walk(source, [*trail, letter])

        walk(start, [index_to_col(start)])
        return results

    transitive_rows: list[list] = []
    for sheet in output_sheets:
        direct_cols: set[int] = set()
        for cell in cells:
            if cell.sheet != sheet:
                continue
            for ref in cell.references:
                if ref.sheet == detail_sheet:
                    direct_cols.update(_expand_columns(ref.col_start, ref.col_end))
        aggregated: dict[int, set[str]] = defaultdict(set)
        for col_index in sorted(direct_cols):
            for raw_index, path_set in paths_to_raw(col_index).items():
                aggregated[raw_index] |= path_set
        for raw_index in sorted(aggregated):
            transitive_rows.append(
                [
                    sheet,
                    index_to_col(raw_index),
                    headers.get(raw_index, ""),
                    "|".join(sorted(aggregated[raw_index])),
                ]
            )

    transitive_raw_columns = sorted(
        (
            (index_to_col(i), headers.get(i, ""))
            for i in used
            if not is_derived(i)
        ),
        key=lambda pair: col_to_index(pair[0]),
    )

    # --- Invariantes / advertencias -----------------------------------
    invariant_warnings: list[InvariantWarning] = []
    for column in columns:
        if column.kind != "derived":
            continue
        if data_rows and column.formula_count != data_rows:
            invariant_warnings.append(
                InvariantWarning(
                    check="formula_count_vs_data_rows",
                    severity="warning",
                    target_column=column.letter,
                    target_header=column.header,
                    detail=(
                        f"formula_count={column.formula_count} != filas de datos={data_rows}"
                    ),
                )
            )
        if column.pattern_count > 1:
            invariant_warnings.append(
                InvariantWarning(
                    check="pattern_count",
                    severity="warning",
                    target_column=column.letter,
                    target_header=column.header,
                    detail=f"pattern_count={column.pattern_count} (>1): la columna no es homogénea",
                )
            )
        if column.gap_rows:
            sample = ", ".join(str(r) for r in column.gap_rows[:5])
            more = "" if len(column.gap_rows) <= 5 else f" (+{len(column.gap_rows) - 5} más)"
            invariant_warnings.append(
                InvariantWarning(
                    check="gaps",
                    severity="warning",
                    target_column=column.letter,
                    target_header=column.header,
                    detail=f"{len(column.gap_rows)} fila(s) con datos pero sin fórmula: {sample}{more}",
                )
            )

    # --- escritura -----------------------------------------------------
    output_dir.mkdir(parents=True, exist_ok=True)
    files = _write_outputs(
        output_dir=output_dir,
        doc_path=doc_path,
        input_name=input_path.name,
        sha256=sha256,
        detail_sheet=detail_sheet,
        detail_present=detail_present,
        output_sheets=output_sheets,
        columns=columns,
        transformations=transformations,
        rule_candidates=rule_candidates,
        transitive_rows=transitive_rows,
        transitive_raw_columns=transitive_raw_columns,
        invariant_warnings=invariant_warnings,
        transformation_type_by_col=transformation_type_by_col,
        data_rows=data_rows,
    )

    return CurrentLogicResult(
        output_dir=output_dir,
        doc_path=doc_path,
        files=files,
        detail_sheet=detail_sheet,
        detail_sheet_present=detail_present,
        output_sheets=output_sheets,
        columns=columns,
        transformations=transformations,
        rule_candidates=rule_candidates,
        transitive_rows=transitive_rows,
        transitive_raw_columns=transitive_raw_columns,
        invariant_warnings=invariant_warnings,
    )


# ---------------------------------------------------------------------------
# Escritura de CSV + documento
# ---------------------------------------------------------------------------

SOURCE_SCHEMA_HEADER = [
    "column",
    "header",
    "kind",
    "has_formulas",
    "formula_count",
    "directly_used_by_outputs",
    "transitively_used_by_outputs",
]
DERIVED_TRANSFORMATIONS_HEADER = [
    "target_column",
    "target_header",
    "transformation_type",
    "source_columns",
    "source_headers",
    "example_formula",
    "formula_count",
    "pattern_count",
]
RULE_CANDIDATES_HEADER = [
    "target_column",
    "target_header",
    "rule_index",
    "source_column",
    "source_header",
    "operator",
    "criterion",
    "output_literal",
    "appended_columns",
    "original_formula",
]
TRANSITIVE_HEADER = [
    "target_sheet",
    "raw_source_column",
    "raw_source_header",
    "dependency_paths",
]
INVARIANTS_HEADER = ["check", "severity", "target_column", "target_header", "detail"]


def _write_csv(path: Path, header: list[str], rows) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


_README = """# Current logic inventory (Sprint 1.3)

Salida de `scripts/inventory_current_logic.py`. Describe la lógica **actual** de la
hoja de detalle tal como está escrita en las fórmulas. No traduce a reglas REMASEP
ni valida corrección; no infiere nada que no esté en una fórmula. No contiene
valores de filas de pacientes: solo encabezados, fórmulas, coordenadas y metadatos.

## `source_schema.csv`
Una fila por columna de la hoja de detalle.
`kind` = `derived` si la columna contiene fórmulas, `raw` en caso contrario.
`directly_used_by_outputs` = alguna hoja output la referencia directamente.
`transitively_used_by_outputs` = la usan los outputs directamente o a través de
columnas derivadas (cierre hacia atrás por las columnas fuente locales).

## `derived_transformations.csv`
Una fila por columna derivada. `transformation_type` es técnico, no clínico:
`CONCAT` (solo `&` de referencias), `AGE_DATEDIF` (usa `DATEDIF`),
`EXACT_MAP` (`IFS`/`IF` con igualdad a literales), `PATTERN_FLAG` (`COUNTIF` con
comodines), `UNKNOWN`. `pattern_count` > 1 indica fórmulas no homogéneas.

## `derived_rule_candidates.csv`
Una fila por condición observada en las columnas que implementan clasificación
(`EXACT_MAP` y `PATTERN_FLAG`). `operator` ∈ {`equals`, `not_equals`, `contains`,
`wildcard_match`}. Para `COUNTIF("*texto*")` el `criterion` es `texto` sin los
asteriscos exteriores; `original_formula` conserva la fórmula completa.
`appended_columns` = columnas concatenadas al resultado (p.ej. `&H2`).
**No hay interpretación semántica médica.**

## `transitive_output_dependencies.csv`
Qué columnas **raw** alimentan finalmente cada hoja output, siguiendo las
columnas derivadas. `dependency_paths` lista rutas `derivada->...->raw`
separadas por `|`. Las columnas derivadas no aparecen como raw.

## `invariants.csv`
Advertencias de análisis (nunca detienen el proceso): `formula_count_vs_data_rows`,
`pattern_count` (>1), `gaps` (filas con datos pero sin fórmula en una columna
derivada). Si está vacío (solo cabecera) no se detectaron problemas.
"""


def _kv(value: bool) -> str:
    return "True" if value else "False"


def _write_outputs(
    *,
    output_dir: Path,
    doc_path: Path,
    input_name: str,
    sha256: str,
    detail_sheet: str,
    detail_present: bool,
    output_sheets: list[str],
    columns: list[ColumnInfo],
    transformations: list[Transformation],
    rule_candidates: list[RuleCandidate],
    transitive_rows: list[list],
    transitive_raw_columns: list[tuple[str, str]],
    invariant_warnings: list[InvariantWarning],
    transformation_type_by_col: dict[str, str],
    data_rows: int,
) -> list[Path]:
    schema_path = output_dir / "source_schema.csv"
    _write_csv(
        schema_path,
        SOURCE_SCHEMA_HEADER,
        (
            [
                c.letter,
                c.header,
                c.kind,
                _kv(c.kind == "derived"),
                c.formula_count,
                _kv(c.directly_used),
                _kv(c.transitively_used),
            ]
            for c in columns
        ),
    )

    transformations_path = output_dir / "derived_transformations.csv"
    _write_csv(
        transformations_path,
        DERIVED_TRANSFORMATIONS_HEADER,
        (
            [
                t.target_column,
                t.target_header,
                t.transformation_type,
                "|".join(t.source_columns),
                "|".join(t.source_headers),
                t.example_formula,
                t.formula_count,
                t.pattern_count,
            ]
            for t in transformations
        ),
    )

    rules_path = output_dir / "derived_rule_candidates.csv"
    _write_csv(
        rules_path,
        RULE_CANDIDATES_HEADER,
        (
            [
                r.target_column,
                r.target_header,
                r.rule_index,
                r.source_column,
                r.source_header,
                r.operator,
                r.criterion,
                r.output_literal,
                r.appended_columns,
                r.original_formula,
            ]
            for r in rule_candidates
        ),
    )

    transitive_path = output_dir / "transitive_output_dependencies.csv"
    _write_csv(transitive_path, TRANSITIVE_HEADER, transitive_rows)

    invariants_path = output_dir / "invariants.csv"
    _write_csv(
        invariants_path,
        INVARIANTS_HEADER,
        (
            [w.check, w.severity, w.target_column, w.target_header, w.detail]
            for w in invariant_warnings
        ),
    )

    readme_path = output_dir / "README.md"
    readme_path.write_text(_README, encoding="utf-8")

    doc_path.parent.mkdir(parents=True, exist_ok=True)
    doc_path.write_text(
        render_current_logic_doc(
            input_name=input_name,
            sha256=sha256,
            detail_sheet=detail_sheet,
            detail_present=detail_present,
            output_sheets=output_sheets,
            columns=columns,
            transformations=transformations,
            rule_candidates=rule_candidates,
            transitive_rows=transitive_rows,
            transitive_raw_columns=transitive_raw_columns,
            invariant_warnings=invariant_warnings,
            data_rows=data_rows,
        ),
        encoding="utf-8",
    )

    return [
        schema_path,
        transformations_path,
        rules_path,
        transitive_path,
        invariants_path,
        readme_path,
        doc_path,
    ]


def _c(value: object) -> str:
    """Escapa `|` para no romper las tablas markdown (los CSV lo mantienen)."""
    return str(value).replace("|", "\\|")


def _find_columns_by_header(columns: list[ColumnInfo], *keywords: str) -> list[ColumnInfo]:
    wanted = [_norm_header(k) for k in keywords]
    return [c for c in columns if any(w in _norm_header(c.header) for w in wanted)]


def render_current_logic_doc(
    *,
    input_name: str,
    sha256: str,
    detail_sheet: str,
    detail_present: bool,
    output_sheets: list[str],
    columns: list[ColumnInfo],
    transformations: list[Transformation],
    rule_candidates: list[RuleCandidate],
    transitive_rows: list[list],
    transitive_raw_columns: list[tuple[str, str]],
    invariant_warnings: list[InvariantWarning],
    data_rows: int,
) -> str:
    lines: list[str] = []
    lines.append("# Lógica actual de la hoja de detalle")
    lines.append("")
    lines.append(
        "> Generado por `scripts/inventory_current_logic.py`. Describe **solo hechos "
        "observables** en las fórmulas de la hoja de detalle. No traduce a reglas "
        "REMASEP, no juzga corrección clínica y no infiere lógica no escrita."
    )
    lines.append(">")
    lines.append(f"> Archivo analizado: `{input_name}`  ")
    lines.append(f"> SHA256: `{sha256}`  ")
    lines.append(f"> Generado: {datetime.now(UTC).strftime('%Y-%m-%dT%H:%M:%SZ')}  ")
    lines.append(f"> Hoja de detalle: `{detail_sheet}` ({data_rows} filas físicas)  ")
    lines.append(
        f"> Hojas output: {', '.join(f'`{s}`' for s in output_sheets) or '—'}"
    )
    lines.append("")
    lines.append(
        "> **Sobre los conteos:** las cifras de filas y fórmulas de este documento "
        "corresponden a **filas físicas** de la hoja y **no deben interpretarse como "
        "la cantidad de atenciones reales**. El workbook de referencia arrastra sus "
        "fórmulas `AC:AL` más allá de las atenciones, por lo que puede haber filas "
        "físicas sin datos. La detección de filas estructuralmente vacías "
        "(`structural_empty_rows`) pertenece al análisis Medinet "
        "(`remasep.services.medinet_analysis`); para el recuento real del dataset "
        "consultar [`docs/MEDINET_ANALYSIS.md`](MEDINET_ANALYSIS.md)."
    )
    lines.append("")

    if not detail_present:
        lines.append(f"**AVISO:** la hoja `{detail_sheet}` no existe en el workbook.")
        lines.append("")
        return "\n".join(lines) + "\n"

    # --- Esquema ---
    lines.append("## Esquema de columnas")
    lines.append("")
    lines.append("| col | header | kind | fórmulas | uso directo | uso transitivo |")
    lines.append("| --- | --- | --- | ---: | :---: | :---: |")
    for c in columns:
        lines.append(
            f"| {c.letter} | {c.header} | {c.kind} | {c.formula_count} | "
            f"{'sí' if c.directly_used else '·'} | {'sí' if c.transitively_used else '·'} |"
        )
    lines.append("")

    # --- Transformaciones ---
    lines.append("## Transformaciones derivadas")
    lines.append("")
    lines.append("| col | header | tipo | fuentes | headers fuente | fórmulas | patrones |")
    lines.append("| --- | --- | --- | --- | --- | ---: | ---: |")
    for t in transformations:
        lines.append(
            f"| {t.target_column} | {t.target_header} | {t.transformation_type} | "
            f"{_c('|'.join(t.source_columns))} | {_c('|'.join(t.source_headers))} | "
            f"{t.formula_count} | {t.pattern_count} |"
        )
    lines.append("")
    lines.append(
        "Tipos técnicos: `CONCAT`, `AGE_DATEDIF`, `EXACT_MAP`, `PATTERN_FLAG`, "
        "`UNKNOWN`. No son categorías clínicas."
    )
    lines.append("")

    # --- Reglas candidatas ---
    lines.append("## Reglas candidatas")
    lines.append("")
    lines.append(
        "Una fila por condición observada en las columnas de clasificación "
        "(`EXACT_MAP`/`PATTERN_FLAG`). Detalle completo en "
        "`artifacts/current_logic/derived_rule_candidates.csv`."
    )
    lines.append("")
    per_column: dict[str, list[RuleCandidate]] = defaultdict(list)
    for rule in rule_candidates:
        per_column[rule.target_column].append(rule)
    lines.append("| col | header | nº reglas | operadores |")
    lines.append("| --- | --- | ---: | --- |")
    for column_letter, rules in per_column.items():
        operators = _c("|".join(sorted({r.operator for r in rules})))
        lines.append(
            f"| {column_letter} | {rules[0].target_header} | {len(rules)} | {operators} |"
        )
    if not per_column:
        lines.append("| — | — | 0 | — |")
    lines.append("")

    # --- Transitivas ---
    lines.append("## Columnas raw que alimentan cada output (transitivo)")
    lines.append("")
    lines.append("| output | raw | header | rutas |")
    lines.append("| --- | --- | --- | --- |")
    for target_sheet, raw_column, raw_header, paths in transitive_rows:
        lines.append(f"| {target_sheet} | {raw_column} | {raw_header} | {_c(paths)} |")
    if not transitive_rows:
        lines.append("| — | — | — | — |")
    lines.append("")
    lines.append("### Conjunto total de columnas raw utilizadas transitivamente")
    lines.append("")
    if transitive_raw_columns:
        lines.append(
            ", ".join(f"`{letter}` ({header})" for letter, header in transitive_raw_columns)
        )
    else:
        lines.append("Ninguna.")
    lines.append("")

    # --- Invariantes ---
    lines.append("## Invariantes / advertencias de análisis")
    lines.append("")
    if invariant_warnings:
        lines.append("| check | col | header | detalle |")
        lines.append("| --- | --- | --- | --- |")
        for w in invariant_warnings:
            lines.append(
                f"| {w.check} | {w.target_column} | {_c(w.target_header)} | {_c(w.detail)} |"
            )
    else:
        lines.append("Sin advertencias: cada columna derivada tiene una única fórmula")
        lines.append(
            f"normalizada y un `formula_count` igual a las {data_rows} filas físicas de "
            "la hoja (no atenciones reales; ver nota inicial), sin huecos."
        )
    lines.append("")

    # --- Observaciones que requieren validación funcional ---
    lines.append("## Observaciones que requieren validación funcional")
    lines.append("")
    lines.append(
        "Hechos observados en headers y dependencias. **No** se afirma que sean "
        "errores; requieren confirmación del responsable funcional."
    )
    lines.append("")

    def _describe(keyword: str, *aliases: str) -> str:
        found = _find_columns_by_header(columns, keyword, *aliases)
        if not found:
            return f"no se encontró ninguna columna cuyo encabezado contenga «{keyword}»"
        parts = []
        for column in found:
            estado = "participa" if column.transitively_used else "no participa"
            parts.append(f"`{column.letter}` («{column.header}») {estado} en las dependencias")
        return "; ".join(parts)

    concat_cols = [t for t in transformations if t.transformation_type == CONCAT]
    first_concat = concat_cols[0] if concat_cols else None
    prestacion_rule_sources = sorted(
        {r.source_header for r in rule_candidates if r.operator in {"contains", "wildcard_match"}}
    )

    lines.append(f"1. **ESTADO** — {_describe('ESTADO')}.")
    if first_concat:
        lines.append(
            f"2. **MODALIDAD** — {_describe('MODALIDAD')}. La primera columna `CONCAT` "
            f"(`{first_concat.target_column}`, «{first_concat.target_header}») se "
            f"construye a partir de {', '.join(f'«{h}»' for h in first_concat.source_headers)}."
        )
    else:
        lines.append(f"2. **MODALIDAD** — {_describe('MODALIDAD')}.")
    lines.append(
        f"3. **PRESTACIÓN REALIZADA** — {_describe('PRESTACION REALIZADA')}. "
        f"Las reglas `PATTERN_FLAG` usan "
        f"{', '.join(f'«{h}»' for h in prestacion_rule_sources) or '—'}."
    )
    lines.append(
        "4. **Columnas raw utilizadas transitivamente:** "
        + (
            ", ".join(f"`{letter}` ({header})" for letter, header in transitive_raw_columns)
            or "ninguna"
        )
        + "."
    )
    lines.append("")
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="inventory_current_logic.py",
        description=(
            "Inventario estructurado de la lógica actual de la hoja de detalle "
            "(esquema, transformaciones, reglas candidatas y dependencias "
            "transitivas). Solo lectura, sin semántica REMASEP."
        ),
    )
    parser.add_argument("workbook", help="Ruta al archivo .xlsx/.xlsm a analizar.")
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"(def: {DEFAULT_OUTPUT})")
    parser.add_argument(
        "--detail-sheet", default=DETAIL_SHEET_DEFAULT, help=f"(def: {DETAIL_SHEET_DEFAULT!r})"
    )
    parser.add_argument("--doc", default=DEFAULT_DOC, help=f"(def: {DEFAULT_DOC})")
    parser.add_argument(
        "--output-sheets",
        default=None,
        help="Lista separada por comas; def: todas las hojas con fórmulas salvo la de detalle.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    input_path = Path(args.workbook)
    output_sheets = (
        [s.strip() for s in args.output_sheets.split(",") if s.strip()]
        if args.output_sheets
        else None
    )

    try:
        result = analyze_current_logic(
            input_path,
            Path(args.output),
            detail_sheet=args.detail_sheet,
            doc_path=Path(args.doc),
            output_sheets=output_sheets,
        )
    except InventoryError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    print(f"Workbook analizado: {input_path}")
    if not result.detail_sheet_present:
        print(
            f"AVISO: la hoja de detalle {result.detail_sheet!r} no existe; "
            "los CSV quedan prácticamente vacíos.",
            file=sys.stderr,
        )
    raw_list = ", ".join(f"{letter} ({header})" for letter, header in result.transitive_raw_columns)
    print(f"  columnas raw usadas transitivamente ({len(result.transitive_raw_columns)}): {raw_list or '—'}")
    rules_by_col: dict[str, int] = defaultdict(int)
    for rule in result.rule_candidates:
        rules_by_col[rule.target_column] += 1
    print(f"  reglas candidatas: {len(result.rule_candidates)} total")
    for column_letter, count in rules_by_col.items():
        print(f"    {column_letter}: {count}")
    print("  transformaciones:")
    for transformation in result.transformations:
        print(
            f"    {transformation.target_column} -> {transformation.transformation_type} "
            f"(fuentes {'|'.join(transformation.source_columns) or '—'})"
        )
    if result.invariant_warnings:
        print(f"  advertencias de invariantes: {len(result.invariant_warnings)}")
        for warning in result.invariant_warnings:
            print(f"    [{warning.check}] {warning.target_column}: {warning.detail}")
    else:
        print("  advertencias de invariantes: ninguna")
    print(f"Artefactos escritos en: {result.output_dir}")
    for path in result.files:
        print(f"  - {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
