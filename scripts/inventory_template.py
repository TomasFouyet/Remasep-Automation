"""Inventario de la plantilla oficial y alineación con el workbook generador.

Sprint 1.4. Solo lectura: NO escribe ni modifica ningún archivo, NO ejecuta
macros ni COM. Ubuntu/Linux + openpyxl (`keep_vba=True` para el `.xlsm`,
`data_only=False`).

No interpreta significado clínico ni inventa mappings: produce evidencia
estructural y *candidatos* de alineación.

Privacidad: del workbook generador solo se leen encabezados, fórmulas y etiquetas
estructurales de las hojas REMASEP; nunca filas de pacientes (la hoja de detalle
se excluye por completo).

Uso:

    python scripts/inventory_template.py \\
        "data/local/REMASEP 2026_V1.4.xlsm" \\
        --generator "data/local/GENERACION DATOS REMASEP.xlsx" \\
        --output artifacts/template_inventory \\
        --doc docs/TEMPLATE_ALIGNMENT.md
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import re
import sys
import unicodedata
import zipfile
from collections import Counter, defaultdict
from dataclasses import dataclass
from datetime import UTC, datetime
from pathlib import Path

from analyze_dependencies import (
    FORMULA_DEPENDENCIES_HEADER,
    SHEET_DEPENDENCIES_HEADER,
    build_formula_dependencies,
    build_sheet_dependencies,
    collect_formula_cells,
)
from inventory_workbook import InventoryError, compute_sha256, load_workbook_safely

DEFAULT_OUTPUT = "artifacts/template_inventory"
DEFAULT_DOC = "docs/TEMPLATE_ALIGNMENT.md"
DEFAULT_GENERATOR_DETAIL_SHEET = "Atenciones - Detalles de citas"

_VBA_ENTRY = "xl/vbaProject.bin"
_MIN_LABEL_LEN = 3
_MAX_AMBIGUOUS_PER_LABEL = 6
# Al recorrer una hoja se amplía su dimensión para cubrir data validations o
# merges cercanos, pero no rangos "columna/fila entera".
_ROW_EXPANSION_CAP = 50000
_COL_EXPANSION_CAP = 16384


# ---------------------------------------------------------------------------
# Utilidades
# ---------------------------------------------------------------------------


def _norm_label(text: object) -> str:
    """Normaliza una etiqueta para comparar: mayúsculas, sin acentos ni puntuación."""
    raw = unicodedata.normalize("NFD", str(text).strip().upper())
    raw = "".join(ch for ch in raw if unicodedata.category(ch) != "Mn")
    raw = re.sub(r"\s+", " ", raw)
    return re.sub(r"[^0-9A-Z ]+", "", raw).strip()


def _looks_like_code(text: str) -> bool:
    return bool(re.fullmatch(r"\d{4,9}", text.strip()))


def _keep_vba_for(path: Path) -> bool:
    return path.suffix.lower() in {".xlsm", ".xltm", ".xlsb"}


def _sha256_bytes(data: bytes) -> str:
    return hashlib.sha256(data).hexdigest()


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------


@dataclass
class SheetInfo:
    name: str
    state: str
    max_row: int
    max_column: int
    formula_count: int
    merged_ranges_count: int
    hidden_rows: int
    hidden_columns: int
    protected: bool
    data_validation_count: int
    comments_count: int
    text_cell_count: int
    unlocked_cell_count: int
    filled_cell_count: int


@dataclass
class CellRecord:
    sheet: str
    cell: str
    row: int
    column: int
    kind: str
    has_formula: bool
    formula: str
    locked: bool
    fill_type: str
    fill_rgb: str
    has_validation: bool
    is_merged: bool
    text_label: str


@dataclass
class InputCandidate:
    sheet: str
    cell: str
    locked: bool
    fill_rgb: str
    has_validation: bool
    label_context: str
    candidate_reasons: str


@dataclass
class AlignmentCandidate:
    generator_sheet: str
    generator_cell: str
    generator_label: str
    template_sheet: str
    template_cell: str
    template_label: str
    match_type: str
    confidence: float
    context: str


@dataclass
class TemplateResult:
    output_dir: Path
    doc_path: Path
    files: list[Path]
    template_name: str
    template_sha256: str
    sheets: list[SheetInfo]
    vba_present: bool
    vba_payload_bytes: int | None
    vba_payload_sha256: str | None
    named_ranges: list[tuple[str, str, str]]
    cell_records: list[CellRecord]
    input_candidates: list[InputCandidate]
    dependency_rows: list[list]
    sheet_dependency_rows: list[list]
    unsupported_counts: dict[str, int]
    fingerprint: dict
    alignment_candidates: list[AlignmentCandidate]
    generator_analyzed: bool
    fill_distribution: list[tuple[str, int]]


# ---------------------------------------------------------------------------
# Lectura de estilos / estructura de celdas
# ---------------------------------------------------------------------------


def _fill_info(cell) -> tuple[str, str]:
    fill = getattr(cell, "fill", None)
    pattern = getattr(fill, "patternType", None)
    if not pattern or pattern == "none":
        return "", ""
    color = getattr(fill, "fgColor", None)
    rgb = ""
    if color is not None:
        color_type = getattr(color, "type", None)
        if color_type == "rgb":
            value = color.rgb
            rgb = value if isinstance(value, str) else ""
        elif color_type == "theme":
            tint = getattr(color, "tint", 0.0) or 0.0
            rgb = f"theme:{color.theme}" + (f"/tint:{round(tint, 3)}" if tint else "")
        elif color_type == "indexed":
            rgb = f"indexed:{color.indexed}"
    return pattern, rgb


def _is_locked(cell) -> bool:
    protection = getattr(cell, "protection", None)
    locked = getattr(protection, "locked", None)
    return True if locked is None else bool(locked)


def _merged_lookup(ws):
    ranges = list(ws.merged_cells.ranges)

    def in_merged(row: int, col: int) -> bool:
        return any(
            rng.min_row <= row <= rng.max_row and rng.min_col <= col <= rng.max_col
            for rng in ranges
        )

    return in_merged


def _validation_lookup(ws):
    boxes: list[tuple[int, int, int, int]] = []
    for dv in ws.data_validations.dataValidation:
        for rng in dv.sqref.ranges:
            boxes.append((rng.min_row, rng.max_row, rng.min_col, rng.max_col))

    def has_validation(row: int, col: int) -> bool:
        return any(r0 <= row <= r1 and c0 <= col <= c1 for r0, r1, c0, c1 in boxes)

    return has_validation


def _hidden_counts(ws) -> tuple[int, int]:
    hidden_rows = sum(
        1 for dim in ws.row_dimensions.values() if getattr(dim, "hidden", False)
    )
    hidden_columns = 0
    for dim in ws.column_dimensions.values():
        if getattr(dim, "hidden", False):
            span = (getattr(dim, "max", 0) or 0) - (getattr(dim, "min", 0) or 0) + 1
            hidden_columns += max(span, 1)
    return hidden_rows, hidden_columns


def _nearby_label(ws, row: int, col: int, text_cells: dict[tuple[int, int], str]) -> str:
    left = ""
    for c in range(col - 1, 0, -1):
        if (row, c) in text_cells:
            left = text_cells[(row, c)]
            break
    above = ""
    for r in range(row - 1, 0, -1):
        if (r, col) in text_cells:
            above = text_cells[(r, col)]
            break
    parts = []
    if left:
        parts.append(f"left={left!r}")
    if above:
        parts.append(f"above={above!r}")
    return " ".join(parts)


# ---------------------------------------------------------------------------
# Parte A + B + C: recorrer la plantilla
# ---------------------------------------------------------------------------


def scan_template(wb) -> tuple[
    list[SheetInfo],
    list[CellRecord],
    list[InputCandidate],
    dict[tuple[str, int], str],
]:
    sheets: list[SheetInfo] = []
    records: list[CellRecord] = []
    candidates: list[InputCandidate] = []
    sentinel_pool: dict[tuple[str, int], str] = {}

    for ws in wb.worksheets:
        in_merged = _merged_lookup(ws)
        has_validation = _validation_lookup(ws)
        hidden_rows, hidden_columns = _hidden_counts(ws)
        protected = bool(getattr(ws.protection, "sheet", False))

        # Rango a recorrer: dimensión de la hoja ampliada para incluir data
        # validations / merges que la exceden ligeramente (se ignoran los
        # rangos "columna entera" tipo C1:C1048576).
        bound_row = ws.max_row or 1
        bound_col = ws.max_column or 1
        extra_ranges = [rng for dv in ws.data_validations.dataValidation for rng in dv.sqref.ranges]
        extra_ranges += list(ws.merged_cells.ranges)
        for rng in extra_ranges:
            if rng.max_row <= _ROW_EXPANSION_CAP:
                bound_row = max(bound_row, rng.max_row)
            if rng.max_col <= _COL_EXPANSION_CAP:
                bound_col = max(bound_col, rng.max_col)

        text_cells: dict[tuple[int, int], str] = {}
        formula_count = comments_count = unlocked_count = filled_count = 0
        pending_inputs: list[tuple[str, int, int, bool, str, bool, str]] = []

        for row in ws.iter_rows(min_row=1, max_row=bound_row, min_col=1, max_col=bound_col):
            for cell in row:
                value = cell.value
                is_formula = cell.data_type == "f"
                is_text = cell.data_type == "s" and isinstance(value, str) and value.strip()
                if getattr(cell, "comment", None) is not None:
                    comments_count += 1
                locked = _is_locked(cell)
                if not locked:
                    unlocked_count += 1
                fill_type, fill_rgb = _fill_info(cell)
                if fill_type:
                    filled_count += 1
                if is_text:
                    text_cells[(cell.row, cell.column)] = value.strip()

                cell_has_validation = has_validation(cell.row, cell.column)
                merged = in_merged(cell.row, cell.column)

                # Parte C: candidato a input (sin fórmula, no merged, desbloqueado en
                # hoja protegida y/o con data validation).
                if (
                    not is_formula
                    and not merged
                    and ((not locked and protected) or cell_has_validation)
                ):
                    pending_inputs.append(
                        (cell.coordinate, cell.row, cell.column, locked, fill_rgb,
                         cell_has_validation, fill_type)
                    )

                relevant = (
                    is_formula
                    or bool(is_text)
                    or not locked
                    or bool(fill_type)
                    or cell_has_validation
                )
                if not relevant:
                    continue
                if is_formula:
                    formula_count += 1
                    kind = "formula"
                elif is_text:
                    kind = "label"
                elif not locked:
                    kind = "unlocked"
                elif cell_has_validation:
                    kind = "validation"
                else:
                    kind = "styled"

                records.append(
                    CellRecord(
                        sheet=ws.title,
                        cell=cell.coordinate,
                        row=cell.row,
                        column=cell.column,
                        kind=kind,
                        has_formula=is_formula,
                        formula=value if is_formula and isinstance(value, str) else "",
                        locked=locked,
                        fill_type=fill_type,
                        fill_rgb=fill_rgb,
                        has_validation=cell_has_validation,
                        is_merged=merged,
                        text_label=value.strip() if is_text else "",
                    )
                )

        # Sentinelas: texto largo, estable, único en la hoja, filas bajas.
        seen_norm: Counter[str] = Counter(_norm_label(v) for v in text_cells.values())
        for (row_index, col_index), value in sorted(text_cells.items()):
            norm = _norm_label(value)
            if len(norm) >= 8 and not _looks_like_code(norm) and seen_norm[norm] == 1:
                sentinel_pool[(ws.title, len(sentinel_pool))] = json.dumps(
                    {"sheet": ws.title, "cell": f"{_col_letter(col_index)}{row_index}", "label": norm}
                )

        # Parte C: completar contexto de etiqueta con el mapa de texto ya completo.
        for coord, row_index, col_index, locked, fill_rgb, cell_has_validation, fill_type in (
            pending_inputs
        ):
            reasons = []
            if not locked:
                reasons.append("unlocked")
            if cell_has_validation:
                reasons.append("has_data_validation")
            if fill_type:
                reasons.append("non_default_fill")
            if protected:
                reasons.append("in_protected_sheet")
            candidates.append(
                InputCandidate(
                    sheet=ws.title,
                    cell=coord,
                    locked=locked,
                    fill_rgb=fill_rgb,
                    has_validation=cell_has_validation,
                    label_context=_nearby_label(ws, row_index, col_index, text_cells),
                    candidate_reasons="|".join(reasons),
                )
            )

        sheets.append(
            SheetInfo(
                name=ws.title,
                state=ws.sheet_state,
                max_row=ws.max_row or 0,
                max_column=ws.max_column or 0,
                formula_count=formula_count,
                merged_ranges_count=len(ws.merged_cells.ranges),
                hidden_rows=hidden_rows,
                hidden_columns=hidden_columns,
                protected=protected,
                data_validation_count=len(ws.data_validations.dataValidation),
                comments_count=comments_count,
                text_cell_count=len(text_cells),
                unlocked_cell_count=unlocked_count,
                filled_cell_count=filled_count,
            )
        )

    return sheets, records, candidates, sentinel_pool


def _col_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


# ---------------------------------------------------------------------------
# Parte E: fingerprint estructural
# ---------------------------------------------------------------------------


def build_fingerprint(
    wb, sheets: list[SheetInfo], cells, sentinel_pool: dict[tuple[str, int], str], full_sha: str
) -> dict:
    formula_coords: dict[str, list[str]] = defaultdict(list)
    for cell in cells:
        formula_coords[cell.sheet].append(cell.coord)

    key_formula_cells = {}
    for name, coords in formula_coords.items():
        ordered = sorted(coords, key=lambda c: (len(c), c))
        key_formula_cells[name] = {
            "count": len(ordered),
            "coords_sha256": _sha256_bytes(",".join(sorted(coords)).encode("utf-8")),
            "first": ordered[:10],
        }

    sentinels = sorted(
        (json.loads(v) for v in sentinel_pool.values()),
        key=lambda d: (d["sheet"], d["cell"]),
    )
    # Como máximo 8 sentinelas por hoja para mantener el fingerprint acotado.
    per_sheet: Counter[str] = Counter()
    trimmed_sentinels = []
    for entry in sentinels:
        if per_sheet[entry["sheet"]] >= 8:
            continue
        per_sheet[entry["sheet"]] += 1
        trimmed_sentinels.append(entry)

    structural = {
        "sheet_names_ordered": list(wb.sheetnames),
        "sheet_states": {s.name: s.state for s in sheets},
        "sheet_dimensions": {s.name: [s.max_row, s.max_column] for s in sheets},
        "formula_count_by_sheet": {s.name: s.formula_count for s in sheets},
        "merged_ranges_by_sheet": {s.name: s.merged_ranges_count for s in sheets},
        "key_formula_cells": key_formula_cells,
        "sentinel_labels": trimmed_sentinels,
    }
    canonical = json.dumps(structural, sort_keys=True, ensure_ascii=False)
    return {
        "full_file_sha256": full_sha,
        "structural_sha256": _sha256_bytes(canonical.encode("utf-8")),
        "structural": structural,
        "note": "Fingerprint experimental. No usar en producción todavía.",
    }


# ---------------------------------------------------------------------------
# Parte F: alineación con el workbook generador
# ---------------------------------------------------------------------------


def _collect_labels(wb, sheet_names: list[str]) -> list[tuple[str, str, str]]:
    """(sheet, cell, texto) de celdas de texto no fórmula en las hojas indicadas."""
    labels: list[tuple[str, str, str]] = []
    for name in sheet_names:
        if name not in wb.sheetnames:
            continue
        ws = wb[name]
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "s" and isinstance(cell.value, str) and cell.value.strip():
                    labels.append((name, cell.coordinate, cell.value.strip()))
    return labels


def build_alignment_candidates(
    generator_labels: list[tuple[str, str, str]],
    template_labels: list[tuple[str, str, str]],
) -> list[AlignmentCandidate]:
    template_by_exact: dict[str, list[tuple[str, str, str]]] = defaultdict(list)
    template_by_norm: dict[str, list[tuple[str, str, str]]] = defaultdict(list)
    template_codes: dict[str, list[tuple[str, str, str]]] = defaultdict(list)
    for sheet, cell, text in template_labels:
        template_by_exact[text].append((sheet, cell, text))
        template_by_norm[_norm_label(text)].append((sheet, cell, text))
        for code in re.findall(r"\d{4,9}", text):
            template_codes[code].append((sheet, cell, text))

    candidates: list[AlignmentCandidate] = []
    for gen_sheet, gen_cell, gen_text in generator_labels:
        norm = _norm_label(gen_text)
        # Etiquetas demasiado cortas o números sueltos generan ruido; se ignoran
        # salvo que sean un código de prestación.
        if len(norm) < _MIN_LABEL_LEN and not _looks_like_code(gen_text):
            continue
        if norm.isdigit() and not _looks_like_code(gen_text):
            continue

        exact = template_by_exact.get(gen_text, [])
        norm_matches = template_by_norm.get(norm, []) if norm else []

        if len(exact) == 1:
            hits, match_type, confidence = exact, "exact_label", 1.0
        elif len(exact) > 1:
            hits, match_type, confidence = exact, "ambiguous", 0.4
        elif len(norm_matches) == 1:
            hits, match_type, confidence = norm_matches, "normalized_label", 0.85
        elif len(norm_matches) > 1:
            hits, match_type, confidence = norm_matches, "ambiguous", 0.4
        elif _looks_like_code(gen_text) and gen_text in template_codes:
            hits = template_codes[gen_text]
            match_type = "contextual" if len(hits) == 1 else "ambiguous"
            confidence = 0.7 if len(hits) == 1 else 0.35
        else:
            continue

        total_hits = len(hits)
        ordered_hits = sorted(hits, key=lambda h: (h[0], h[1]))
        if match_type == "ambiguous":
            ordered_hits = ordered_hits[:_MAX_AMBIGUOUS_PER_LABEL]

        for tmpl_sheet, tmpl_cell, tmpl_text in ordered_hits:
            context = f"norm={norm!r}"
            if match_type == "ambiguous":
                context += f"; ambiguous_total={total_hits}"
            candidates.append(
                AlignmentCandidate(
                    generator_sheet=gen_sheet,
                    generator_cell=gen_cell,
                    generator_label=gen_text,
                    template_sheet=tmpl_sheet,
                    template_cell=tmpl_cell,
                    template_label=tmpl_text,
                    match_type=match_type,
                    confidence=round(confidence, 3),
                    context=context,
                )
            )

    candidates.sort(
        key=lambda c: (c.generator_sheet, c.generator_cell, -c.confidence, c.template_sheet, c.template_cell)
    )
    return candidates


# ---------------------------------------------------------------------------
# Escritura
# ---------------------------------------------------------------------------

TEMPLATE_CELLS_HEADER = [
    "sheet", "cell", "row", "column", "kind", "has_formula", "formula", "locked",
    "fill_type", "fill_rgb", "has_validation", "is_merged", "text_label",
]
INPUT_CANDIDATES_HEADER = [
    "sheet", "cell", "locked", "fill_rgb", "has_validation", "label_context",
    "candidate_reasons",
]
ALIGNMENT_HEADER = [
    "generator_sheet", "generator_cell", "generator_label", "template_sheet",
    "template_cell", "template_label", "match_type", "confidence", "context",
]
FILL_DISTRIBUTION_HEADER = ["fill_rgb", "cell_count"]


def _write_csv(path: Path, header: list[str], rows) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


def _bool(value: bool) -> str:
    return "True" if value else "False"


def _md(value: object) -> str:
    return str(value).replace("|", "\\|")


# ---------------------------------------------------------------------------
# Orquestación
# ---------------------------------------------------------------------------


def analyze_template(
    template_path: Path,
    output_dir: Path,
    *,
    generator_path: Path | None = None,
    doc_path: Path | None = None,
    generator_detail_sheet: str = DEFAULT_GENERATOR_DETAIL_SHEET,
    generator_label_sheets: list[str] | None = None,
) -> TemplateResult:
    doc_path = doc_path or Path(DEFAULT_DOC)
    full_sha = compute_sha256(template_path) if template_path.is_file() else ""

    wb = load_workbook_safely(template_path, keep_vba=_keep_vba_for(template_path))
    try:
        sheet_names = list(wb.sheetnames)
        sheets, cell_records, input_candidates, sentinel_pool = scan_template(wb)
        cells = collect_formula_cells(wb)
        named_ranges = [
            (name, defined.value or "", "workbook")
            for name, defined in wb.defined_names.items()
        ]
        fingerprint = build_fingerprint(wb, sheets, cells, sentinel_pool, full_sha)
        template_labels = _collect_labels(wb, sheet_names)
    finally:
        wb.close()

    sheet_order = {name: index for index, name in enumerate(sheet_names)}
    dependency_rows, _unknown = build_formula_dependencies(
        cells, sheet_order, set(sheet_names)
    )
    sheet_dependency_rows = build_sheet_dependencies(dependency_rows)

    unsupported_counts: Counter[str] = Counter()
    for cell in cells:
        for construct in cell.unsupported:
            unsupported_counts[construct] += 1

    fill_distribution = Counter(
        r.fill_rgb for r in cell_records if r.fill_type and r.fill_rgb
    )

    # --- VBA (sin ejecutar nada) -----------------------------------------
    vba_present = False
    vba_bytes: int | None = None
    vba_sha: str | None = None
    if zipfile.is_zipfile(template_path):
        with zipfile.ZipFile(template_path) as archive:
            if _VBA_ENTRY in archive.namelist():
                payload = archive.read(_VBA_ENTRY)
                vba_present = True
                vba_bytes = len(payload)
                vba_sha = _sha256_bytes(payload)

    # --- Parte F --------------------------------------------------------
    alignment_candidates: list[AlignmentCandidate] = []
    generator_analyzed = False
    if generator_path is not None:
        if generator_label_sheets is None:
            gen_wb_peek = load_workbook_safely(generator_path)
            try:
                generator_label_sheets = [
                    name for name in gen_wb_peek.sheetnames if name != generator_detail_sheet
                ]
            finally:
                gen_wb_peek.close()
        if generator_detail_sheet in generator_label_sheets:
            raise InventoryError(
                "La hoja de detalle no puede estar en generator_label_sheets "
                "(contiene datos de pacientes)."
            )
        gen_wb = load_workbook_safely(generator_path)
        try:
            generator_labels = _collect_labels(gen_wb, generator_label_sheets)
        finally:
            gen_wb.close()
        alignment_candidates = build_alignment_candidates(generator_labels, template_labels)
        generator_analyzed = True

    files = _write_outputs(
        output_dir=output_dir,
        doc_path=doc_path,
        template_path=template_path,
        full_sha=full_sha,
        sheets=sheets,
        cell_records=cell_records,
        input_candidates=input_candidates,
        dependency_rows=dependency_rows,
        sheet_dependency_rows=sheet_dependency_rows,
        unsupported_counts=dict(unsupported_counts),
        fingerprint=fingerprint,
        alignment_candidates=alignment_candidates,
        named_ranges=named_ranges,
        vba_present=vba_present,
        vba_bytes=vba_bytes,
        vba_sha=vba_sha,
        fill_distribution=sorted(fill_distribution.items(), key=lambda kv: (-kv[1], kv[0])),
        generator_analyzed=generator_analyzed,
    )

    return TemplateResult(
        output_dir=output_dir,
        doc_path=doc_path,
        files=files,
        template_name=template_path.name,
        template_sha256=full_sha,
        sheets=sheets,
        vba_present=vba_present,
        vba_payload_bytes=vba_bytes,
        vba_payload_sha256=vba_sha,
        named_ranges=named_ranges,
        cell_records=cell_records,
        input_candidates=input_candidates,
        dependency_rows=dependency_rows,
        sheet_dependency_rows=sheet_dependency_rows,
        unsupported_counts=dict(unsupported_counts),
        fingerprint=fingerprint,
        alignment_candidates=alignment_candidates,
        generator_analyzed=generator_analyzed,
        fill_distribution=sorted(fill_distribution.items(), key=lambda kv: (-kv[1], kv[0])),
    )


def _write_outputs(**kw) -> list[Path]:
    output_dir: Path = kw["output_dir"]
    doc_path: Path = kw["doc_path"]
    template_path: Path = kw["template_path"]
    sheets: list[SheetInfo] = kw["sheets"]
    cell_records: list[CellRecord] = kw["cell_records"]
    input_candidates: list[InputCandidate] = kw["input_candidates"]
    output_dir.mkdir(parents=True, exist_ok=True)

    summary = {
        "filename": template_path.name,
        "sha256": kw["full_sha"],
        "generated_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "sheet_names": [s.name for s in sheets],
        "sheet_states": {s.name: s.state for s in sheets},
        "totals": {
            "formula_count": sum(s.formula_count for s in sheets),
            "merged_ranges_count": sum(s.merged_ranges_count for s in sheets),
            "data_validation_count": sum(s.data_validation_count for s in sheets),
            "comments_count": sum(s.comments_count for s in sheets),
            "hidden_rows_count": sum(s.hidden_rows for s in sheets),
            "hidden_columns_count": sum(s.hidden_columns for s in sheets),
        },
        "sheets": [
            {
                "name": s.name,
                "state": s.state,
                "max_row": s.max_row,
                "max_column": s.max_column,
                "formula_count": s.formula_count,
                "merged_ranges_count": s.merged_ranges_count,
                "hidden_rows_count": s.hidden_rows,
                "hidden_columns_count": s.hidden_columns,
                "protected": s.protected,
                "data_validation_count": s.data_validation_count,
                "comments_count": s.comments_count,
                "text_cell_count": s.text_cell_count,
                "unlocked_cell_count": s.unlocked_cell_count,
                "filled_cell_count": s.filled_cell_count,
            }
            for s in sheets
        ],
        "named_ranges": [
            {"name": name, "refers_to": value, "scope": scope}
            for name, value, scope in kw["named_ranges"]
        ],
        "vba": {
            "present": kw["vba_present"],
            "payload_bytes": kw["vba_bytes"],
            "payload_sha256": kw["vba_sha"],
        },
        "unsupported_constructs": kw["unsupported_counts"],
    }
    summary_path = output_dir / "template_summary.json"
    summary_path.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")

    cells_path = output_dir / "template_cells.csv"
    _write_csv(
        cells_path,
        TEMPLATE_CELLS_HEADER,
        (
            [
                r.sheet, r.cell, r.row, r.column, r.kind, _bool(r.has_formula), r.formula,
                _bool(r.locked), r.fill_type, r.fill_rgb, _bool(r.has_validation),
                _bool(r.is_merged), r.text_label,
            ]
            for r in cell_records
        ),
    )

    input_path = output_dir / "input_candidates.csv"
    _write_csv(
        input_path,
        INPUT_CANDIDATES_HEADER,
        (
            [c.sheet, c.cell, _bool(c.locked), c.fill_rgb, _bool(c.has_validation),
             c.label_context, c.candidate_reasons]
            for c in input_candidates
        ),
    )

    dependencies_path = output_dir / "template_formula_dependencies.csv"
    _write_csv(dependencies_path, FORMULA_DEPENDENCIES_HEADER, kw["dependency_rows"])

    sheet_dependencies_path = output_dir / "template_sheet_dependencies.csv"
    _write_csv(sheet_dependencies_path, SHEET_DEPENDENCIES_HEADER, kw["sheet_dependency_rows"])

    fingerprint_path = output_dir / "template_fingerprint.json"
    fingerprint_path.write_text(
        json.dumps(kw["fingerprint"], indent=2, ensure_ascii=False), encoding="utf-8"
    )

    fill_distribution_path = output_dir / "template_fill_distribution.csv"
    _write_csv(fill_distribution_path, FILL_DISTRIBUTION_HEADER, kw["fill_distribution"])

    alignment_path = output_dir / "alignment_candidates.csv"
    _write_csv(
        alignment_path,
        ALIGNMENT_HEADER,
        (
            [
                a.generator_sheet, a.generator_cell, a.generator_label, a.template_sheet,
                a.template_cell, a.template_label, a.match_type, a.confidence, a.context,
            ]
            for a in kw["alignment_candidates"]
        ),
    )

    readme_path = output_dir / "README.md"
    readme_path.write_text(_README, encoding="utf-8")

    doc_path.parent.mkdir(parents=True, exist_ok=True)
    doc_path.write_text(
        render_alignment_doc(
            template_name=template_path.name,
            full_sha=kw["full_sha"],
            sheets=sheets,
            input_candidates=input_candidates,
            sheet_dependency_rows=kw["sheet_dependency_rows"],
            alignment_candidates=kw["alignment_candidates"],
            fill_distribution=kw["fill_distribution"],
            vba_present=kw["vba_present"],
            vba_bytes=kw["vba_bytes"],
            unsupported_counts=kw["unsupported_counts"],
            generator_analyzed=kw["generator_analyzed"],
        ),
        encoding="utf-8",
    )

    return [
        summary_path,
        cells_path,
        input_path,
        dependencies_path,
        sheet_dependencies_path,
        fingerprint_path,
        fill_distribution_path,
        alignment_path,
        readme_path,
        doc_path,
    ]


_README = """# Template inventory (Sprint 1.4)

Salida de `scripts/inventory_template.py`. Evidencia estructural de la plantilla
oficial `.xlsm` y *candidatos* de alineación con el workbook generador. Solo
lectura, sin COM, sin ejecutar macros. No interpreta clínica ni confirma mappings.

- `template_summary.json` — estructura por hoja, protección, validaciones,
  comentarios, named ranges, presencia y hash del payload VBA.
- `template_cells.csv` — celdas estructuralmente relevantes (fórmula, texto,
  desbloqueada, fill no-default o con data validation). `text_label` solo lleva
  texto de la plantilla oficial.
- `input_candidates.csv` — celdas que *podrían* ser de ingreso (sin fórmula, no
  merged, desbloqueadas y/o con validación). Son candidatos, no confirmaciones.
- `template_formula_dependencies.csv` / `template_sheet_dependencies.csv` —
  dependencias de fórmula (reutiliza `formula_refs`). Solo aristas observadas.
- `template_fingerprint.json` — hash completo + fingerprint estructural
  (hojas, dimensiones, sentinelas de texto estable, coords de fórmulas, merges)
  para detectar reemplazos silenciosos de plantilla. Experimental.
- `template_fill_distribution.csv` — recuento de celdas por color de relleno.
- `alignment_candidates.csv` — pares (generador, plantilla) por etiqueta.
  `match_type` ∈ {`exact_label`, `normalized_label`, `contextual`, `ambiguous`}.
  `confidence` es una métrica de matching, NO una confirmación funcional. Los
  casos con varios candidatos quedan como `ambiguous` (nunca se elige uno).

Privacidad: del workbook generador solo se leen etiquetas de texto de las hojas
REMASEP; la hoja de detalle de pacientes se excluye por completo.
"""


def render_alignment_doc(
    *,
    template_name: str,
    full_sha: str,
    sheets: list[SheetInfo],
    input_candidates: list[InputCandidate],
    sheet_dependency_rows: list[list],
    alignment_candidates: list[AlignmentCandidate],
    fill_distribution: list[tuple[str, int]],
    vba_present: bool,
    vba_bytes: int | None,
    unsupported_counts: dict[str, int],
    generator_analyzed: bool,
) -> str:
    lines: list[str] = []
    lines.append("# Alineación con la plantilla oficial REMASEP")
    lines.append("")
    lines.append(
        "> Generado por `scripts/inventory_template.py`. Evidencia **estructural** y "
        "*candidatos* de alineación. No interpreta significado clínico, no confirma "
        "mappings y no ejecuta macros."
    )
    lines.append(">")
    lines.append(f"> Plantilla: `{template_name}`  ")
    lines.append(f"> SHA256: `{full_sha}`  ")
    lines.append(f"> Generado: {datetime.now(UTC).strftime('%Y-%m-%dT%H:%M:%SZ')}  ")
    lines.append(f"> VBA presente: {'sí' if vba_present else 'no'}"
                 + (f" ({vba_bytes} bytes)" if vba_bytes else ""))
    lines.append("")

    # 1. Estructura
    lines.append("## 1. Estructura de la plantilla oficial")
    lines.append("")
    lines.append("| hoja | estado | filas | cols | fórmulas | merges | prot. | val. | filas ocultas | cols ocultas |")
    lines.append("| --- | --- | ---: | ---: | ---: | ---: | :---: | ---: | ---: | ---: |")
    for s in sheets:
        lines.append(
            f"| {s.name} | {s.state} | {s.max_row} | {s.max_column} | {s.formula_count} | "
            f"{s.merged_ranges_count} | {'sí' if s.protected else 'no'} | "
            f"{s.data_validation_count} | {s.hidden_rows} | {s.hidden_columns} |"
        )
    lines.append("")

    # 2. Hojas con fórmulas
    with_formulas = [s for s in sheets if s.formula_count]
    lines.append("## 2. Hojas con fórmulas")
    lines.append("")
    lines.append(
        ", ".join(f"`{s.name}` ({s.formula_count})" for s in with_formulas) or "ninguna"
    )
    lines.append("")

    # 3. Flujo entre hojas
    cross = [(s, t, n) for s, t, n in sheet_dependency_rows if s != t]
    lines.append("## 3. Hojas que se alimentan entre sí (observado)")
    lines.append("")
    lines.append("```mermaid")
    lines.append("graph LR")
    ids: dict[str, str] = {}
    for s in sheets:
        ids[s.name] = f"n{len(ids)}"
        lines.append(f'    {ids[s.name]}["{s.name}"]')
    for source, target, count in sorted(cross, key=lambda r: (-r[2], r[0], r[1])):
        src = ids.setdefault(source, f"n{len(ids)}")
        dst = ids.setdefault(target, f"n{len(ids)}")
        lines.append(f"    {src} -->|{count}| {dst}")
    if not cross:
        lines.append("    %% sin dependencias entre hojas distintas")
    lines.append("```")
    lines.append("")
    lines.append("| origen | destino | referencias |")
    lines.append("| --- | --- | ---: |")
    for source, target, count in sorted(cross, key=lambda r: (-r[2], r[0], r[1])):
        lines.append(f"| {source} | {target} | {count} |")
    if not cross:
        lines.append("| — | — | 0 |")
    lines.append("")

    # 4. Input candidates
    per_sheet_inputs: Counter[str] = Counter(c.sheet for c in input_candidates)
    lines.append("## 4. Input candidates")
    lines.append("")
    lines.append(f"Total: **{len(input_candidates)}** celdas candidatas a ingreso.")
    lines.append("")
    lines.append("| hoja | candidatos |")
    lines.append("| --- | ---: |")
    for name, count in sorted(per_sheet_inputs.items(), key=lambda kv: (-kv[1], kv[0])):
        lines.append(f"| {name} | {count} |")
    if not per_sheet_inputs:
        lines.append("| — | 0 |")
    lines.append("")

    # 5. Fills / protección
    lines.append("## 5. Distribución de rellenos y protección")
    lines.append("")
    protected = [s.name for s in sheets if s.protected]
    unprotected = [s.name for s in sheets if not s.protected]
    lines.append(f"- Hojas protegidas: {', '.join(f'`{n}`' for n in protected) or '—'}")
    lines.append(f"- Hojas sin proteger: {', '.join(f'`{n}`' for n in unprotected) or '—'}")
    lines.append(f"- Colores de relleno distintos: **{len(fill_distribution)}**")
    lines.append("")
    lines.append("| fill_rgb | celdas |")
    lines.append("| --- | ---: |")
    for rgb, count in fill_distribution[:15]:
        lines.append(f"| {_md(rgb)} | {count} |")
    if not fill_distribution:
        lines.append("| — | 0 |")
    lines.append("")

    # 6 + 7. Alignment
    lines.append("## 6. Candidatos de alineación por hoja")
    lines.append("")
    if not generator_analyzed:
        lines.append("_No se analizó el workbook generador (`--generator` no indicado)._")
        lines.append("")
    else:
        by_type: Counter[str] = Counter(a.match_type for a in alignment_candidates)
        by_pair: Counter[tuple[str, str]] = Counter(
            (a.generator_sheet, a.template_sheet) for a in alignment_candidates
        )
        lines.append(
            f"Total candidatos: **{len(alignment_candidates)}** — "
            + ", ".join(f"{k}: {v}" for k, v in sorted(by_type.items()))
        )
        lines.append("")
        lines.append("| generador → plantilla | candidatos |")
        lines.append("| --- | ---: |")
        for (gen_sheet, tmpl_sheet), count in sorted(by_pair.items(), key=lambda kv: -kv[1]):
            lines.append(f"| {gen_sheet} → {tmpl_sheet} | {count} |")
        lines.append("")

        lines.append("## 7. Matches ambiguos")
        lines.append("")
        ambiguous = [a for a in alignment_candidates if a.match_type == "ambiguous"]
        if ambiguous:
            lines.append(
                f"**{len(ambiguous)}** filas ambiguas "
                f"({len({a.generator_cell for a in ambiguous})} etiquetas del generador con "
                "más de un candidato). No se elige ninguno automáticamente."
            )
            lines.append("")
            lines.append("| gen cell | etiqueta | plantilla |")
            lines.append("| --- | --- | --- |")
            for a in ambiguous[:25]:
                lines.append(
                    f"| {a.generator_sheet}!{a.generator_cell} | {_md(a.generator_label)} | "
                    f"{a.template_sheet}!{a.template_cell} |"
                )
        else:
            lines.append("Ninguno.")
        lines.append("")

    # 8. Observaciones
    lines.append("## 8. Observaciones que requieren revisión humana")
    lines.append("")
    lines.append(
        "- Los `input_candidates` se basan solo en protección/estilo/validación; "
        "hay que confirmar cuáles son realmente celdas de ingreso del proceso."
    )
    lines.append(
        "- Los `alignment_candidates` son coincidencias de **etiqueta**, no de "
        "significado. Cada `exact_label`/`normalized_label` debe validarse; los "
        "`ambiguous` requieren decisión explícita."
    )
    lines.append(
        "- Constructs de fórmula no soportados por el analizador: "
        + (", ".join(f"{k} ({v})" for k, v in sorted(unsupported_counts.items())) or "ninguno")
        + "."
    )
    lines.append(
        f"- VBA: {'presente' if vba_present else 'ausente'}"
        + (f", payload {vba_bytes} bytes (no interpretado)." if vba_present else ".")
    )
    lines.append(
        "- El `template_fingerprint.json` es experimental; sirve para detectar si "
        "una futura plantilla 'V1.4' cambia de estructura sin cambiar de nombre."
    )
    lines.append("")
    return "\n".join(lines) + "\n"


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="inventory_template.py",
        description=(
            "Inventario estructural de la plantilla oficial REMASEP y candidatos de "
            "alineación con el workbook generador. Solo lectura, sin COM ni macros."
        ),
    )
    parser.add_argument("template", help="Ruta al .xlsm de la plantilla oficial.")
    parser.add_argument(
        "--generator",
        default=None,
        help="Ruta al workbook generador (opcional). Sin esto se omite la Parte F.",
    )
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT, help=f"(def: {DEFAULT_OUTPUT})")
    parser.add_argument("--doc", default=DEFAULT_DOC, help=f"(def: {DEFAULT_DOC})")
    parser.add_argument(
        "--generator-detail-sheet",
        default=DEFAULT_GENERATOR_DETAIL_SHEET,
        help="Hoja de pacientes del generador a excluir SIEMPRE de la lectura de etiquetas.",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    template_path = Path(args.template)
    generator_path = Path(args.generator) if args.generator else None

    try:
        result = analyze_template(
            template_path,
            Path(args.output),
            generator_path=generator_path,
            doc_path=Path(args.doc),
            generator_detail_sheet=args.generator_detail_sheet,
        )
    except InventoryError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    total_formulas = sum(s.formula_count for s in result.sheets)
    cross = [(s, t, n) for s, t, n in result.sheet_dependency_rows if s != t]
    by_type: Counter[str] = Counter(a.match_type for a in result.alignment_candidates)

    print(f"Plantilla analizada: {template_path}")
    print(f"  hojas oficiales ({len(result.sheets)}): {', '.join(s.name for s in result.sheets)}")
    print("  fórmulas por hoja:")
    for s in result.sheets:
        print(f"    {s.name}: {s.formula_count}")
    print(f"  fórmulas totales: {total_formulas}")
    print(f"  input candidates: {len(result.input_candidates)}")
    print(f"  colores de relleno distintos: {len(result.fill_distribution)}")
    print("  hojas cross-linked por fórmulas:")
    for source, target, count in sorted(cross, key=lambda r: (-r[2], r[0], r[1])):
        print(f"    {source} -> {target} : {count}")
    if not cross:
        print("    (ninguna)")
    print(f"  alignment candidates: {len(result.alignment_candidates)}")
    print(f"    exact_label: {by_type.get('exact_label', 0)}")
    print(f"    normalized_label: {by_type.get('normalized_label', 0)}")
    print(f"    contextual: {by_type.get('contextual', 0)}")
    print(f"    ambiguous: {by_type.get('ambiguous', 0)}")
    print(f"  VBA presente: {result.vba_present}"
          + (f" (payload {result.vba_payload_bytes} bytes, sha256 {result.vba_payload_sha256[:16]}…)"
             if result.vba_present else ""))
    print(
        "  constructs no soportados: "
        + (", ".join(f"{k}={v}" for k, v in sorted(result.unsupported_counts.items())) or "ninguno")
    )
    print(f"Artefactos escritos en: {result.output_dir}")
    for path in result.files:
        print(f"  - {path}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
