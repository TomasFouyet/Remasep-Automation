"""Comparación segura de dos workbooks REMASEP (Sprint 3.4).

Sólo lectura: abre cada archivo con ``openpyxl`` (``data_only`` en ``False`` para
las fórmulas y en ``True`` para los valores cacheados por Excel), **nunca**
escribe, **nunca** guarda, **no** usa COM, LibreOffice ni macros, y **no**
recalcula nada.

Expone:

- `WorkbookPair` — los cuatro workbooks abiertos (before/after × formula/value).
- `compare_structure` -> `StructureComparison` — nombres/orden de hojas,
  dimensiones, recuento y expresión de fórmulas, merges, VBA, SHA256.
- `diff_formulas` -> lista de `FormulaChange` — cambios de **expresión** de
  fórmula (se espera 0 para estos archivos, no se *hardcodea*).
- `diff_values` -> lista de `DeltaCell` — celdas cuyo **valor** cambió. El
  ``change_kind`` se decide con la fórmula ANTES **y** DESPUÉS:
  ``DIRECT_INPUT_CHANGE`` (sin fórmula en ninguna versión) ·
  ``FORMULA_RESULT_CHANGE`` (misma expresión, distinto cache) ·
  ``FORMULA_EXPRESSION_CHANGE`` (la expresión se reescribió: add/remove/modify).
- `build_delta_edges` -> aristas de dependencia estructural entre celdas
  cambiadas (referencias A1 de la fórmula *después*).

Privacidad: estos workbooks son outputs REMASEP **agregados**; aun así los
artefactos sólo llevan coordenadas, etiquetas del formulario, fórmulas y valores
agregados — nunca filas individuales.
"""

from __future__ import annotations

import contextlib
import hashlib
import re
import zipfile
from collections.abc import Iterable
from dataclasses import dataclass
from pathlib import Path

import openpyxl

from remasep.core.errors import RemasepError

_VBA_ENTRY = "xl/vbaProject.bin"

# ``change_kind`` de una celda cuyo valor cambió, decidido con la fórmula
# ANTES **y** DESPUÉS (no sólo la de la versión final):
#   - sin fórmula en ninguna versión                -> DIRECT_INPUT_CHANGE
#   - misma expresión de fórmula, distinto cache    -> FORMULA_RESULT_CHANGE
#   - expresión de fórmula distinta (add/remove/mod)-> FORMULA_EXPRESSION_CHANGE
DIRECT_INPUT_CHANGE = "DIRECT_INPUT_CHANGE"
FORMULA_RESULT_CHANGE = "FORMULA_RESULT_CHANGE"
FORMULA_EXPRESSION_CHANGE = "FORMULA_EXPRESSION_CHANGE"

FORMULA_ADDED = "FORMULA_ADDED"
FORMULA_REMOVED = "FORMULA_REMOVED"
FORMULA_MODIFIED = "FORMULA_MODIFIED"


def _formula_change_type(before_formula: str, after_formula: str) -> str:
    if before_formula and after_formula:
        return FORMULA_MODIFIED
    if after_formula:
        return FORMULA_ADDED
    return FORMULA_REMOVED


class WorkbookDeltaError(RemasepError):
    """No se pudo comparar el par de workbooks."""


# ---------------------------------------------------------------------------
# Apertura segura
# ---------------------------------------------------------------------------


def compute_sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1024 * 1024), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _keep_vba(path: Path) -> bool:
    return path.suffix.lower() in {".xlsm", ".xltm", ".xlsb"}


def _load(path: Path, *, data_only: bool):
    try:
        return openpyxl.load_workbook(
            path, data_only=data_only, read_only=False, keep_vba=_keep_vba(path)
        )
    except FileNotFoundError as exc:
        raise WorkbookDeltaError(f"no existe el workbook: {path}") from exc
    except Exception as exc:  # mensaje claro, no traceback
        raise WorkbookDeltaError(
            f"no se pudo leer '{path}': {exc.__class__.__name__}: {exc}"
        ) from exc


@dataclass
class WorkbookPair:
    before_path: Path
    after_path: Path
    before_sha256: str
    after_sha256: str
    before_formula: openpyxl.Workbook
    after_formula: openpyxl.Workbook
    before_value: openpyxl.Workbook
    after_value: openpyxl.Workbook

    def close(self) -> None:
        for wb in (
            self.before_formula, self.after_formula,
            self.before_value, self.after_value,
        ):
            with contextlib.suppress(Exception):
                wb.close()


def open_pair(before_path: str | Path, after_path: str | Path) -> WorkbookPair:
    before, after = Path(before_path), Path(after_path)
    for p in (before, after):
        if not p.is_file():
            raise WorkbookDeltaError(f"no es un archivo: {p}")
    return WorkbookPair(
        before_path=before,
        after_path=after,
        before_sha256=compute_sha256(before),
        after_sha256=compute_sha256(after),
        before_formula=_load(before, data_only=False),
        after_formula=_load(after, data_only=False),
        before_value=_load(before, data_only=True),
        after_value=_load(after, data_only=True),
    )


# ---------------------------------------------------------------------------
# Utilidades de celda
# ---------------------------------------------------------------------------


def _formula_text(value: object) -> str:
    if isinstance(value, str) and value.startswith("="):
        return value
    text = getattr(value, "text", None)  # ArrayFormula
    if isinstance(text, str):
        return text if text.startswith("=") else f"={text}"
    return ""


def _is_blankish(value: object) -> bool:
    return value is None or value == ""


def _values_equal(a: object, b: object) -> bool:
    if a == b:
        return True
    if _is_blankish(a) and _is_blankish(b):
        return True
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return abs(float(a) - float(b)) < 1e-9
    return False


def _sheet_bounds(ws_a, ws_b) -> tuple[int, int]:
    return (
        max(ws_a.max_row or 1, ws_b.max_row or 1),
        max(ws_a.max_column or 1, ws_b.max_column or 1),
    )


# ---------------------------------------------------------------------------
# 4. Comparación de estructura
# ---------------------------------------------------------------------------


@dataclass
class SheetStructure:
    name: str
    max_row: int
    max_column: int
    formula_count: int
    merged_ranges: int


@dataclass
class StructureComparison:
    sheet_names_before: list[str]
    sheet_names_after: list[str]
    sheet_order_changed: bool
    added_sheets: list[str]
    removed_sheets: list[str]
    per_sheet: list[dict]
    vba_before: dict
    vba_after: dict
    sha256_before: str
    sha256_after: str
    significant_changes: list[str]

    def to_dict(self) -> dict:
        return {
            "sha256_before": self.sha256_before,
            "sha256_after": self.sha256_after,
            "sheet_names_before": self.sheet_names_before,
            "sheet_names_after": self.sheet_names_after,
            "sheet_order_changed": self.sheet_order_changed,
            "added_sheets": self.added_sheets,
            "removed_sheets": self.removed_sheets,
            "vba_before": self.vba_before,
            "vba_after": self.vba_after,
            "per_sheet": self.per_sheet,
            "significant_changes": self.significant_changes,
        }


def _vba_info(path: Path) -> dict:
    if not zipfile.is_zipfile(path):
        return {"present": False, "payload_bytes": None, "payload_sha256": None}
    with zipfile.ZipFile(path) as archive:
        if _VBA_ENTRY not in archive.namelist():
            return {"present": False, "payload_bytes": None, "payload_sha256": None}
        payload = archive.read(_VBA_ENTRY)
    return {
        "present": True,
        "payload_bytes": len(payload),
        "payload_sha256": hashlib.sha256(payload).hexdigest(),
    }


def _sheet_structs(wb) -> dict[str, SheetStructure]:
    out: dict[str, SheetStructure] = {}
    for ws in wb.worksheets:
        fcount = 0
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    fcount += 1
        out[ws.title] = SheetStructure(
            name=ws.title,
            max_row=ws.max_row or 0,
            max_column=ws.max_column or 0,
            formula_count=fcount,
            merged_ranges=len(list(ws.merged_cells.ranges)),
        )
    return out


def compare_structure(pair: WorkbookPair) -> StructureComparison:
    names_before = list(pair.before_formula.sheetnames)
    names_after = list(pair.after_formula.sheetnames)
    before = _sheet_structs(pair.before_formula)
    after = _sheet_structs(pair.after_formula)

    added = [n for n in names_after if n not in before]
    removed = [n for n in names_before if n not in after]
    order_changed = [n for n in names_before if n in after] != [
        n for n in names_after if n in before
    ]

    significant: list[str] = []
    if added:
        significant.append(f"hojas añadidas: {added}")
    if removed:
        significant.append(f"hojas eliminadas: {removed}")
    if order_changed:
        significant.append("cambió el orden de las hojas")

    per_sheet: list[dict] = []
    for name in names_after:
        b = before.get(name)
        a = after[name]
        row = {
            "sheet": name,
            "present_before": b is not None,
            "max_row_before": b.max_row if b else None,
            "max_row_after": a.max_row,
            "max_column_before": b.max_column if b else None,
            "max_column_after": a.max_column,
            "formula_count_before": b.formula_count if b else None,
            "formula_count_after": a.formula_count,
            "merged_ranges_before": b.merged_ranges if b else None,
            "merged_ranges_after": a.merged_ranges,
        }
        per_sheet.append(row)
        if b is not None:
            if b.formula_count != a.formula_count:
                significant.append(
                    f"{name}: recuento de fórmulas {b.formula_count} -> {a.formula_count}"
                )
            if (b.max_row, b.max_column) != (a.max_row, a.max_column):
                significant.append(
                    f"{name}: dimensiones {b.max_row}x{b.max_column} -> "
                    f"{a.max_row}x{a.max_column}"
                )
            if b.merged_ranges != a.merged_ranges:
                significant.append(
                    f"{name}: merges {b.merged_ranges} -> {a.merged_ranges}"
                )

    vba_before = _vba_info(pair.before_path)
    vba_after = _vba_info(pair.after_path)
    if vba_before.get("present") != vba_after.get("present"):
        significant.append(
            f"presencia de VBA cambió: {vba_before.get('present')} -> {vba_after.get('present')}"
        )
    elif vba_before.get("payload_sha256") != vba_after.get("payload_sha256"):
        # payload distinto pero VBA presente en ambos: sin más contexto no se puede
        # afirmar un cambio de lógica (Excel recompila el proyecto al guardar).
        significant.append(
            "payload VBA con distinto sha256 (VBA presente en ambos; sin cambios de "
            "fórmula ni de dimensiones: compatible con un simple re-guardado)"
        )

    return StructureComparison(
        sheet_names_before=names_before,
        sheet_names_after=names_after,
        sheet_order_changed=order_changed,
        added_sheets=added,
        removed_sheets=removed,
        per_sheet=per_sheet,
        vba_before=vba_before,
        vba_after=vba_after,
        sha256_before=pair.before_sha256,
        sha256_after=pair.after_sha256,
        significant_changes=significant,
    )


# ---------------------------------------------------------------------------
# 5. Diff de expresiones de fórmula
# ---------------------------------------------------------------------------


@dataclass
class FormulaChange:
    sheet: str
    cell: str
    before_formula: str
    after_formula: str
    change_type: str


def diff_formulas(pair: WorkbookPair) -> list[FormulaChange]:
    changes: list[FormulaChange] = []
    common = [n for n in pair.after_formula.sheetnames if n in pair.before_formula.sheetnames]
    for name in common:
        ws_b = pair.before_formula[name]
        ws_a = pair.after_formula[name]
        max_row, max_col = _sheet_bounds(ws_b, ws_a)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                fb = _formula_text(ws_b.cell(row=r, column=c).value)
                fa = _formula_text(ws_a.cell(row=r, column=c).value)
                if fb == fa:
                    continue
                changes.append(
                    FormulaChange(
                        sheet=name,
                        cell=ws_a.cell(row=r, column=c).coordinate,
                        before_formula=fb,
                        after_formula=fa,
                        change_type=_formula_change_type(fb, fa),
                    )
                )
    return changes


# ---------------------------------------------------------------------------
# 6. Delta de valores
# ---------------------------------------------------------------------------


@dataclass
class DeltaCell:
    sheet: str
    cell: str
    row: int
    column: int
    before_value: object
    after_value: object
    formula: str  # fórmula en el workbook *después* ("" si no hay)
    change_kind: str
    before_formula: str = ""  # fórmula en el workbook *antes* ("" si no había)
    section_path: tuple[str, ...] = ()
    row_path: tuple[str, ...] = ()
    column_path: tuple[str, ...] = ()

    @property
    def key(self) -> tuple[str, str]:
        return (self.sheet, self.cell)

    @property
    def formula_change_type(self) -> str:
        """Sólo para ``FORMULA_EXPRESSION_CHANGE``: ADDED / REMOVED / MODIFIED."""
        if self.change_kind != FORMULA_EXPRESSION_CHANGE:
            return ""
        return _formula_change_type(self.before_formula, self.formula)


def _scalar(value: object) -> object:
    """Normaliza para el CSV: sin objetos raros, blanco -> ''."""
    if _is_blankish(value):
        return ""
    if isinstance(value, bool):
        return value
    if isinstance(value, float) and value.is_integer():
        return int(value)
    return value


def diff_values(pair: WorkbookPair) -> list[DeltaCell]:
    deltas: list[DeltaCell] = []
    common = [n for n in pair.after_value.sheetnames if n in pair.before_value.sheetnames]
    for name in common:
        vb = pair.before_value[name]
        va = pair.after_value[name]
        fb = pair.before_formula[name]
        fa = pair.after_formula[name]
        max_row, max_col = _sheet_bounds(vb, va)
        for r in range(1, max_row + 1):
            for c in range(1, max_col + 1):
                a = vb.cell(row=r, column=c).value
                b = va.cell(row=r, column=c).value
                if _values_equal(a, b):
                    continue
                before_formula = _formula_text(fb.cell(row=r, column=c).value)
                after_formula = _formula_text(fa.cell(row=r, column=c).value)
                if before_formula != after_formula:
                    kind = FORMULA_EXPRESSION_CHANGE
                elif after_formula:
                    kind = FORMULA_RESULT_CHANGE
                else:
                    kind = DIRECT_INPUT_CHANGE
                deltas.append(
                    DeltaCell(
                        sheet=name,
                        cell=va.cell(row=r, column=c).coordinate,
                        row=r,
                        column=c,
                        before_value=_scalar(a),
                        after_value=_scalar(b),
                        formula=after_formula,
                        before_formula=before_formula,
                        change_kind=kind,
                    )
                )
    return deltas


# ---------------------------------------------------------------------------
# 8. Contexto semántico (best-effort, reutiliza el layout de Sprint 3.1)
# ---------------------------------------------------------------------------


def attach_semantic_context(
    pair: WorkbookPair,
    deltas: Iterable[DeltaCell],
    *,
    skip_sheets: set[str] | None = None,
) -> None:
    """Rellena ``section_path`` / ``row_path`` / ``column_path`` in-place.

    Usa `remasep.services.semantic_layout.SheetLayout` sobre el workbook
    ``data_only=True`` (así las celdas de fórmula aportan su *valor*, no el texto
    ``=SUM(...)``, y no contaminan los rótulos). Si una hoja o celda no encaja en
    ese modelo, deja las tuplas vacías (sin fallar). ``skip_sheets`` excluye hojas
    que no son grillas de datos (p.ej. ``CONTROL``).
    """
    try:
        from remasep.services.semantic_layout import SheetLayout
    except Exception:  # noqa: BLE001
        return

    skip = skip_sheets or set()
    layouts: dict[str, object] = {}
    for delta in deltas:
        if delta.sheet in skip:
            continue
        layout = layouts.get(delta.sheet, "?")
        if layout == "?":
            try:
                layout = SheetLayout(pair.after_value[delta.sheet])
            except Exception:  # noqa: BLE001
                layout = None
            layouts[delta.sheet] = layout
        if layout is None:
            continue
        ctx = None
        with contextlib.suppress(Exception):
            ctx = layout.cell_context(delta.cell)
        if ctx is None:
            continue
        delta.section_path = tuple(ctx.section_labels_raw)
        delta.row_path = tuple(ctx.row_labels_raw)
        delta.column_path = tuple(ctx.column_labels_raw)


# ---------------------------------------------------------------------------
# 19. Grafo de dependencia estructural entre celdas cambiadas
# ---------------------------------------------------------------------------


@dataclass
class DeltaEdge:
    source_sheet: str
    source_cell: str
    target_sheet: str
    target_cell: str
    dependency_type: str  # cell | range | cross_sheet


@dataclass(frozen=True)
class _Ref:
    sheet: str | None
    col_start: int
    col_end: int
    row_start: int
    row_end: int
    is_single: bool


# literales de texto, para no confundir "A1" dentro de "..." con una referencia
_STRING_RE = re.compile(r'"(?:[^"]|"")*"')
_SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)"
_A1 = r"\$?[A-Za-z]{1,3}\$?[0-9]+"
_COLREF = r"\$?[A-Za-z]{1,3}"
_REF_RE = re.compile(
    rf"(?<![A-Za-z0-9_.!$\]])"
    rf"(?:(?P<sheet>{_SHEET})!)?"
    rf"(?:(?P<r1>{_A1}):(?P<r2>{_A1})"
    rf"|(?P<c1>{_COLREF}):(?P<c2>{_COLREF})"
    rf"|(?P<a>{_A1}))"
    rf"(?![A-Za-z0-9_.(\[])"
)
_A1_SPLIT_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")


def _col_index(letters: str) -> int:
    index = 0
    for char in letters.upper():
        index = index * 26 + (ord(char) - 64)
    return index


def _unquote(sheet: str | None) -> str | None:
    if sheet is None:
        return None
    if len(sheet) >= 2 and sheet[0] == sheet[-1] == "'":
        return sheet[1:-1].replace("''", "'")
    return sheet


def _parse_refs(formula: str) -> list[_Ref]:
    body = _STRING_RE.sub(lambda m: " " * len(m.group(0)), formula)
    out: list[_Ref] = []
    for m in _REF_RE.finditer(body):
        sheet = _unquote(m.group("sheet"))
        if m.group("r1"):
            c1, r1 = _A1_SPLIT_RE.match(m.group("r1")).groups()
            c2, r2 = _A1_SPLIT_RE.match(m.group("r2")).groups()
            ci1, ci2 = _col_index(c1), _col_index(c2)
            ri1, ri2 = int(r1), int(r2)
            out.append(_Ref(sheet, min(ci1, ci2), max(ci1, ci2),
                            min(ri1, ri2), max(ri1, ri2),
                            is_single=ci1 == ci2 and ri1 == ri2))
        elif m.group("c1"):
            ci1, ci2 = _col_index(m.group("c1").replace("$", "")), _col_index(m.group("c2").replace("$", ""))
            out.append(_Ref(sheet, min(ci1, ci2), max(ci1, ci2), 1, 1_048_576, is_single=False))
        else:
            col, row = _A1_SPLIT_RE.match(m.group("a")).groups()
            ci = _col_index(col)
            out.append(_Ref(sheet, ci, ci, int(row), int(row), is_single=True))
    return out


def _ref_contains(ref: _Ref, sheet: str, col: int, row: int, host_sheet: str) -> bool:
    if (ref.sheet or host_sheet) != sheet:
        return False
    return ref.col_start <= col <= ref.col_end and ref.row_start <= row <= ref.row_end


def build_delta_edges(deltas: list[DeltaCell]) -> list[DeltaEdge]:
    edges: list[DeltaEdge] = []
    for target in deltas:
        if not target.formula:
            continue
        refs = _parse_refs(target.formula)
        for src in deltas:
            if src.key == target.key:
                continue
            for ref in refs:
                if _ref_contains(ref, src.sheet, src.column, src.row, target.sheet):
                    dep = "cross_sheet" if src.sheet != target.sheet else (
                        "cell" if ref.is_single else "range"
                    )
                    edges.append(
                        DeltaEdge(
                            source_sheet=src.sheet, source_cell=src.cell,
                            target_sheet=target.sheet, target_cell=target.cell,
                            dependency_type=dep,
                        )
                    )
                    break
    return edges
