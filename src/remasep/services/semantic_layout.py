"""Extracción del contexto **visual** de las celdas de una hoja legacy (Sprint 3.1).

Empieza la capa semántica: en vez de razonar con coordenadas (`G54`) queremos
recuperar los rótulos que un humano *ve* alrededor de una métrica — encabezados
de fila, encabezados de columna y encabezados de sección — resolviendo celdas
combinadas, sin modificar el workbook y **sin interpretar** todavía qué
significan (eso es Sprint 3.2).

Contiene:

- :func:`normalize_semantic_label` — normalización propia de la capa semántica
  (mayúsculas, sin tildes, espacios colapsados). **No** reutiliza
  ``normalize_legacy_text`` ni ``core.text``: la capa semántica evoluciona
  aparte.
- :class:`SheetLayout` — lee una hoja (openpyxl, sólo lectura), resuelve merges
  y expone ``cell_context(coord) -> CellContext``.
- :func:`parse_age_label` / :func:`age_bounds_from_formula` /
  :func:`text_criteria_from_formula` — **evidencia técnica** para comparar más
  tarde rótulo vs. fórmula. No es un mapping clínico.
"""

from __future__ import annotations

import re
import unicodedata
from dataclasses import dataclass

# ---------------------------------------------------------------------------
# Normalización de rótulos semánticos
# ---------------------------------------------------------------------------


def normalize_semantic_label(value: object) -> str:
    """Normaliza un rótulo para *matching* posterior; nunca reemplaza al raw.

    Mayúsculas, sin diacríticos, whitespace colapsado. Conserva dígitos y
    puntuación (``5010009 - VIDRIO IONÓMERO`` -> ``5010009 - VIDRIO IONOMERO``;
    ``20 A 24 AÑOS`` -> ``20 A 24 ANOS``). Deliberadamente independiente de
    ``remasep.core.text`` y de ``normalize_legacy_text``.
    """
    if value is None:
        return ""
    text = unicodedata.normalize("NFD", str(value))
    text = "".join(ch for ch in text if unicodedata.category(ch) != "Mn")
    text = text.upper()
    return " ".join(text.split())


# ---------------------------------------------------------------------------
# Contexto de celda
# ---------------------------------------------------------------------------

CONTEXT_COMPLETE = "COMPLETE"
CONTEXT_PARTIAL = "PARTIAL"
CONTEXT_AMBIGUOUS = "AMBIGUOUS"
CONTEXT_NO_CONTEXT = "NO_CONTEXT"

_MAX_ROW_LEVELS = 3
_MAX_COLUMN_LEVELS = 3
_MAX_SECTION_LEVELS = 2


@dataclass(frozen=True)
class CellContext:
    """Rótulos visibles alrededor de una celda (raw + normalizado)."""

    sheet: str
    coord: str
    row_labels_raw: tuple[str, ...]
    column_labels_raw: tuple[str, ...]
    section_labels_raw: tuple[str, ...]
    context_status: str
    notes: tuple[str, ...] = ()

    @property
    def row_labels_norm(self) -> tuple[str, ...]:
        return tuple(normalize_semantic_label(x) for x in self.row_labels_raw)

    @property
    def column_labels_norm(self) -> tuple[str, ...]:
        return tuple(normalize_semantic_label(x) for x in self.column_labels_raw)

    @property
    def section_labels_norm(self) -> tuple[str, ...]:
        return tuple(normalize_semantic_label(x) for x in self.section_labels_raw)

    @property
    def row_hierarchy_depth(self) -> int:
        return len(self.row_labels_raw)

    @property
    def column_hierarchy_depth(self) -> int:
        return len(self.column_labels_raw)

    def semantic_signature(self) -> str:
        """Firma preliminar ``hoja :: sección… :: fila… :: columna…`` (normalizada).

        **Candidata**: no reemplaza al ``metric_id`` técnico
        (``LEGACY::<hoja>::<celda>``).
        """
        parts = [
            *self.section_labels_norm,
            *self.row_labels_norm,
            *self.column_labels_norm,
        ]
        parts = [p for p in parts if p]
        body = " :: ".join(parts) if parts else "<NO_CONTEXT>"
        return f"{normalize_semantic_label(self.sheet)} :: {body}"


_COORD_RE = re.compile(r"^([A-Za-z]{1,3})([0-9]+)$")
# Encabezado de grupo típico del formato MINSAL: "A.1.", "A.4.", "B.-", "C -"
_GROUP_HEADER_RE = re.compile(r"^\s*[A-Z]\s*[.\-]\s*(\d+\s*[.\-]|\-|)", re.IGNORECASE)
_SECTION_KEYWORD_RE = re.compile(r"(?i)^\s*secci[oó]n\b")


def _col_to_index(col: str) -> int:
    index = 0
    for ch in col.upper():
        index = index * 26 + (ord(ch) - 64)
    return index


def _index_to_col(index: int) -> str:
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


class SheetLayout:
    """Vista de sólo lectura de una hoja: valores resueltos por merge + `bold`."""

    def __init__(self, worksheet) -> None:
        self.title: str = worksheet.title
        self._value: dict[tuple[int, int], object] = {}
        self._bold: set[tuple[int, int]] = set()
        max_row = 0
        max_col = 0
        for row in worksheet.iter_rows():
            for cell in row:
                if cell.value is None:
                    continue
                self._value[(cell.row, cell.column)] = cell.value
                max_row = max(max_row, cell.row)
                max_col = max(max_col, cell.column)
                font = cell.font
                if font is not None and font.bold:
                    self._bold.add((cell.row, cell.column))
        self.max_row = max_row or 1
        self.max_col = max_col or 1
        self._anchor: dict[tuple[int, int], tuple[int, int]] = {}
        for mr in worksheet.merged_cells.ranges:
            anchor = (mr.min_row, mr.min_col)
            for r in range(mr.min_row, mr.max_row + 1):
                for c in range(mr.min_col, mr.max_col + 1):
                    self._anchor[(r, c)] = anchor

        # Clasificación de filas de encabezado (una pasada). Una fila-rótulo de
        # encabezado tiene: una celda de texto en **negrita** en las columnas
        # izquierdas, ≤ 2 valores de texto distintos y no está "repartida" por la
        # grilla (esa es la banda de encabezados de columna).
        self._bold_label_rows: dict[int, tuple[int, str]] = {}
        self._keyword_section_rows: set[int] = set()
        for r in range(1, self.max_row + 1):
            texts = self._row_text_columns(r)
            if not texts or self._is_wide_header_row(r):
                continue
            first_col, first_text = texts[0]
            if first_col > 4 or len({t for _c, t in texts}) > 2 or not self.is_bold(r, first_col):
                continue
            if _SECTION_KEYWORD_RE.match(first_text):
                self._keyword_section_rows.add(r)
                self._bold_label_rows[r] = (first_col, first_text)
            elif first_col >= 2 and len(first_text) >= 4:
                # col A suele ser una columna técnica (nº de línea): no es sección
                self._bold_label_rows[r] = (first_col, first_text)

        # Si la hoja usa la palabra "SECCIÓN", ésas son las secciones y el resto
        # de encabezados en negrita (p.ej. "A.4. …") son niveles de fila. Si no
        # (p.ej. B2 ANEXO), los encabezados en negrita hacen de sección.
        if self._keyword_section_rows:
            self._section_rows = set(self._keyword_section_rows)
        else:
            self._section_rows = set(self._bold_label_rows)
        self._group_rows = set(self._bold_label_rows) - self._section_rows

    # --- acceso a celdas -------------------------------------------------

    def _resolved_key(self, row: int, col: int) -> tuple[int, int]:
        return self._anchor.get((row, col), (row, col))

    def value(self, row: int, col: int) -> object:
        return self._value.get(self._resolved_key(row, col))

    def text(self, row: int, col: int) -> str:
        value = self.value(row, col)
        return "" if value is None else str(value).strip()

    def is_text(self, row: int, col: int) -> bool:
        value = self.value(row, col)
        return isinstance(value, str) and value.strip() != ""

    def is_number(self, row: int, col: int) -> bool:
        value = self.value(row, col)
        return isinstance(value, (int, float)) and not isinstance(value, bool)

    def is_bold(self, row: int, col: int) -> bool:
        return self._resolved_key(row, col) in self._bold

    # --- filas de encabezado / sección --------------------------------

    def _row_text_columns(self, row: int) -> list[tuple[int, str]]:
        out: list[tuple[int, str]] = []
        last = None
        for col in range(1, self.max_col + 1):
            if self.is_text(row, col):
                t = self.text(row, col)
                if t != last:  # dedup de spillover de merge (ya resuelto)
                    out.append((col, t))
                last = t
            else:
                last = None
        return out

    def _is_wide_header_row(self, row: int) -> bool:
        """Fila de banda de encabezados de columna (texto repartido por la grilla)."""
        texts = self._row_text_columns(row)
        if len(texts) >= 4:
            return True
        if texts:
            spread = texts[-1][0] - texts[0][0]
            return spread >= 15
        return False

    # --- extracción de contexto --------------------------------------

    def _nearest_section_row(self, row: int) -> int:
        return max((rr for rr in self._section_rows if rr < row), default=0)

    def section_labels(self, row: int, col: int) -> list[str]:
        inner = self._nearest_section_row(row)
        if not inner:
            return []
        labels = [self._bold_label_rows[inner][1]]
        # Anidamiento real: una sección inmediatamente encima (bloque de
        # encabezados contiguo) es la sección "padre".
        outer = self._nearest_section_row(inner)
        if outer and inner - outer <= 2:
            labels.insert(0, self._bold_label_rows[outer][1])
        return labels

    def _leaf_row_labels(self, row: int, col: int) -> tuple[list[str], list[str]]:
        labels: list[str] = []
        notes: list[str] = []
        saw_number_after_text = False
        for cc in range(1, col):
            if self.is_text(row, cc):
                if saw_number_after_text:
                    notes.append(f"fila {row}: número intercalado entre rótulos de texto")
                t = self.text(row, cc)
                if not labels or labels[-1] != t:
                    labels.append(t)
            elif self.is_number(row, cc) and labels:
                saw_number_after_text = True
        return labels, notes

    def _group_header(self, row: int, col: int, section_row: int) -> list[str]:
        """Nivel de fila intermedio: el encabezado de grupo MÁS CERCANO por encima
        (``A.4. …``), sin cruzar otro grupo ni la sección."""
        label_col = next((cc for cc in range(1, col) if self.is_text(row, cc)), 2)
        for rr in range(row - 1, section_row, -1):
            if self._is_wide_header_row(rr) or rr in self._section_rows:
                continue
            if rr in self._group_rows:
                return [self._bold_label_rows[rr][1]]
            if self.is_text(rr, label_col) and _GROUP_HEADER_RE.match(self.text(rr, label_col)):
                return [self.text(rr, label_col)]
        return []

    def row_labels(self, row: int, col: int) -> tuple[list[str], list[str]]:
        section_row = self._nearest_section_row(row)
        group = self._group_header(row, col, section_row)
        leaf, notes = self._leaf_row_labels(row, col)
        return group + leaf, notes

    def column_labels(self, row: int, col: int) -> tuple[list[str], list[str]]:
        labels: list[str] = []
        notes: list[str] = []
        seen = False
        for rr in range(row - 1, 0, -1):
            if rr in self._section_rows:
                break  # la sección no es un encabezado de columna
            if self.is_text(rr, col):
                t = self.text(rr, col)
                if not labels or labels[0] != t:
                    labels.insert(0, t)
                seen = True
                continue
            if self.is_number(rr, col):
                if seen:
                    break
                continue
            # celda vacía en la columna de la métrica
            if seen:
                if self.is_text(rr, col + 1) or (col > 1 and self.is_text(rr, col - 1)):
                    notes.append(
                        f"fila {rr}: banda de encabezados continúa en columna vecina "
                        f"pero la columna de la métrica está vacía"
                    )
                break
        return labels, notes

    def cell_context(self, coord: str) -> CellContext:
        match = _COORD_RE.match(coord)
        if not match:
            raise ValueError(f"coordenada no reconocida: {coord!r}")
        col = _col_to_index(match.group(1))
        row = int(match.group(2))

        row_labels, row_notes = self.row_labels(row, col)
        column_labels, col_notes = self.column_labels(row, col)
        section_labels = self.section_labels(row, col)

        row_labels = row_labels[:_MAX_ROW_LEVELS]
        column_labels = column_labels[:_MAX_COLUMN_LEVELS]
        section_labels = section_labels[:_MAX_SECTION_LEVELS]
        notes = tuple(dict.fromkeys([*row_notes, *col_notes]))

        if not row_labels and not column_labels and not section_labels:
            status = CONTEXT_NO_CONTEXT
        elif notes:
            status = CONTEXT_AMBIGUOUS
        elif row_labels and column_labels:
            status = CONTEXT_COMPLETE
        else:
            status = CONTEXT_PARTIAL

        return CellContext(
            sheet=self.title,
            coord=coord,
            row_labels_raw=tuple(row_labels),
            column_labels_raw=tuple(column_labels),
            section_labels_raw=tuple(section_labels),
            context_status=status,
            notes=notes,
        )

    def text_cells(self) -> list[tuple[int, int, str]]:
        """Todas las celdas con texto resuelto (para volcados de layout)."""
        out: list[tuple[int, int, str]] = []
        for row in range(1, self.max_row + 1):
            for col in range(1, self.max_col + 1):
                if self.is_text(row, col):
                    out.append((row, col, self.text(row, col)))
        return out


# ---------------------------------------------------------------------------
# Evidencia técnica: rótulo de edad y criterios de fórmula
# ---------------------------------------------------------------------------

_AGE_PATTERNS: tuple[tuple[re.Pattern[str], object], ...] = (
    (re.compile(r"^MENOS DE 1 ANO\s*-\s*1 ANO$"), lambda m: (0, 1)),
    (re.compile(r"^MENOS DE\s*(\d+)\s*ANOS?$"), lambda m: (None, int(m.group(1)) - 1)),
    (re.compile(r"(\d+)\s*Y\s*MAS\s*ANOS?"), lambda m: (int(m.group(1)), None)),
    (re.compile(r"^(\d+)\s*[-A]\s*(\d+)\s*ANOS?$"), lambda m: (int(m.group(1)), int(m.group(2)))),
    (re.compile(r"^(\d+)\s*-\s*(\d+)$"), lambda m: (int(m.group(1)), int(m.group(2)))),
    (re.compile(r"^(\d+)\s*ANOS?$"), lambda m: (int(m.group(1)), int(m.group(1)))),
)

_AF_CRITERION_RE = re.compile(
    r"\$?AF\$?:\$?AF\s*,\s*\"\s*(>=|<=|>|<|=)\s*(\d+)\s*\"", re.IGNORECASE
)
_WILDCARD_LITERAL_RE = re.compile(r"\"([^\"]*\*[^\"]*)\"")


def parse_age_label(text: object) -> tuple[int | None, int | None] | None:
    """``(lower, upper)`` de un rótulo de edad, o ``None`` si no aplica.

    Evidencia técnica, no un mapping. ``lower``/``upper`` son inclusivos;
    ``None`` = extremo abierto (``75 y más años`` -> ``(75, None)``).
    """
    norm = normalize_semantic_label(text)
    if not norm:
        return None
    for pattern, build in _AGE_PATTERNS:
        match = pattern.search(norm)
        if match:
            return build(match)  # type: ignore[operator]
    return None


def age_bounds_from_formula(formula: object) -> tuple[int | None, int | None] | None:
    """``(lower, upper)`` a partir de los criterios ``$AF:$AF`` de la fórmula.

    ``>=N``/``>N`` fijan cota inferior; ``<=N``/``<N`` cota superior; ``=N`` ambas.
    Devuelve ``None`` si la fórmula no filtra por edad.
    """
    if not isinstance(formula, str):
        return None
    lower: int | None = None
    upper: int | None = None
    found = False
    for op, raw in _AF_CRITERION_RE.findall(formula):
        found = True
        n = int(raw)
        if op == ">=":
            lower = n if lower is None else max(lower, n)
        elif op == ">":
            lower = n + 1 if lower is None else max(lower, n + 1)
        elif op == "<=":
            upper = n if upper is None else min(upper, n)
        elif op == "<":
            upper = n - 1 if upper is None else min(upper, n - 1)
        elif op == "=":
            lower = n if lower is None else max(lower, n)
            upper = n if upper is None else min(upper, n)
    return (lower, upper) if found else None


def text_criteria_from_formula(formula: object) -> tuple[str, ...]:
    """Literales con comodín (``"*...*"``) usados como criterio en la fórmula."""
    if not isinstance(formula, str):
        return ()
    return tuple(dict.fromkeys(_WILDCARD_LITERAL_RE.findall(formula)))


CONSISTENCY_CONSISTENT = "CONSISTENT"
CONSISTENCY_CONFLICT = "CONFLICT"
CONSISTENCY_NO_COMPARABLE = "NO_COMPARABLE"


def age_intervals_overlap(
    a: tuple[int | None, int | None], b: tuple[int | None, int | None]
) -> bool:
    (a_lo, a_hi), (b_lo, b_hi) = a, b
    if a_lo is not None and b_hi is not None and a_lo > b_hi:
        return False
    return not (b_lo is not None and a_hi is not None and b_lo > a_hi)


def compare_label_and_formula_age(
    label_bounds: tuple[int | None, int | None] | None,
    formula_bounds: tuple[int | None, int | None] | None,
) -> str:
    """Comparación **objetiva** rótulo vs. fórmula (no interpreta clínicamente)."""
    if label_bounds is None or formula_bounds is None:
        return CONSISTENCY_NO_COMPARABLE
    return (
        CONSISTENCY_CONSISTENT
        if age_intervals_overlap(label_bounds, formula_bounds)
        else CONSISTENCY_CONFLICT
    )
