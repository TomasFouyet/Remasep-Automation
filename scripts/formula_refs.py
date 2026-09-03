"""Analizador ligero de referencias en fórmulas de Excel.

Trabaja SIEMPRE sobre el texto original de la fórmula (nunca sobre la versión
normalizada de ``formula_patterns``). No implementa la gramática completa de
Excel: reconoce referencias en notación A1 y las representa de forma simbólica;
los constructs dinámicos (``INDIRECT``, ``OFFSET``), las referencias externas y
las *structured references* de tabla se reportan como "no soportados" en lugar
de inferir nada silenciosamente.

Solo depende de la librería estándar.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field

# --- Tipos de referencia -----------------------------------------------------

CELL = "cell"
RANGE = "range"
WHOLE_COLUMN = "whole_column"
WHOLE_ROW = "whole_row"

# --- Gramática mínima (regex) ----------------------------------------------

# Literal de texto de Excel: "..."; las comillas internas se escriben "".
_STRING_RE = re.compile(r'"(?:[^"]|"")*"')

# Nombre de hoja: entre comillas simples ('' escapa una comilla) o token simple.
_SHEET = r"(?:'(?:[^']|'')+'|[A-Za-z_][A-Za-z0-9_.]*)"
_COL = r"\$?[A-Za-z]{1,3}"
_ROW = r"\$?[0-9]+"
_A1 = rf"{_COL}\$?[0-9]+"

# Una referencia: prefijo de hoja opcional + (rango A1 | rango de columnas |
# rango de filas | celda). El orden de las alternativas importa: primero las
# más largas para no partir "A1:B10" en "A1" y "B10".
_REF_RE = re.compile(
    rf"(?<![A-Za-z0-9_.!$\]])"
    rf"(?:(?P<sheet>{_SHEET})!)?"
    rf"(?:"
    rf"(?P<r1>{_A1}):(?P<r2>{_A1})"
    rf"|(?P<c1>{_COL}):(?P<c2>{_COL})"
    rf"|(?P<w1>{_ROW}):(?P<w2>{_ROW})"
    rf"|(?P<a>{_A1})"
    rf")"
    rf"(?![A-Za-z0-9_.(\[])"
)

_A1_SPLIT_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")

# Constructs que NO intentamos resolver; se reportan tal cual.
_UNSUPPORTED_PATTERNS: tuple[tuple[str, re.Pattern[str]], ...] = (
    ("indirect", re.compile(r"(?<![A-Za-z0-9_])INDIRECT\s*\(", re.IGNORECASE)),
    ("offset", re.compile(r"(?<![A-Za-z0-9_])OFFSET\s*\(", re.IGNORECASE)),
    ("external_reference", re.compile(r"\[[^\]]*\][A-Za-z0-9_. ]*!|'\[[^\]]*\]")),
    ("structured_reference", re.compile(r"[A-Za-z_][A-Za-z0-9_.]*\[[^\]]*\]|\[[@#][^\]]*\]")),
    ("three_d_reference", re.compile(rf"{_SHEET}:{_SHEET}!")),
    ("ref_error", re.compile(r"#REF!")),
    ("name_error", re.compile(r"#NAME\?")),
)


@dataclass(frozen=True)
class Reference:
    """Una referencia simbólica. Los rangos NO se expanden a celdas."""

    raw: str
    sheet: str | None  # None => misma hoja (referencia local)
    ref_type: str
    col_start: str | None  # letras, sin '$'
    col_end: str | None
    row_start: int | None
    row_end: int | None


@dataclass
class ParseResult:
    references: list[Reference] = field(default_factory=list)
    unsupported: list[str] = field(default_factory=list)


# --- Utilidades de columna (puras) ----------------------------------------


def col_to_index(col: str) -> int:
    """'A' -> 1, 'Z' -> 26, 'AA' -> 27."""
    index = 0
    for char in col.upper():
        index = index * 26 + (ord(char) - 64)
    return index


def index_to_col(index: int) -> str:
    """1 -> 'A', 27 -> 'AA'."""
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


# --- Analizador ----------------------------------------------------------------


def _strip_strings(formula: str) -> str:
    """Reemplaza los literales "..." por espacios de igual longitud."""
    return _STRING_RE.sub(lambda m: " " * len(m.group(0)), formula)


def _unquote_sheet(name: str | None) -> str | None:
    if name is None:
        return None
    if len(name) >= 2 and name[0] == "'" and name[-1] == "'":
        return name[1:-1].replace("''", "'")
    return name


def _split_a1(token: str) -> tuple[str, int]:
    match = _A1_SPLIT_RE.match(token)
    assert match is not None  # el token ya casó contra _A1
    return match.group(1).upper(), int(match.group(2))


def _norm_col(token: str) -> str:
    return token.replace("$", "").upper()


def _ordered(start: str, end: str) -> tuple[str, str]:
    return (start, end) if col_to_index(start) <= col_to_index(end) else (end, start)


def detect_unsupported(formula: str) -> list[str]:
    """Constructs no soportados presentes en la fórmula (orden estable, sin repetir)."""
    body = _strip_strings(formula)
    found: list[str] = []
    for label, pattern in _UNSUPPORTED_PATTERNS:
        if label not in found and pattern.search(body):
            found.append(label)
    return found


def parse_references(formula: str) -> ParseResult:
    """Extrae referencias simbólicas únicas de una fórmula (texto original).

    - Ignora todo lo que aparezca dentro de literales ``"..."``.
    - Deduplica referencias equivalentes (``=A1+A1`` -> una sola).
    - No expande rangos; ``A:A`` se mantiene simbólico.
    """
    result = ParseResult()
    if not isinstance(formula, str) or not formula:
        return result

    body = _strip_strings(formula)
    result.unsupported = detect_unsupported(formula)

    seen: set[tuple[str | None, str, str | None, str | None, int | None, int | None]] = set()
    for match in _REF_RE.finditer(body):
        sheet = _unquote_sheet(match.group("sheet"))

        if match.group("r1") is not None:
            c1, r1 = _split_a1(match.group("r1"))
            c2, r2 = _split_a1(match.group("r2"))
            col_start, col_end = _ordered(c1, c2)
            row_start, row_end = min(r1, r2), max(r1, r2)
            ref_type = CELL if (col_start == col_end and row_start == row_end) else RANGE
        elif match.group("c1") is not None:
            col_start, col_end = _ordered(_norm_col(match.group("c1")), _norm_col(match.group("c2")))
            row_start = row_end = None
            ref_type = WHOLE_COLUMN
        elif match.group("w1") is not None:
            col_start = col_end = None
            a, b = int(match.group("w1").replace("$", "")), int(match.group("w2").replace("$", ""))
            row_start, row_end = min(a, b), max(a, b)
            ref_type = WHOLE_ROW
        else:
            col, row = _split_a1(match.group("a"))
            col_start = col_end = col
            row_start = row_end = row
            ref_type = CELL

        key = (sheet, ref_type, col_start, col_end, row_start, row_end)
        if key in seen:
            continue
        seen.add(key)
        result.references.append(
            Reference(
                raw=match.group(0),
                sheet=sheet,
                ref_type=ref_type,
                col_start=col_start,
                col_end=col_end,
                row_start=row_start,
                row_end=row_end,
            )
        )

    return result
