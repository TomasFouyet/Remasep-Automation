"""Evaluador del subset de fórmulas de agregación legacy (Sprint 2.3).

Cubre **únicamente** los constructs que aparecen realmente en
``GENERACION DATOS REMASEP.xlsx`` para las fórmulas que referencian directamente
``'Atenciones - Detalles de citas'``:

- ``COUNTIF(rango, criterio)`` y ``COUNTIFS(r1, c1, r2, c2, ...)``;
- sumas de esas llamadas con ``+``;
- rango = columna entera de la hoja de detalle (``!$AG:$AG`` / ``!O:O``);
- criterio = concatenación con ``&`` de literales ``"..."`` y referencias de celda;
- criterios numéricos con ``< <= > >= = <>`` y criterios de texto con comodines.

Semántica de criterios de texto (consistente con Sprint 2.2): case-insensitive
(``casefold``), **con tildes**, **whitespace exacto**, comodines Excel (``*`` ``?``
``~`` como escape). NO se usa ``normalize_text`` global. No hay fuzzy matching.

Cualquier cosa fuera de este subset se marca ``UNSUPPORTED`` (no se evalúa a
medias, no se infiere).
"""

from __future__ import annotations

import math
import re
from collections.abc import Callable, Mapping, Sequence
from dataclasses import dataclass

from remasep.core.text import normalize_legacy_text

DERIVED_COLUMNS: tuple[str, ...] = ("AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL")
AGE_COLUMN = "AF"

_OPERATORS = ("<>", "<=", ">=", "<", ">", "=")
_COUNT_CALL_RE = re.compile(r"(?i)(?<![A-Za-z0-9_.])(?:_xl\w+\.)?(COUNTIFS?)\s*\(")
_WHOLE_COL_RANGE_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?:\$?([A-Za-z]{1,3})$")
_CELL_REF_RE = re.compile(r"^\$?[A-Za-z]{1,3}\$?[0-9]+$")


class UnsupportedFormulaError(Exception):
    """La fórmula usa un construct fuera del subset soportado."""


# ---------------------------------------------------------------------------
# Utilidades de parsing
# ---------------------------------------------------------------------------


def _split_top_level(text: str, separator: str) -> list[str]:
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
        elif ch == separator and depth == 0:
            parts.append("".join(buf))
            buf = []
        else:
            buf.append(ch)
        i += 1
    parts.append("".join(buf))
    return parts


def _call_content(text: str, open_paren_index: int) -> tuple[str, int]:
    depth = 0
    in_str = False
    j = open_paren_index
    while j < len(text):
        ch = text[j]
        if in_str:
            if ch == '"':
                if j + 1 < len(text) and text[j + 1] == '"':
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
                return text[open_paren_index + 1 : j], j
        j += 1
    raise UnsupportedFormulaError("paréntesis sin cerrar")


def _unquote(literal: str) -> str:
    inner = literal[1:-1]
    return inner.replace('""', '"')


def _to_number(text: str) -> float | None:
    text = text.strip()
    if not text:
        return None
    try:
        value = float(text)
    except ValueError:
        return None
    return value if math.isfinite(value) else None


# ---------------------------------------------------------------------------
# Criterio Excel
# ---------------------------------------------------------------------------


def _wildcard_to_regex(pattern: str) -> re.Pattern[str]:
    out: list[str] = ["\\A"]
    i = 0
    while i < len(pattern):
        ch = pattern[i]
        if ch == "~" and i + 1 < len(pattern) and pattern[i + 1] in "*?~":
            out.append(re.escape(pattern[i + 1]))
            i += 2
            continue
        if ch == "*":
            out.append(".*")
        elif ch == "?":
            out.append(".")
        else:
            out.append(re.escape(ch))
        i += 1
    out.append("\\Z")
    return re.compile("".join(out), re.DOTALL)


@dataclass(frozen=True)
class Criterion:
    raw: str
    operator: str  # "" (igualdad) | "<" | "<=" | ">" | ">=" | "=" | "<>"
    numeric: float | None
    text_cf: str | None  # texto casefolded para igualdad exacta
    regex: re.Pattern[str] | None  # si el texto tiene comodines

    def matches(self, value: object) -> bool:
        if self.numeric is not None:
            actual = _coerce_number(value)
            if actual is None:
                return False
            return _compare_numeric(actual, self.operator, self.numeric)

        text = normalize_legacy_text(value)
        if self.operator == "<>":
            return not self._text_equal(text)
        return self._text_equal(text)

    def _text_equal(self, text_cf: str) -> bool:
        if self.regex is not None:
            return self.regex.match(text_cf) is not None
        return text_cf == self.text_cf


def _coerce_number(value: object) -> float | None:
    if isinstance(value, bool):
        return None
    if isinstance(value, (int, float)):
        return float(value) if math.isfinite(float(value)) else None
    if isinstance(value, str):
        return _to_number(value)
    return None


def _compare_numeric(actual: float, operator: str, expected: float) -> bool:
    if operator in ("", "="):
        return actual == expected
    if operator == "<>":
        return actual != expected
    if operator == "<":
        return actual < expected
    if operator == "<=":
        return actual <= expected
    if operator == ">":
        return actual > expected
    if operator == ">=":
        return actual >= expected
    raise UnsupportedFormulaError(f"operador no soportado: {operator!r}")


def build_criterion(resolved_text: str) -> Criterion:
    operator = ""
    body = resolved_text
    for candidate in _OPERATORS:
        if body.startswith(candidate):
            operator = candidate
            body = body[len(candidate) :]
            break

    number = None if any(w in body for w in ("*", "?")) else _to_number(body)
    if number is not None:
        return Criterion(resolved_text, operator, number, None, None)

    if operator in ("<", "<=", ">", ">="):
        raise UnsupportedFormulaError(
            f"criterio de comparación no numérico no soportado: {resolved_text!r}"
        )

    if any(ch in body for ch in ("*", "?", "~")):
        return Criterion(resolved_text, operator, None, None, _wildcard_to_regex(body.casefold()))
    return Criterion(resolved_text, operator, None, body.casefold(), None)


# ---------------------------------------------------------------------------
# Fórmula de agregación
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class CountTerm:
    pairs: tuple[tuple[str, Criterion], ...]  # (columna_hoja_detalle, criterio)

    def count(self, rows: Sequence[Mapping[str, object]]) -> int:
        total = 0
        for row in rows:
            if all(criterion.matches(row.get(column)) for column, criterion in self.pairs):
                total += 1
        return total


@dataclass(frozen=True)
class AggregationFormula:
    terms: tuple[CountTerm, ...]
    referenced_columns: tuple[str, ...]

    @property
    def depends_on_age(self) -> bool:
        return AGE_COLUMN in self.referenced_columns

    def evaluate(self, rows: Sequence[Mapping[str, object]]) -> int:
        return sum(term.count(rows) for term in self.terms)


def _resolve_criterion_expr(expr: str, resolve_ref: Callable[[str], object]) -> str:
    parts = _split_top_level(expr.strip(), "&")
    resolved: list[str] = []
    for part in parts:
        token = part.strip()
        if len(token) >= 2 and token[0] == '"' and token[-1] == '"':
            resolved.append(_unquote(token))
        elif _CELL_REF_RE.match(token):
            value = resolve_ref(token)
            if value is None:
                raise UnsupportedFormulaError(f"referencia de criterio sin valor: {token}")
            resolved.append(str(value))
        else:
            raise UnsupportedFormulaError(f"parte de criterio no soportada: {token!r}")
    return "".join(resolved)


def _parse_count_call(
    func: str, content: str, detail_sheet: str, resolve_ref: Callable[[str], object]
) -> CountTerm:
    args = [a.strip() for a in _split_top_level(content, ",")]
    if func.upper() == "COUNTIF" and len(args) != 2:
        raise UnsupportedFormulaError("COUNTIF con un nº de argumentos inesperado")
    if func.upper() == "COUNTIFS" and (len(args) < 2 or len(args) % 2 != 0):
        raise UnsupportedFormulaError("COUNTIFS con un nº de argumentos impar")

    pairs: list[tuple[str, Criterion]] = []
    for i in range(0, len(args), 2):
        column = _parse_range(args[i], detail_sheet)
        criterion = build_criterion(_resolve_criterion_expr(args[i + 1], resolve_ref))
        pairs.append((column, criterion))
    return CountTerm(tuple(pairs))


def _parse_range(range_expr: str, detail_sheet: str) -> str:
    text = range_expr.strip()
    if "!" not in text:
        raise UnsupportedFormulaError(f"rango sin hoja: {text!r}")
    sheet_part, _, cells = text.rpartition("!")
    sheet_name = sheet_part.strip()
    if sheet_name.startswith("'") and sheet_name.endswith("'"):
        sheet_name = sheet_name[1:-1].replace("''", "'")
    if sheet_name != detail_sheet:
        raise UnsupportedFormulaError(f"rango que no apunta a la hoja de detalle: {sheet_name!r}")
    match = _WHOLE_COL_RANGE_RE.match(cells.strip())
    if not match or match.group(1).upper() != match.group(2).upper():
        raise UnsupportedFormulaError(f"rango que no es columna entera: {cells!r}")
    return match.group(1).upper()


def parse_aggregation_formula(
    formula: str,
    detail_sheet: str,
    resolve_ref: Callable[[str], object],
) -> AggregationFormula:
    """Parsea ``=COUNTIF(...)[+COUNTIFS(...)...]`` sobre la hoja de detalle.

    Lanza :class:`UnsupportedFormulaError` ante cualquier construct fuera del subset.
    """
    body = formula.strip()
    if body.startswith("="):
        body = body[1:].strip()
    if body.startswith("(") and body.endswith(")"):
        # solo si el paréntesis envuelve toda la expresión
        inner, end = _call_content(body, 0)
        if end == len(body) - 1:
            body = inner.strip()

    terms: list[CountTerm] = []
    columns: list[str] = []
    for raw_term in _split_top_level(body, "+"):
        term_text = raw_term.strip()
        match = _COUNT_CALL_RE.match(term_text)
        if not match:
            raise UnsupportedFormulaError(f"término no es COUNTIF/COUNTIFS: {term_text[:40]!r}")
        content, end = _call_content(term_text, match.end() - 1)
        if term_text[end + 1 :].strip():
            raise UnsupportedFormulaError("hay contenido tras el COUNTIF(S)")
        term = _parse_count_call(match.group(1), content, detail_sheet, resolve_ref)
        terms.append(term)
        columns.extend(column for column, _ in term.pairs)

    if not terms:
        raise UnsupportedFormulaError("sin términos COUNTIF/COUNTIFS")

    ordered_columns = tuple(sorted(set(columns)))
    return AggregationFormula(tuple(terms), ordered_columns)


# ---------------------------------------------------------------------------
# Sensibilidad al padding (filas estructuralmente vacías)
# ---------------------------------------------------------------------------


def padding_can_affect(
    formula: AggregationFormula, empty_row_values: Mapping[str, object]
) -> tuple[bool, str]:
    """¿Contaría alguna fila estructuralmente vacía en esta fórmula?

    Se evalúan los criterios contra los valores que Excel deja en las columnas
    referenciadas para una fila vacía. Si algún término cuenta la fila -> sí.
    """
    for term_index, term in enumerate(formula.terms):
        failing = [
            column
            for column, criterion in term.pairs
            if not criterion.matches(empty_row_values.get(column))
        ]
        if not failing:
            return True, (
                f"El término #{term_index + 1} contaría una fila vacía "
                f"({', '.join(f'{c}={_short(empty_row_values.get(c))}' for c, _ in term.pairs)})."
            )
    return False, "Ningún término cuenta una fila estructuralmente vacía."


def _short(value: object) -> str:
    text = "" if value is None else str(value)
    return repr(text if len(text) <= 20 else text[:17] + "…")


def _pattern_signature(formula: str) -> str:
    """Firma estable de la 'familia' de una fórmula (para agrupar/reportar)."""
    signature = re.sub(r'"(?:[^"]|"")*"', '"S"', formula)
    signature = re.sub(r"\$?[A-Za-z]{1,3}\$?[0-9]+", "REF", signature)
    signature = re.sub(r"[0-9]+", "#", signature)
    return re.sub(r"\s+", "", signature)


def formula_pattern(formula: str) -> str:
    return _pattern_signature(formula)


def functions_used(formula: str) -> tuple[str, ...]:
    found: set[str] = set()
    stripped = re.sub(r'"(?:[^"]|"")*"', "", formula)
    for match in re.finditer(r"(?<![A-Za-z0-9_.])(?:_xl\w+\.)?([A-Za-z][A-Za-z0-9_.]*)\s*\(", stripped):
        found.add(match.group(1).upper())
    return tuple(sorted(found))
