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


def _resolve_criterion_expr(
    expr: str, resolve_ref: Callable[[str], object]
) -> tuple[str, tuple[str, ...]]:
    """Resuelve un criterio ``"..."[&$A1&...]`` a texto; devuelve también las
    referencias de celda usadas **como criterio** (no son dependencias de valor)."""
    parts = _split_top_level(expr.strip(), "&")
    resolved: list[str] = []
    criterion_refs: list[str] = []
    for part in parts:
        token = part.strip()
        if len(token) >= 2 and token[0] == '"' and token[-1] == '"':
            resolved.append(_unquote(token))
        elif _CELL_REF_RE.match(token):
            value = resolve_ref(token)
            if value is None:
                raise UnsupportedFormulaError(f"referencia de criterio sin valor: {token}")
            resolved.append(str(value))
            criterion_refs.append(token.replace("$", "").upper())
        else:
            raise UnsupportedFormulaError(f"parte de criterio no soportada: {token!r}")
    return "".join(resolved), tuple(criterion_refs)


def _parse_count_call(
    func: str, content: str, detail_sheet: str, resolve_ref: Callable[[str], object]
) -> tuple[CountTerm, tuple[str, ...]]:
    args = [a.strip() for a in _split_top_level(content, ",")]
    if func.upper() == "COUNTIF" and len(args) != 2:
        raise UnsupportedFormulaError("COUNTIF con un nº de argumentos inesperado")
    if func.upper() == "COUNTIFS" and (len(args) < 2 or len(args) % 2 != 0):
        raise UnsupportedFormulaError("COUNTIFS con un nº de argumentos impar")

    pairs: list[tuple[str, Criterion]] = []
    criterion_refs: list[str] = []
    for i in range(0, len(args), 2):
        column = _parse_range(args[i], detail_sheet)
        text, refs = _resolve_criterion_expr(args[i + 1], resolve_ref)
        pairs.append((column, build_criterion(text)))
        criterion_refs.extend(refs)
    return CountTerm(tuple(pairs)), tuple(criterion_refs)


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
        term, _refs = _parse_count_call(match.group(1), content, detail_sheet, resolve_ref)
        terms.append(term)
        columns.extend(column for column, _ in term.pairs)

    if not terms:
        raise UnsupportedFormulaError("sin términos COUNTIF/COUNTIFS")

    ordered_columns = tuple(sorted(set(columns)))
    return AggregationFormula(tuple(terms), ordered_columns)


# ---------------------------------------------------------------------------
# Fórmulas derivadas — aritmética (+/-) entre agregaciones y celdas agregadas
# ---------------------------------------------------------------------------

_VALUE_CELL_REF_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+)$")

# Funciones que denotan una comprobación/validación, no una agregación.
_VALIDATION_FUNCS = frozenset(
    {"IF", "IFS", "AND", "OR", "NOT", "IFERROR", "IFNA", "ISBLANK", "ISERROR", "ISNA"}
)


def _split_signed_terms(text: str) -> list[tuple[int, str]]:
    """Parte una expresión en términos ``(signo, texto)`` separados por ``+``/``-``
    de nivel superior (respeta strings y paréntesis). El signo inicial es ``+``."""
    terms: list[tuple[int, str]] = []
    buf: list[str] = []
    sign = 1
    depth = 0
    in_str = False
    i = 0
    while i < len(text):
        ch = text[i]
        if in_str:
            buf.append(ch)
            if ch == '"':
                if i + 1 < len(text) and text[i + 1] == '"':
                    buf.append('"')
                    i += 2
                    continue
                in_str = False
        elif ch == '"':
            in_str = True
            buf.append(ch)
        elif ch in "([{":
            depth += 1
            buf.append(ch)
        elif ch in ")]}":
            depth -= 1
            buf.append(ch)
        elif depth == 0 and ch in "+-":
            term = "".join(buf).strip()
            if term:
                terms.append((sign, term))
                buf = []
                sign = 1 if ch == "+" else -1
            else:
                sign = sign * (1 if ch == "+" else -1)
        else:
            buf.append(ch)
        i += 1
    last = "".join(buf).strip()
    if last:
        terms.append((sign, last))
    return terms


def arithmetic_constructs(formula: str) -> tuple[str, ...]:
    """Constructs aritméticos presentes en la fórmula, como etiquetas estables
    (``function:SUM``, ``operator:*``, ``grouping_parens``). ``COUNTIF``/``COUNTIFS``
    y los operadores ``+``/``-`` **no** se reportan (sí están soportados)."""
    body = formula.strip()
    body = body.removeprefix("=")
    nostr = re.sub(r'"(?:[^"]|"")*"', "", body)
    found: list[str] = []
    for match in re.finditer(r"(?<![A-Za-z0-9_.])(?:_xl\w+\.)?([A-Za-z][A-Za-z0-9_.]*)\s*\(", nostr):
        name = match.group(1).upper()
        if name in ("COUNTIF", "COUNTIFS"):
            continue
        found.append(f"function:{name}")
    depth = 0
    for idx, ch in enumerate(nostr):
        if ch == "(":
            prev = nostr[:idx].rstrip()
            if depth == 0 and not (prev and (prev[-1].isalnum() or prev[-1] in "._")):
                found.append("grouping_parens")
            depth += 1
        elif ch == ")":
            depth -= 1
        elif depth == 0 and ch == "*":
            found.append("operator:*")
        elif depth == 0 and ch == "/":
            found.append("operator:/")
        elif depth == 0 and ch == "&":
            found.append("operator:&")
    return tuple(dict.fromkeys(found))


@dataclass(frozen=True)
class DerivedFormula:
    """``=Σ COUNTIF(S)(...)  ±  celdas agregadas  ±  literales`` (solo ``+``/``-``).

    - ``count_terms``: agregaciones sobre la hoja de detalle (con su signo);
    - ``value_refs``: referencias de celda usadas **como valor** → dependencias
      del grafo (``(signo, coord)``);
    - ``literals``: constantes numéricas (``(signo, valor)``);
    - ``criterion_refs``: referencias usadas **como criterio** dentro de un
      ``COUNTIF(S)`` (se resuelven a texto, **no** son dependencias de valor).
    """

    count_terms: tuple[tuple[int, CountTerm], ...]
    value_refs: tuple[tuple[int, str], ...]
    literals: tuple[tuple[int, float], ...]
    referenced_columns: tuple[str, ...]
    criterion_refs: tuple[str, ...]

    @property
    def base_call_count(self) -> int:
        return len(self.count_terms)

    @property
    def is_pure_aggregation(self) -> bool:
        return bool(self.count_terms) and not self.value_refs and not self.literals

    @property
    def is_derived_aggregation(self) -> bool:
        return bool(self.count_terms) and bool(self.value_refs or self.literals)

    @property
    def depends_on_age_local(self) -> bool:
        return AGE_COLUMN in self.referenced_columns

    @property
    def value_ref_coords(self) -> tuple[str, ...]:
        return tuple(dict.fromkeys(coord for _sign, coord in self.value_refs))

    def evaluate(
        self, rows: Sequence[Mapping[str, object]], value_lookup: Mapping[str, float]
    ) -> int | float:
        total: float = 0.0
        for sign, term in self.count_terms:
            total += sign * term.count(rows)
        for sign, coord in self.value_refs:
            if coord not in value_lookup:
                raise UnsupportedFormulaError(f"referencia de valor sin resolver: {coord}")
            total += sign * value_lookup[coord]
        for sign, literal in self.literals:
            total += sign * literal
        return int(total) if float(total).is_integer() else total


def parse_derived_formula(
    formula: str,
    detail_sheet: str,
    resolve_ref: Callable[[str], object],
) -> DerivedFormula:
    """Parsea ``=COUNTIF(S)(...) ± celda ± ...`` (solo ``+``/``-``).

    Lanza :class:`UnsupportedFormulaError` ante ``* / & SUM IF``, paréntesis de
    agrupación aritmética, rangos como valor o cualquier otra función.
    """
    body = formula.strip()
    if body.startswith("="):
        body = body[1:].strip()
    if body.startswith("(") and body.endswith(")"):
        inner, end = _call_content(body, 0)
        if end == len(body) - 1:
            body = inner.strip()

    bad = arithmetic_constructs("=" + body)
    if bad:
        raise UnsupportedFormulaError("construct aritmético no soportado: " + ", ".join(bad))

    count_terms: list[tuple[int, CountTerm]] = []
    value_refs: list[tuple[int, str]] = []
    literals: list[tuple[int, float]] = []
    columns: list[str] = []
    criterion_refs: list[str] = []

    for sign, raw_term in _split_signed_terms(body):
        term_text = raw_term.strip()
        if not term_text:
            raise UnsupportedFormulaError("término aritmético vacío")
        match = _COUNT_CALL_RE.match(term_text)
        if match:
            content, end = _call_content(term_text, match.end() - 1)
            if term_text[end + 1 :].strip():
                raise UnsupportedFormulaError("hay contenido tras el COUNTIF(S)")
            term, refs = _parse_count_call(match.group(1), content, detail_sheet, resolve_ref)
            count_terms.append((sign, term))
            columns.extend(column for column, _ in term.pairs)
            criterion_refs.extend(refs)
            continue
        if _VALUE_CELL_REF_RE.match(term_text):
            value_refs.append((sign, term_text.replace("$", "").upper()))
            continue
        number = _to_number(term_text)
        if number is not None:
            literals.append((sign, number))
            continue
        raise UnsupportedFormulaError(f"término aritmético no soportado: {term_text[:40]!r}")

    if not count_terms and not value_refs:
        raise UnsupportedFormulaError("sin términos evaluables")

    return DerivedFormula(
        tuple(count_terms),
        tuple(value_refs),
        tuple(literals),
        tuple(sorted(set(columns))),
        tuple(dict.fromkeys(criterion_refs)),
    )


# ---------------------------------------------------------------------------
# SUM downstream (Sprint 2.4/2.5) — ``Σ SUM(...) ± celda ± literal``
# ---------------------------------------------------------------------------

_SUM_CALL_RE = re.compile(r"(?i)(?<![A-Za-z0-9_.])(?:_xl\w+\.)?(SUM)\s*\(")
_RANGE_CELL_RE = re.compile(r"^\$?([A-Za-z]{1,3})\$?([0-9]+):\$?([A-Za-z]{1,3})\$?([0-9]+)$")
_MAX_SUM_RANGE_CELLS = 5000


def _col_to_index(col: str) -> int:
    index = 0
    for ch in col.upper():
        index = index * 26 + (ord(ch) - 64)
    return index


def _index_to_col(index: int) -> str:
    letters = ""
    while index > 0:
        index, remainder = divmod(index - 1, 26)
        letters = chr(65 + remainder) + letters
    return letters


def _expand_rectangular_range(match: re.Match[str]) -> tuple[str, ...]:
    col1, row1, col2, row2 = match.group(1), int(match.group(2)), match.group(3), int(match.group(4))
    col_start, col_end = sorted((_col_to_index(col1), _col_to_index(col2)))
    row_start, row_end = sorted((row1, row2))
    if (col_end - col_start + 1) * (row_end - row_start + 1) > _MAX_SUM_RANGE_CELLS:
        raise UnsupportedFormulaError("rango SUM demasiado grande para expandir")
    return tuple(
        f"{_index_to_col(col)}{row}"
        for row in range(row_start, row_end + 1)
        for col in range(col_start, col_end + 1)
    )


def _parse_sum_content(content: str) -> tuple[str, ...]:
    text = content.strip()
    if not text:
        raise UnsupportedFormulaError("SUM sin argumentos")
    if "!" in text:
        raise UnsupportedFormulaError("SUM con referencia a otra hoja no soportado")
    match = _RANGE_CELL_RE.match(text)
    if match:
        return _expand_rectangular_range(match)
    coords: list[str] = []
    for sign, token in _split_signed_terms(text):
        if sign < 0:
            raise UnsupportedFormulaError("SUM con resta interna no soportado")
        token = token.strip()
        if not _VALUE_CELL_REF_RE.match(token):
            raise UnsupportedFormulaError(f"argumento de SUM no soportado: {token[:40]!r}")
        coords.append(token.replace("$", "").upper())
    if not coords:
        raise UnsupportedFormulaError("SUM sin celdas")
    return tuple(coords)


@dataclass(frozen=True)
class SumFormula:
    """``=Σ SUM(rango o cadena de celdas) ± celda ± literal`` (solo ``+``/``-``).

    Cada ``SUM(...)`` observado en el workbook real es, o bien un rango
    rectangular de la misma hoja (``SUM(Q86:Q92)``), o bien una cadena de
    celdas individuales unidas con ``+`` (``SUM(F40+H40+...+AL40)``). No se
    admite resta dentro del `SUM`, referencias a otra hoja ni más de un
    argumento — no observados en el workbook de referencia.
    """

    sum_terms: tuple[tuple[int, tuple[str, ...]], ...]
    value_refs: tuple[tuple[int, str], ...]
    literals: tuple[tuple[int, float], ...]

    @property
    def value_ref_coords(self) -> tuple[str, ...]:
        coords: list[str] = []
        for _sign, cells in self.sum_terms:
            coords.extend(cells)
        coords.extend(coord for _sign, coord in self.value_refs)
        return tuple(dict.fromkeys(coords))

    def evaluate(self, value_lookup: Mapping[str, object]) -> int | float:
        total = 0.0
        for sign, cells in self.sum_terms:
            for coord in cells:
                if coord not in value_lookup:
                    raise UnsupportedFormulaError(f"referencia de valor sin resolver: {coord}")
                number = _coerce_number(value_lookup[coord])
                total += sign * (number if number is not None else 0.0)
        for sign, coord in self.value_refs:
            if coord not in value_lookup:
                raise UnsupportedFormulaError(f"referencia de valor sin resolver: {coord}")
            number = _coerce_number(value_lookup[coord])
            total += sign * (number if number is not None else 0.0)
        for sign, literal in self.literals:
            total += sign * literal
        return int(total) if float(total).is_integer() else total


def parse_sum_formula(formula: str) -> SumFormula:
    """Parsea ``=SUM(...) ± celda ± ...`` (solo ``+``/``-``, sin ``eval()``).

    Lanza :class:`UnsupportedFormulaError` ante cualquier otro construct:
    ``* / &``, paréntesis de agrupación, otra función, rango no rectangular,
    referencia a otra hoja o resta dentro del `SUM`.
    """
    body = formula.strip()
    if body.startswith("="):
        body = body[1:].strip()
    if body.startswith("(") and body.endswith(")"):
        inner, end = _call_content(body, 0)
        if end == len(body) - 1:
            body = inner.strip()

    sum_terms: list[tuple[int, tuple[str, ...]]] = []
    value_refs: list[tuple[int, str]] = []
    literals: list[tuple[int, float]] = []

    for sign, raw_term in _split_signed_terms(body):
        term_text = raw_term.strip()
        if not term_text:
            raise UnsupportedFormulaError("término aritmético vacío")
        match = _SUM_CALL_RE.match(term_text)
        if match:
            content, end = _call_content(term_text, match.end() - 1)
            if term_text[end + 1 :].strip():
                raise UnsupportedFormulaError("hay contenido tras el SUM")
            sum_terms.append((sign, _parse_sum_content(content)))
            continue
        if _VALUE_CELL_REF_RE.match(term_text):
            value_refs.append((sign, term_text.replace("$", "").upper()))
            continue
        number = _to_number(term_text)
        if number is not None:
            literals.append((sign, number))
            continue
        raise UnsupportedFormulaError(f"término aritmético no soportado: {term_text[:40]!r}")

    if not sum_terms and not value_refs:
        raise UnsupportedFormulaError("sin términos SUM evaluables")

    return SumFormula(tuple(sum_terms), tuple(value_refs), tuple(literals))


# ---------------------------------------------------------------------------
# IF validation (Sprint 2.5) — comprobaciones legacy, no métricas clínicas
# ---------------------------------------------------------------------------

_IF_CALL_RE = re.compile(r"(?i)(?<![A-Za-z0-9_.])(?:_xl\w+\.)?(IF)\s*\(")
_IF_OPERATORS = ("<>", "<=", ">=", "<", ">", "=")


@dataclass(frozen=True)
class IfOperand:
    """Un lado de la condición de un ``IF``: celda, número o texto literal."""

    kind: str  # "CELL" | "NUM" | "STR"
    coord: str | None = None
    number: float | None = None
    text: str | None = None


@dataclass(frozen=True)
class IfBranch:
    """Rama ``then``/``else`` de un ``IF``: literal, celda o `IF` anidado."""

    kind: str  # "NUM" | "STR" | "CELL" | "NESTED"
    number: float | None = None
    text: str | None = None
    coord: str | None = None
    nested: IfFormula | None = None

    @property
    def value_ref_coords(self) -> tuple[str, ...]:
        if self.kind == "CELL":
            return (self.coord,) if self.coord else ()
        if self.kind == "NESTED" and self.nested is not None:
            return self.nested.value_ref_coords
        return ()

    def resolve(self, value_lookup: Mapping[str, object]) -> object:
        if self.kind == "NUM":
            return self.number
        if self.kind == "STR":
            return self.text
        if self.kind == "CELL":
            if self.coord not in value_lookup:
                raise UnsupportedFormulaError(f"referencia de valor sin resolver: {self.coord}")
            return value_lookup[self.coord]
        assert self.nested is not None
        return self.nested.evaluate(value_lookup)


def _operand_value(operand: IfOperand, value_lookup: Mapping[str, object]) -> object:
    if operand.kind == "NUM":
        return operand.number
    if operand.kind == "STR":
        return operand.text
    if operand.coord not in value_lookup:
        raise UnsupportedFormulaError(f"referencia de valor sin resolver: {operand.coord}")
    return value_lookup[operand.coord]


def _evaluate_if_condition(
    left_operand: IfOperand, operator: str, right_operand: IfOperand, value_lookup: Mapping[str, object]
) -> bool:
    left_value = _operand_value(left_operand, value_lookup)
    right_value = _operand_value(right_operand, value_lookup)
    text_mode = operator in ("=", "<>") and (left_operand.kind == "STR" or right_operand.kind == "STR")
    if text_mode:
        equal = normalize_legacy_text(left_value) == normalize_legacy_text(right_value)
        return equal if operator == "=" else not equal
    left_number = _coerce_number(left_value)
    right_number = _coerce_number(right_value)
    if left_number is None or right_number is None:
        raise UnsupportedFormulaError("condición IF no numérica ni de texto reconocible")
    return _compare_numeric(left_number, operator, right_number)


@dataclass(frozen=True)
class IfFormula:
    """``=IF(izquierda OP derecha, entonces, si_no)``, con ``entonces``/``si_no``
    literales, referencias de celda o un ``IF`` anidado (la única composición
    observada en el workbook real)."""

    left: IfOperand
    operator: str
    right: IfOperand
    then_branch: IfBranch
    else_branch: IfBranch

    @property
    def value_ref_coords(self) -> tuple[str, ...]:
        coords: list[str] = []
        for operand in (self.left, self.right):
            if operand.kind == "CELL" and operand.coord:
                coords.append(operand.coord)
        coords.extend(self.then_branch.value_ref_coords)
        coords.extend(self.else_branch.value_ref_coords)
        return tuple(dict.fromkeys(coords))

    def evaluate(self, value_lookup: Mapping[str, object]) -> object:
        condition = _evaluate_if_condition(self.left, self.operator, self.right, value_lookup)
        branch = self.then_branch if condition else self.else_branch
        result = branch.resolve(value_lookup)
        if isinstance(result, float) and result.is_integer():
            return int(result)
        return result


def _parse_if_operand(token: str) -> IfOperand:
    token = token.strip()
    if len(token) >= 2 and token[0] == '"' and token[-1] == '"':
        return IfOperand("STR", text=_unquote(token))
    if _VALUE_CELL_REF_RE.match(token):
        return IfOperand("CELL", coord=token.replace("$", "").upper())
    number = _to_number(token)
    if number is not None:
        return IfOperand("NUM", number=number)
    raise UnsupportedFormulaError(f"operando IF no soportado: {token[:40]!r}")


def _parse_if_condition(condition: str) -> tuple[IfOperand, str, IfOperand]:
    masked = re.sub(r'"(?:[^"]|"")*"', lambda m: " " * len(m.group(0)), condition)
    for operator in _IF_OPERATORS:
        index = masked.find(operator)
        if index != -1:
            left = _parse_if_operand(condition[:index])
            right = _parse_if_operand(condition[index + len(operator) :])
            return left, operator, right
    raise UnsupportedFormulaError(f"condición IF sin operador reconocido: {condition[:40]!r}")


def _parse_if_branch(token: str) -> IfBranch:
    token = token.strip()
    if _IF_CALL_RE.match(token):
        return IfBranch("NESTED", nested=parse_if_formula(token))
    if len(token) >= 2 and token[0] == '"' and token[-1] == '"':
        return IfBranch("STR", text=_unquote(token))
    if _VALUE_CELL_REF_RE.match(token):
        return IfBranch("CELL", coord=token.replace("$", "").upper())
    number = _to_number(token)
    if number is not None:
        return IfBranch("NUM", number=number)
    raise UnsupportedFormulaError(f"resultado IF no soportado: {token[:40]!r}")


def parse_if_formula(formula: str) -> IfFormula:
    """Parsea ``=IF(izq OP der, entonces, si_no)`` (``IF`` anidado permitido en
    una rama). Operadores soportados: ``= <> < <= > >=``. Sin ``eval()``.

    Lanza :class:`UnsupportedFormulaError` ante cualquier otra forma: más de un
    `IF` en la condición, funciones no soportadas, resultado que no sea
    literal/celda/`IF` anidado, etc.
    """
    body = formula.strip()
    if body.startswith("="):
        body = body[1:].strip()
    match = _IF_CALL_RE.match(body)
    if not match:
        raise UnsupportedFormulaError("no es una fórmula IF")
    content, end = _call_content(body, match.end() - 1)
    if body[end + 1 :].strip():
        raise UnsupportedFormulaError("hay contenido tras el IF")
    args = [a.strip() for a in _split_top_level(content, ",")]
    if len(args) != 3:
        raise UnsupportedFormulaError("IF con un nº de argumentos distinto de 3")
    left, operator, right = _parse_if_condition(args[0])
    then_branch = _parse_if_branch(args[1])
    else_branch = _parse_if_branch(args[2])
    return IfFormula(left, operator, right, then_branch, else_branch)


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
