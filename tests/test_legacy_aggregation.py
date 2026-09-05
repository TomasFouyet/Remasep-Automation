"""Tests del evaluador de agregaciones legacy (Sprint 2.3).

Casos sintéticos, sin workbook. No se usa data/local.
"""

from __future__ import annotations

import pytest

from remasep.services.legacy_aggregation import (
    UnsupportedFormulaError,
    build_criterion,
    functions_used,
    padding_can_affect,
    parse_aggregation_formula,
)

DETAIL = "Atenciones - Detalles de citas"


def _no_refs(_ref: str) -> object:
    raise AssertionError("esta fórmula no debería resolver referencias de celda")


def _refs(mapping):
    return lambda ref: mapping.get(ref.replace("$", ""))


def _countif(col: str, crit: str) -> str:
    return f"=COUNTIF('{DETAIL}'!${col}:${col},{crit})"


# --- criterios de texto ------------------------------------------------


def test_countif_exact_text_case_insensitive():
    crit = build_criterion("Consulta General")
    assert crit.matches("CONSULTA GENERAL")
    assert crit.matches("consulta general")
    assert not crit.matches("consulta general x")


def test_countif_wildcard_star():
    crit = build_criterion("*consulta*")
    assert crit.matches("una consulta larga")
    assert crit.matches("consulta")
    assert not crit.matches("cnslt")


def test_countif_wildcard_question():
    crit = build_criterion("a?c")
    assert crit.matches("abc")
    assert crit.matches("aXc")
    assert not crit.matches("ac")
    assert not crit.matches("abbc")


def test_criteria_are_accent_sensitive():
    crit = build_criterion("*máscara*")
    assert crit.matches("con máscara puesta")
    assert not crit.matches("con mascara puesta")


def test_criteria_are_whitespace_exact():
    crit = build_criterion("control aparato removible")
    assert crit.matches("CONTROL APARATO REMOVIBLE")
    assert not crit.matches("control  aparato removible")  # doble espacio
    assert not crit.matches(" control aparato removible")  # espacio inicial


def test_tilde_escape_wildcard():
    crit = build_criterion("cod~*fin")  # ~* -> asterisco literal
    assert crit.matches("cod*fin")
    assert not crit.matches("cod XYZ fin")


def test_not_equal_operator():
    crit = build_criterion("<>X")
    assert crit.matches("Y")
    assert not crit.matches("x")  # case-insensitive


# --- criterios numéricos --------------------------------------------


@pytest.mark.parametrize(
    ("criterion", "value", "expected"),
    [
        (">=10", 10, True), (">=10", 9, False),
        ("<=14", 14, True), ("<=14", 15, False),
        (">2", 3, True), (">2", 2, False),
        ("<2", 1, True), ("<2", 2, False),
        ("=5", 5, True), ("=5", 6, False),
        ("<>7", 8, True), ("<>7", 7, False),
        (">=8", "8", True),  # texto numérico
        (">=8", "ocho", False),  # texto no numérico -> no cuenta
        (">=8", None, False),
    ],
)
def test_numeric_criteria(criterion, value, expected):
    assert build_criterion(criterion).matches(value) is expected


# --- parsing de fórmulas -----------------------------------------------


def test_parse_countif_single():
    formula = _countif("AG", '"*consulta médico*"')
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    assert parsed.referenced_columns == ("AG",)
    assert not parsed.depends_on_age
    assert len(parsed.terms) == 1


def test_parse_countifs_multiple_criteria():
    formula = (
        f"=COUNTIFS('{DETAIL}'!$AG:$AG,\"*consulta médico*\","
        f"'{DETAIL}'!$AF:$AF,\">=10\",'{DETAIL}'!$AF:$AF,\"<=14\")"
    )
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    assert parsed.referenced_columns == ("AF", "AG")
    assert parsed.depends_on_age
    assert len(parsed.terms[0].pairs) == 3


def test_parse_sum_of_countifs():
    formula = _countif("AC", '"a"') + "+" + _countif("AC", '"b"').removeprefix("=")
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    assert len(parsed.terms) == 2


def test_parse_criterion_concatenation():
    formula = _countif("AG", '"*consulta médico*"&"Hombre"')
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    crit = parsed.terms[0].pairs[0][1]
    assert crit.matches("consulta médicoHombre")
    assert not crit.matches("consulta médicoMujer")


def test_parse_criterion_cell_reference():
    formula = (
        f"=COUNTIFS('{DETAIL}'!$O:$O,$A5,'{DETAIL}'!$AF:$AF,\"<2\")"
    )
    parsed = parse_aggregation_formula(formula, DETAIL, _refs({"A5": "CONSULTA MEDICA X"}))
    crit = parsed.terms[0].pairs[0][1]
    assert crit.matches("consulta medica x")


def test_parse_criterion_cell_reference_with_concat():
    formula = _countif("AD", '$A5&"Hombre"')
    parsed = parse_aggregation_formula(formula, DETAIL, _refs({"A5": "control odon esp"}))
    crit = parsed.terms[0].pairs[0][1]
    assert crit.matches("control odon espHombre")


# --- no soportado -----------------------------------------------------


def test_unsupported_other_function():
    with pytest.raises(UnsupportedFormulaError):
        parse_aggregation_formula(f"=SUMIF('{DETAIL}'!$A:$A,\"x\")", DETAIL, _no_refs)


def test_unsupported_trailing_expression():
    formula = _countif("AE", '"x"') + "-A1"
    with pytest.raises(UnsupportedFormulaError, match="tras el COUNTIF"):
        parse_aggregation_formula(formula, DETAIL, _no_refs)


def test_unsupported_non_whole_column_range():
    with pytest.raises(UnsupportedFormulaError, match="columna entera"):
        parse_aggregation_formula(f"=COUNTIF('{DETAIL}'!$A2:$A100,\"x\")", DETAIL, _no_refs)


def test_unsupported_range_not_detail_sheet():
    with pytest.raises(UnsupportedFormulaError, match="hoja de detalle"):
        parse_aggregation_formula("=COUNTIF('Otra'!$A:$A,\"x\")", DETAIL, _no_refs)


def test_unsupported_unresolved_cell_reference():
    with pytest.raises(UnsupportedFormulaError, match="sin valor"):
        parse_aggregation_formula(_countif("O", "$A5"), DETAIL, _refs({}))


# --- evaluación -----------------------------------------------------


def _rows(*specs):
    return [dict(s) for s in specs]


def test_evaluate_countif():
    formula = _countif("AG", '"*consulta médico*Hombre"')
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    rows = _rows(
        {"AG": "consulta médicoHombre"},
        {"AG": "consulta médicoHombre"},
        {"AG": "consulta médicoMujer"},
        {"AG": "#N/A"},
    )
    assert parsed.evaluate(rows) == 2


def test_evaluate_countifs_age_band():
    formula = (
        f"=COUNTIFS('{DETAIL}'!$AG:$AG,\"*consulta médico*\","
        f"'{DETAIL}'!$AF:$AF,\">=10\",'{DETAIL}'!$AF:$AF,\"<=14\")"
    )
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    rows = _rows(
        {"AG": "consulta médicoHombre", "AF": 10},
        {"AG": "consulta médicoHombre", "AF": 14},
        {"AG": "consulta médicoHombre", "AF": 15},  # fuera de banda
        {"AG": "consulta médicoHombre", "AF": 9},  # fuera de banda
        {"AG": "#N/A", "AF": 12},  # categoría no coincide
    )
    assert parsed.evaluate(rows) == 2


def test_evaluate_sum_of_countif():
    formula = _countif("AC", '"a"') + "+" + _countif("AC", '"b"').removeprefix("=")
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    rows = _rows({"AC": "a"}, {"AC": "a"}, {"AC": "b"}, {"AC": "c"})
    assert parsed.evaluate(rows) == 3


# --- padding sensitivity ------------------------------------------


_EMPTY = {
    "O": "", "AC": "", "AD": "", "AE": "", "AF": 0,
    "AG": "#N/A", "AH": "#N/A", "AI": "#N/A", "AJ": "0", "AK": "0", "AL": "0",
}


def test_padding_not_sensitive_with_category_filter():
    formula = (
        f"=COUNTIFS('{DETAIL}'!$AG:$AG,\"*consulta médico*\",'{DETAIL}'!$AF:$AF,\"<10\")"
    )
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    affect, reason = padding_can_affect(parsed, _EMPTY)
    assert affect is False
    assert "Ningún término" in reason


def test_padding_sensitive_age_only_formula():
    # una fórmula que SOLO filtra por AF: la fila vacía (AF=0) contaría.
    formula = _countif("AF", '"<10"')
    parsed = parse_aggregation_formula(formula, DETAIL, _no_refs)
    affect, reason = padding_can_affect(parsed, _EMPTY)
    assert affect is True
    assert "contaría una fila vacía" in reason


# --- functions_used --------------------------------------------------


def test_functions_used():
    assert functions_used(_countif("AG", '"x"')) == ("COUNTIF",)
    assert functions_used(f"=COUNTIFS('{DETAIL}'!$AG:$AG,\"x\")-A1") == ("COUNTIFS",)
