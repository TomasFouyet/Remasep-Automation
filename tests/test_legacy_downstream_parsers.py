"""Tests del subset SUM/IF de `legacy_aggregation` (Sprint 2.5).

Casos sintéticos, sin workbook. No se usa data/local.
"""

from __future__ import annotations

import pytest

from remasep.services.legacy_aggregation import (
    UnsupportedFormulaError,
    parse_if_formula,
    parse_sum_formula,
)

# --- SUM ---------------------------------------------------------------


def test_sum_of_rectangular_range():
    parsed = parse_sum_formula("=SUM(F10:F12)")
    assert parsed.value_ref_coords == ("F10", "F11", "F12")
    assert parsed.evaluate({"F10": 1, "F11": 2, "F12": 3}) == 6


def test_sum_of_individual_cell_chain():
    parsed = parse_sum_formula("=SUM(F40+H40+J40)")
    assert parsed.sum_terms == ((1, ("F40", "H40", "J40")),)
    assert parsed.evaluate({"F40": 1, "H40": 2, "J40": 3}) == 6


def test_sum_range_plus_trailing_cell():
    parsed = parse_sum_formula("=SUM(Q86:Q88)+Q84")
    assert set(parsed.value_ref_coords) == {"Q86", "Q87", "Q88", "Q84"}
    assert parsed.evaluate({"Q86": 1, "Q87": 1, "Q88": 1, "Q84": 10}) == 13


def test_sum_of_single_cell_models_sum_of_sum():
    # SUM(P10) donde P10 es a su vez un SUM: la evaluación es composicional.
    parsed = parse_sum_formula("=SUM(P10)")
    assert parsed.value_ref_coords == ("P10",)
    assert parsed.evaluate({"P10": 42}) == 42


def test_sum_treats_blank_range_cell_as_zero():
    parsed = parse_sum_formula("=SUM(F10:F12)")
    # F11 no está en el lookup con valor -> se considera 0 (celda en blanco), no
    # "dependencia no disponible": eso lo decide quien arma el lookup, no el
    # evaluador (ver DEPENDENCY_UNAVAILABLE en evaluate_legacy_downstream).
    assert parsed.evaluate({"F10": 1, "F11": None, "F12": 3}) == 4


def test_sum_raises_when_dependency_missing_from_lookup():
    parsed = parse_sum_formula("=SUM(F10:F11)")
    with pytest.raises(UnsupportedFormulaError, match="sin resolver"):
        parsed.evaluate({"F10": 1})


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(F30*2)",           # operador no soportado dentro del SUM
        "=SUM(A1,B1)",           # argumentos separados por coma
        "=SUM('Otra'!A1:A2)",    # otra hoja
        "=SUM(A1:A2:A3)",        # rango mal formado
    ],
)
def test_unsupported_sum_variants(formula):
    with pytest.raises(UnsupportedFormulaError):
        parse_sum_formula(formula)


def test_sum_minus_cell_is_supported():
    parsed = parse_sum_formula("=SUM(A1:A2)-B1")
    assert parsed.value_refs == ((-1, "B1"),)
    assert parsed.evaluate({"A1": 5, "A2": 5, "B1": 3}) == 7


def test_sum_rejects_trailing_content_after_call():
    with pytest.raises(UnsupportedFormulaError, match="tras el SUM"):
        parse_sum_formula("=SUM(F10:F11)*2")


# --- IF: operadores -----------------------------------------------------


@pytest.mark.parametrize(
    ("operator", "left", "right", "expected"),
    [
        ("<", 1, 2, True), ("<", 2, 1, False),
        ("<=", 2, 2, True), ("<=", 3, 2, False),
        (">", 3, 2, True), (">", 2, 3, False),
        (">=", 2, 2, True), (">=", 1, 2, False),
        ("=", 5, 5, True), ("=", 5, 6, False),
        ("<>", 5, 6, True), ("<>", 5, 5, False),
    ],
)
def test_if_reference_vs_reference_operators(operator, left, right, expected):
    formula = f"=IF(A1{operator}A2,1,0)"
    parsed = parse_if_formula(formula)
    assert parsed.operator == operator
    result = parsed.evaluate({"A1": left, "A2": right})
    assert result == (1 if expected else 0)


def test_if_reference_vs_literal():
    parsed = parse_if_formula("=IF(A1<>0,1,0)")
    assert parsed.right.kind == "NUM"
    assert parsed.evaluate({"A1": 0}) == 0
    assert parsed.evaluate({"A1": 5}) == 1


def test_if_result_literal_text():
    parsed = parse_if_formula('=IF(A1<A2,"* aviso *","")')
    assert parsed.evaluate({"A1": 1, "A2": 2}) == "* aviso *"
    assert parsed.evaluate({"A1": 2, "A2": 1}) == ""


def test_if_result_is_cell_reference():
    parsed = parse_if_formula("=IF(A1>0,B1,0)")
    assert parsed.value_ref_coords == ("A1", "B1")
    assert parsed.evaluate({"A1": 1, "B1": 99}) == 99
    assert parsed.evaluate({"A1": 0, "B1": 99}) == 0


def test_if_empty_string_condition_uses_legacy_text_semantics():
    # Replica el patrón real: IF(celda="", ...) — celda en blanco (None) cuenta
    # como "" (como Excel); un 0 numérico NO cuenta como "".
    parsed = parse_if_formula('=IF(A1="","vacío","lleno")')
    assert parsed.evaluate({"A1": None}) == "vacío"
    assert parsed.evaluate({"A1": 0}) == "lleno"
    assert parsed.evaluate({"A1": ""}) == "vacío"


# --- IF anidado ----------------------------------------------------------


def test_nested_if_matches_real_workbook_shape():
    formula = '=IF(C9<>0, IF(AO9="","* falta dato *",""), "")'
    parsed = parse_if_formula(formula)
    assert parsed.value_ref_coords == ("C9", "AO9")
    assert parsed.evaluate({"C9": 0, "AO9": 0}) == ""          # C9=0 -> ni se mira AO9
    assert parsed.evaluate({"C9": 5, "AO9": 0}) == ""          # AO9 lleno (0) -> sin aviso
    assert parsed.evaluate({"C9": 5, "AO9": None}) == "* falta dato *"  # AO9 en blanco


def test_nested_if_numeric_branches():
    formula = "=IF(C9<>0, IF(AO9=\"\",1,0), 0)"
    parsed = parse_if_formula(formula)
    assert parsed.evaluate({"C9": 5, "AO9": None}) == 1
    assert parsed.evaluate({"C9": 5, "AO9": 0}) == 0
    assert parsed.evaluate({"C9": 0, "AO9": None}) == 0


# --- IF no soportado ------------------------------------------------------


@pytest.mark.parametrize(
    "formula",
    [
        "=IF(A1>0,1,0,0)",              # 4 argumentos
        '=IF(A1&"x"="x",1,0)',          # operando con "&"
        "=IF(A1>0,A2*2,0)",             # resultado con operación aritmética
        "=SUM(A1:A2)",                  # no es un IF
        "=IF(A1>0,1,0)+1",              # contenido tras el IF
    ],
)
def test_unsupported_if_variants(formula):
    with pytest.raises(UnsupportedFormulaError):
        parse_if_formula(formula)
