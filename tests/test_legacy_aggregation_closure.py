"""Tests del subset aritmético de `legacy_aggregation` (Sprint 2.4).

Fórmulas derivadas ``COUNTIF(S)(...) ± celda agregada`` y utilidades asociadas.
Casos sintéticos, sin workbook. No se usa data/local.
"""

from __future__ import annotations

import pytest

from remasep.services.legacy_aggregation import (
    UnsupportedFormulaError,
    arithmetic_constructs,
    parse_derived_formula,
)

DETAIL = "Atenciones - Detalles de citas"


def _no_refs(ref: str) -> object:  # pragma: no cover - no debería llamarse
    raise AssertionError(f"no se esperaba resolver {ref!r}")


def _refs(mapping):
    return lambda ref: mapping.get(ref.replace("$", ""))


def _countifs_ae_af(extra: str = "") -> str:
    return (
        f"=COUNTIFS('{DETAIL}'!$AE:$AE,\"*COMPOSITE*\"&\"*Hombre*\","
        f"'{DETAIL}'!$AF:$AF,\"<2\"){extra}"
    )


# --- clasificación de la expresión --------------------------------------


def test_pure_aggregation_is_not_derived():
    parsed = parse_derived_formula(_countifs_ae_af(), DETAIL, _no_refs)
    assert parsed.is_pure_aggregation
    assert not parsed.is_derived_aggregation
    assert parsed.value_ref_coords == ()
    assert parsed.base_call_count == 1


def test_countifs_minus_cell_is_derived_aggregation():
    parsed = parse_derived_formula(_countifs_ae_af("-F30"), DETAIL, _no_refs)
    assert parsed.is_derived_aggregation
    assert not parsed.is_pure_aggregation
    assert parsed.value_ref_coords == ("F30",)
    assert parsed.value_refs == ((-1, "F30"),)


def test_countif_plus_cell_reference_is_derived():
    formula = f"=COUNTIF('{DETAIL}'!$AC:$AC,\"x\")+X7"
    parsed = parse_derived_formula(formula, DETAIL, _no_refs)
    assert parsed.is_derived_aggregation
    assert parsed.value_refs == ((1, "X7"),)


def test_sum_of_two_countifs_minus_cell():
    formula = _countifs_ae_af() + "+" + _countifs_ae_af().removeprefix("=") + "-F30"
    parsed = parse_derived_formula(formula, DETAIL, _no_refs)
    assert parsed.base_call_count == 2
    assert parsed.value_refs == ((-1, "F30"),)


# --- criterio ref vs valor ref (Parte E) ------------------------------


def test_criterion_cell_ref_is_not_a_value_dependency():
    formula = (
        f"=COUNTIFS('{DETAIL}'!$O:$O,$A5,'{DETAIL}'!$AF:$AF,\"<2\")-F30"
    )
    parsed = parse_derived_formula(formula, DETAIL, _refs({"A5": "CONTROL PERIODONCIA"}))
    assert parsed.criterion_refs == ("A5",)
    assert parsed.value_ref_coords == ("F30",)  # solo F30 es dependencia de valor


def test_value_cell_ref_creates_dependency():
    parsed = parse_derived_formula(_countifs_ae_af("-F30"), DETAIL, _no_refs)
    assert "F30" in parsed.value_ref_coords


# --- depends_on_age local ------------------------------------------


def test_depends_on_age_local_true_when_af_referenced():
    assert parse_derived_formula(_countifs_ae_af("-F30"), DETAIL, _no_refs).depends_on_age_local


def test_depends_on_age_local_false_without_af():
    formula = f"=COUNTIF('{DETAIL}'!$AC:$AC,\"x\")-F30"
    assert not parse_derived_formula(formula, DETAIL, _no_refs).depends_on_age_local


# --- evaluación con lookup de dependencias --------------------------


def test_evaluate_subtracts_dependency_value():
    parsed = parse_derived_formula(_countifs_ae_af("-F30"), DETAIL, _no_refs)
    rows = [
        {"AE": "una COMPOSITE de Hombre", "AF": 1},
        {"AE": "otra COMPOSITE Hombre", "AF": 1},
        {"AE": "COMPOSITE Hombre", "AF": 5},  # AF fuera de "<2"
    ]
    assert parsed.evaluate(rows, {"F30": 1}) == 1  # 2 - 1
    assert parsed.evaluate(rows, {"F30": 0}) == 2


def test_evaluate_raises_when_value_ref_missing():
    parsed = parse_derived_formula(_countifs_ae_af("-F30"), DETAIL, _no_refs)
    with pytest.raises(UnsupportedFormulaError, match="sin resolver"):
        parsed.evaluate([], {})


def test_evaluate_with_literal():
    formula = f"=COUNTIF('{DETAIL}'!$AC:$AC,\"x\")-1"
    parsed = parse_derived_formula(formula, DETAIL, _no_refs)
    assert parsed.literals == ((-1, 1.0),)
    assert parsed.evaluate([{"AC": "x"}, {"AC": "x"}], {}) == 1


# --- operadores / constructs no soportados (Parte D) -----------------


@pytest.mark.parametrize(
    ("formula", "label"),
    [
        (f"=COUNTIF('{DETAIL}'!$AC:$AC,\"x\")*2", "operator:*"),
        (f"=COUNTIF('{DETAIL}'!$AC:$AC,\"x\")/2", "operator:/"),
        (f"=COUNTIF('{DETAIL}'!$AC:$AC,\"x\")&\"y\"", "operator:&"),
        (f"=SUM('{DETAIL}'!$AF:$AF)", "function:SUM"),
        ('=IF(A1>0,1,0)', "function:IF"),
        (f"=(COUNTIF('{DETAIL}'!$AC:$AC,\"x\")-F30)+2", "grouping_parens"),
    ],
)
def test_arithmetic_constructs_labels(formula, label):
    assert label in arithmetic_constructs(formula)


def test_arithmetic_constructs_ignores_plus_minus_and_countif():
    assert arithmetic_constructs(_countifs_ae_af("-F30+G31")) == ()


@pytest.mark.parametrize(
    "formula",
    [
        "=SUM(A1:A9)",
        "=IF(B2=0,1,0)",
        "=A1*B2",
        "=A1/B2",
        "=(A1+A2)-A3",
        "=ROUND(A1,0)",
    ],
)
def test_parse_derived_rejects_unobserved_constructs(formula):
    with pytest.raises(UnsupportedFormulaError):
        parse_derived_formula(formula, DETAIL, _no_refs)


def test_parse_derived_rejects_range_as_value():
    formula = f"=COUNTIF('{DETAIL}'!$AC:$AC,\"x\")-A1:A5"
    with pytest.raises(UnsupportedFormulaError):
        parse_derived_formula(formula, DETAIL, _no_refs)


def test_parse_derived_still_rejects_non_detail_range():
    with pytest.raises(UnsupportedFormulaError):
        parse_derived_formula("=COUNTIF('Otra'!$A:$A,\"x\")-F30", DETAIL, _no_refs)
