"""Tests de `scripts/validate_af_age_excel_com.py` (validación AF vs Excel real).

La extracción de pares y la comparación son puro Python y corren siempre (sin
Excel) en el suite Linux habitual. La comparación real contra Excel vía COM está
marcada ``@pytest.mark.excel`` y, como el resto de la integración COM de este
proyecto (ver `test_excel_com_integration.py`), requiere una variable de entorno
explícita — nunca se dispara sola porque el entorno *podría* tener Excel
disponible (en este proyecto, vía WSL interop a `powershell.exe`, eso también es
cierto en la máquina de desarrollo Linux habitual, y no queremos que un
`pytest` normal abra Excel sin que alguien lo pida a propósito).

Ejecutar la integración real::

    REMASEP_VALIDATE_AF_EXCEL_COM=1 pytest -m excel tests/test_validate_af_age_excel_com.py -vv
"""

from __future__ import annotations

import os
from datetime import date

import pytest
import validate_af_age_excel_com as va

from remasep.core.errors import RemasepError
from remasep.services.legacy_transform import legacy_age_years

_ENV_RUN_EXCEL_COM = "REMASEP_VALIDATE_AF_EXCEL_COM"

# ---------------------------------------------------------------------------
# extract_date_pairs — dedup, scope de período, exclusión sin FECHA_NACIMIENTO
# ---------------------------------------------------------------------------


def _row(dia_cita, fecha_nacimiento, **overrides):
    base = {
        "DIA CITA": dia_cita,
        "FECHA NACIMIENTO": fecha_nacimiento,
        "SEXO": "Mujer",
        "SUCURSAL": "Santiago",
        "ESPECIALIDAD": "ODONTOLOGIA",
        "TIPO DE CITA": "CONSULTA GENERAL",
        "PRESTACIÓN": "EVALUACIÓN ODONTOLÓGICA",
        "ESTADO": "Atendido",
        "MODALIDAD": "Presencial",
        "PRESTACIÓN REALIZADA": "",
    }
    base.update(overrides)
    return list(base.values())


def test_extract_date_pairs_dedups_identical_pairs(make_medinet):
    rows = [
        _row(date(2026, 7, 10), date(1990, 1, 1)),
        _row(date(2026, 7, 11), date(1990, 1, 1)),  # mismo nacimiento, otra cita
        _row(date(2026, 7, 10), date(1990, 1, 1)),  # par idéntico a la primera
    ]
    medinet = make_medinet(rows)
    pairs = va.extract_date_pairs(medinet, months=[7], year=2026)
    assert pairs == sorted(
        {(date(1990, 1, 1), date(2026, 7, 10)), (date(1990, 1, 1), date(2026, 7, 11))}
    )


def test_extract_date_pairs_excludes_missing_birth_date(make_medinet):
    rows = [
        _row(date(2026, 7, 10), date(1990, 1, 1)),
        _row(date(2026, 7, 12), None),  # FECHA_NACIMIENTO ausente: no comparable
    ]
    medinet = make_medinet(rows)
    pairs = va.extract_date_pairs(medinet, months=[7], year=2026)
    assert pairs == [(date(1990, 1, 1), date(2026, 7, 10))]


def test_extract_date_pairs_respects_month_year_scope(make_medinet):
    rows = [
        _row(date(2026, 7, 10), date(1990, 1, 1)),
        _row(date(2026, 8, 10), date(1985, 5, 5)),  # fuera del período pedido
    ]
    medinet = make_medinet(rows)
    pairs = va.extract_date_pairs(medinet, months=[7], year=2026)
    assert pairs == [(date(1990, 1, 1), date(2026, 7, 10))]


def test_extract_date_pairs_covers_multiple_months(make_medinet):
    rows = [
        _row(date(2026, 4, 1), date(1990, 1, 1)),
        _row(date(2026, 5, 1), date(1991, 1, 1)),
    ]
    medinet = make_medinet(rows)
    pairs = va.extract_date_pairs(medinet, months=[4, 5], year=2026)
    assert pairs == [
        (date(1990, 1, 1), date(2026, 4, 1)),
        (date(1991, 1, 1), date(2026, 5, 1)),
    ]


# ---------------------------------------------------------------------------
# python_expected_ages — delega en legacy_age_years, no reimplementa nada
# ---------------------------------------------------------------------------


def test_python_expected_ages_matches_legacy_age_years_directly():
    pairs = [
        (date(1990, 1, 1), date(2026, 7, 10)),  # cumpleaños ya pasado este año
        (date(1990, 12, 31), date(2026, 7, 10)),  # cumpleaños aún no llega
        (date(2027, 1, 1), date(2026, 7, 10)),  # nacimiento posterior -> excepción legacy 0
    ]
    expected = va.python_expected_ages(pairs)
    for pair in pairs:
        assert expected[pair] == legacy_age_years(pair[0], pair[1])
    assert expected[pairs[0]] == 36
    assert expected[pairs[1]] == 35
    assert expected[pairs[2]] == 0


# ---------------------------------------------------------------------------
# compare_ages — clasificación, sin fechas en el resultado
# ---------------------------------------------------------------------------


def test_compare_ages_classifies_match_mismatch_and_errors():
    pairs = [
        (date(1990, 1, 1), date(2026, 7, 10)),
        (date(1991, 1, 1), date(2026, 7, 10)),
        (date(1992, 1, 1), date(2026, 7, 10)),
    ]
    python_ages = {pairs[0]: 36, pairs[1]: 35, pairs[2]: 34}
    excel_ages = [36, 99, None]  # match, mismatch, error (Excel no evaluó)

    result = va.compare_ages(pairs, python_ages, excel_ages)
    assert result.total == 3
    assert result.match == 1
    assert result.mismatch == 1
    assert result.errors == 1
    assert result.equivalence_rate == pytest.approx(1 / 3)


def test_compare_ages_mismatch_details_never_carry_dates():
    pairs = [(date(1990, 1, 1), date(2026, 7, 10))]
    python_ages = {pairs[0]: 36}
    excel_ages = [37]

    result = va.compare_ages(pairs, python_ages, excel_ages)
    assert result.mismatch == 1
    detail = result.mismatches[0]
    assert detail.index == 0
    assert detail.python_value == 36
    assert detail.excel_value == 37
    # el detalle expone sólo índice + valores enteros, nunca las fechas del par
    assert not hasattr(detail, "birth")
    assert not hasattr(detail, "service")
    assert set(vars(detail)) == {"index", "python_value", "excel_value"}


def test_compare_ages_all_match_gives_100_percent_equivalence():
    pairs = [(date(1990, 1, 1), date(2026, 7, 10)), (date(1991, 1, 1), date(2026, 7, 10))]
    python_ages = {pairs[0]: 36, pairs[1]: 35}
    excel_ages = [36, 35]
    result = va.compare_ages(pairs, python_ages, excel_ages)
    assert result.mismatch == 0
    assert result.errors == 0
    assert result.equivalence_rate == 1.0


def test_compare_ages_raises_on_length_mismatch():
    pairs = [(date(1990, 1, 1), date(2026, 7, 10))]
    with pytest.raises(RemasepError):
        va.compare_ages(pairs, {pairs[0]: 36}, [])


# ---------------------------------------------------------------------------
# excel_com_available / run_excel_com_ages — sin powershell.exe -> falla limpio
# ---------------------------------------------------------------------------


def test_excel_com_available_false_when_powershell_missing(monkeypatch):
    monkeypatch.setattr(va, "find_powershell", lambda: None)
    assert va.excel_com_available() is False


def test_run_excel_com_ages_raises_remasep_error_when_powershell_missing(monkeypatch):
    monkeypatch.setattr(va, "find_powershell", lambda: None)
    with pytest.raises(RemasepError):
        va.run_excel_com_ages([(date(1990, 1, 1), date(2026, 7, 10))])


def test_run_excel_com_ages_empty_pairs_short_circuits(monkeypatch):
    # Ni siquiera debería intentar resolver powershell.exe para una lista vacía.
    monkeypatch.setattr(
        va, "find_powershell", lambda: (_ for _ in ()).throw(AssertionError("no debería llamarse"))
    )
    assert va.run_excel_com_ages([], powershell_exe="powershell.exe") == []


# ---------------------------------------------------------------------------
# Integración real Excel COM (Windows nativo, o WSL con interop a Excel) —
# requiere opt-in explícito por variable de entorno, nunca se autodetecta.
# ---------------------------------------------------------------------------


@pytest.mark.excel
def test_excel_com_matches_legacy_age_years_for_known_pairs():
    if not os.environ.get(_ENV_RUN_EXCEL_COM):
        pytest.skip(f"integración COM: define la variable de entorno {_ENV_RUN_EXCEL_COM}")
    if not va.excel_com_available():
        pytest.skip("Excel COM no disponible en este entorno")

    pairs = [
        (date(1990, 1, 1), date(2026, 7, 10)),
        (date(1990, 12, 31), date(2026, 7, 10)),
        (date(2027, 1, 1), date(2026, 7, 10)),
        (date(2000, 2, 29), date(2026, 3, 1)),
    ]
    excel_ages = va.run_excel_com_ages(pairs)
    python_ages = va.python_expected_ages(pairs)
    result = va.compare_ages(pairs, python_ages, excel_ages)

    assert result.errors == 0, result.mismatches
    assert result.mismatch == 0, result.mismatches
    assert result.match == len(pairs)
