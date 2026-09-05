"""Tests de las transformaciones derivadas legacy AC:AL (Sprint 2.2).

Reproducen el comportamiento EXACTO del workbook; casos sintéticos, sin data/local.
"""

from __future__ import annotations

from datetime import date

import pytest

from remasep.core.text import normalize_legacy_text, normalize_text
from remasep.services.legacy_rules import load_legacy_rules
from remasep.services.legacy_transform import (
    LEGACY_NA,
    excel_str,
    legacy_age_years,
    legacy_derived_record,
    legacy_derived_value,
)


@pytest.fixture(scope="module")
def rules():
    return load_legacy_rules()


def _rec(**kw):
    base = {
        "TIPO_DE_CITA": "",
        "SUCURSAL": "",
        "SEXO": "",
        "ESPECIALIDAD": "",
        "PRESTACION": "",
        "DIA_CITA": "01-07-2026",
        "FECHA_NACIMIENTO": "01/01/2000",
    }
    base.update(kw)
    return base


# --- concatenaciones AC/AD/AE -----------------------------------------


def test_ac_concat(rules):
    rec = _rec(TIPO_DE_CITA="CONTROL ODONTOPEDIATRÍA", SUCURSAL="Santiago")
    assert legacy_derived_value("AC", rec, rules) == "CONTROL ODONTOPEDIATRÍASantiago"


def test_ad_concat(rules):
    rec = _rec(TIPO_DE_CITA="CONSULTA X", SEXO="Mujer")
    assert legacy_derived_value("AD", rec, rules) == "CONSULTA XMujer"


def test_ae_concat(rules):
    rec = _rec(PRESTACION="5001003 - CONTROL", ESPECIALIDAD="Odontopediatría", SEXO="Hombre")
    assert legacy_derived_value("AE", rec, rules) == "5001003 - CONTROLOdontopediatríaHombre"


def test_concat_with_empty_fields(rules):
    # Excel concatena "" cuando la celda está vacía; no salta el campo.
    rec = _rec(TIPO_DE_CITA="ABC", SUCURSAL="")
    assert legacy_derived_value("AC", rec, rules) == "ABC"
    rec = _rec(TIPO_DE_CITA="", SEXO="Mujer")
    assert legacy_derived_value("AD", rec, rules) == "Mujer"


def test_concat_preserves_casing_and_accents(rules):
    rec = _rec(TIPO_DE_CITA="Evaluación", SUCURSAL="ÑUÑOA")
    assert legacy_derived_value("AC", rec, rules) == "EvaluaciónÑUÑOA"


# --- AF edad --------------------------------------------------------------


def test_af_birthday_today():
    assert legacy_age_years(date(2000, 7, 1), date(2026, 7, 1)) == 26


def test_af_day_before_birthday():
    assert legacy_age_years(date(2000, 7, 2), date(2026, 7, 1)) == 25


def test_af_day_after_birthday():
    assert legacy_age_years(date(2000, 6, 30), date(2026, 7, 1)) == 26


def test_af_birth_after_service_is_zero():
    assert legacy_age_years(date(2030, 1, 1), date(2026, 7, 1)) == 0


def test_af_missing_returns_none():
    assert legacy_age_years(None, date(2026, 7, 1)) is None
    assert legacy_age_years(date(2000, 1, 1), "") is None


def test_af_parses_ddmm_strings_like_excel():
    # dd/mm/aaaa y dd-mm-aaaa
    assert legacy_age_years("28/06/2024", "01-07-2026") == 2
    rec = {"FECHA_NACIMIENTO": "07/03/1980", "DIA_CITA": "01-07-2026"}
    assert legacy_derived_value("AF", rec, load_legacy_rules()) == 46


# --- AG/AH/AI EXACT_MAP ------------------------------------------------


def test_ag_exact_map_match(rules):
    rec = _rec(
        TIPO_DE_CITA="CONSULTA MEDICA DE ESPECIALIDAD EN GENETICA CLINICA - COD.0101325",
        SEXO="Mujer",
    )
    assert legacy_derived_value("AG", rec, rules) == "consulta médicoMujer"


def test_ah_exact_map_match(rules):
    rec = _rec(TIPO_DE_CITA="CONTROL PERIODONCIA", SEXO="Hombre")
    assert legacy_derived_value("AH", rec, rules) == "control odon espHombre"


def test_ai_exact_map_match(rules):
    rec = _rec(TIPO_DE_CITA="EVALUACIÓN DE ORTODONCIA", SEXO="Mujer")
    assert legacy_derived_value("AI", rec, rules) == "evaluacion odon espMujer"


def test_exact_map_no_match_is_na(rules):
    rec = _rec(TIPO_DE_CITA="CONSULTA CUALQUIERA", SEXO="Mujer")
    assert legacy_derived_value("AG", rec, rules) == LEGACY_NA
    assert legacy_derived_value("AH", rec, rules) == LEGACY_NA


def test_exact_map_is_case_insensitive(rules):
    rec = _rec(TIPO_DE_CITA="control periodoncia", SEXO="Hombre")  # minúsculas
    assert legacy_derived_value("AH", rec, rules) == "control odon espHombre"


def test_exact_map_is_accent_sensitive(rules):
    # "CONTROL EVOLUCION DENTARIA" sin tilde NO coincide con "CONTROL EVOLUCIÓN DENTARIA".
    rec = _rec(TIPO_DE_CITA="CONTROL EVOLUCION DENTARIA", SEXO="Hombre")
    assert legacy_derived_value("AH", rec, rules) == LEGACY_NA


# --- AJ/AK/AL PATTERN_FLAG (COUNTIF) ---------------------------------


def test_aj_contains_match(rules):
    rec = _rec(PRESTACION="algo CONTROL APARATO REMOVIBLE algo", SEXO="Mujer")
    assert legacy_derived_value("AJ", rec, rules) == "1Mujer"


def test_ak_contains_match(rules):
    rec = _rec(PRESTACION="x CONTROL TAPE LABIAL PREQ. y", SEXO="Hombre")
    assert legacy_derived_value("AK", rec, rules) == "1Hombre"


def test_al_contains_match(rules):
    rec = _rec(PRESTACION="INSTALACIÓN PLACA OBTURADORA POSTQ.", SEXO="Mujer")
    assert legacy_derived_value("AL", rec, rules) == "1Mujer"


def test_pattern_flag_no_match_is_zero(rules):
    rec = _rec(PRESTACION="OTRA PRESTACION", SEXO="Hombre")
    assert legacy_derived_value("AJ", rec, rules) == "0Hombre"


def test_pattern_flag_empty_prestacion(rules):
    rec = _rec(PRESTACION="", SEXO="Mujer")
    assert legacy_derived_value("AJ", rec, rules) == "0Mujer"


def test_pattern_flag_counts_multiple_patterns(rules):
    # COUNTIF suma: si la prestación contiene dos patrones de la misma categoría -> "2".
    rec = _rec(
        PRESTACION="CONTROL APARATO REMOVIBLE y también CONTROL SPLINT",
        SEXO="Hombre",
    )
    assert legacy_derived_value("AJ", rec, rules) == "2Hombre"


def test_pattern_flag_case_insensitive(rules):
    rec = _rec(PRESTACION="previo control aparato removible posterior", SEXO="Mujer")
    assert legacy_derived_value("AJ", rec, rules) == "1Mujer"


# --- registro completo + excel_str -----------------------------------


def test_legacy_derived_record_row_two(rules):
    rec = _rec(
        TIPO_DE_CITA="CONTROL ODONTOPEDIATRÍA",
        SUCURSAL="Santiago",
        SEXO="Mujer",
        ESPECIALIDAD="Odontopediatría",
        PRESTACION="5001003 - CONTROL ODONTOLOGICO",
        FECHA_NACIMIENTO="28/06/2024",
        DIA_CITA="01-07-2026",
    )
    out = legacy_derived_record(rec, rules)
    assert out == {
        "AC": "CONTROL ODONTOPEDIATRÍASantiago",
        "AD": "CONTROL ODONTOPEDIATRÍAMujer",
        "AE": "5001003 - CONTROL ODONTOLOGICOOdontopediatríaMujer",
        "AF": 2,
        "AG": LEGACY_NA,
        "AH": "control odon espMujer",
        "AI": LEGACY_NA,
        "AJ": "0Mujer",
        "AK": "0Mujer",
        "AL": "0Mujer",
    }


def test_excel_str():
    assert excel_str(None) == ""
    assert excel_str(2) == "2"
    assert excel_str(2.0) == "2"
    assert excel_str("Mujer") == "Mujer"


# --- semántica única de matching legacy (Sprint 2.2 · revisión) -----------


def test_normalize_legacy_text_semantics():
    # case-insensitive
    assert normalize_legacy_text("CONTROL EVOLUCIÓN DENTARIA") == "control evolución dentaria"
    assert normalize_legacy_text("control evolución dentaria") == "control evolución dentaria"
    # preserva tildes (a diferencia de normalize_text global)
    assert "ó" in normalize_legacy_text("EVOLUCIÓN")
    assert normalize_legacy_text("CONTROL EVOLUCIÓN DENTARIA") != normalize_legacy_text(
        "CONTROL EVOLUCION DENTARIA"
    )
    # preserva el whitespace EXACTAMENTE: no recorta ni colapsa
    assert normalize_legacy_text(" CONTROL EVOLUCIÓN DENTARIA") == " control evolución dentaria"
    assert normalize_legacy_text("CONTROL  EVOLUCIÓN DENTARIA") == "control  evolución dentaria"
    assert normalize_legacy_text("CONTROL EVOLUCIÓN DENTARIA ") != normalize_legacy_text(
        "CONTROL EVOLUCIÓN DENTARIA"
    )
    assert normalize_legacy_text(None) == ""
    # normalize_text global NO se tocó (sigue eliminando tildes y colapsando):
    assert normalize_text("  EVOLUCIÓN   X  ") == "EVOLUCION X"


@pytest.mark.parametrize(
    ("data", "should_match"),
    [
        ("CONTROL EVOLUCIÓN DENTARIA", True),
        ("control evolución dentaria", True),  # solo casing -> MATCH
        ("CONTROL EVOLUCION DENTARIA", False),  # falta la tilde -> NO MATCH
        ("CONTROL  EVOLUCIÓN DENTARIA", False),  # doble espacio -> NO MATCH
        (" CONTROL EVOLUCIÓN DENTARIA", False),  # espacio inicial -> NO MATCH
        ("CONTROL EVOLUCIÓN DENTARIA ", False),  # espacio final -> NO MATCH
    ],
)
def test_equals_semantics_exact_transform(rules, data, should_match):
    rec = _rec(TIPO_DE_CITA=data, SEXO="Hombre")
    expected = "control odon espHombre" if should_match else LEGACY_NA
    assert legacy_derived_value("AH", rec, rules) == expected


@pytest.mark.parametrize(
    ("data", "should_match"),
    [
        ("previo CONTROL MÁSCARA Y/O DISYUNTOR posterior", True),
        ("previo control máscara y/o disyuntor posterior", True),  # casing -> MATCH
        ("previo CONTROL MASCARA Y/O DISYUNTOR posterior", False),  # sin tilde -> NO MATCH
        # doble espacio dentro del patrón: COUNTIF "*CONTROL MÁSCARA Y/O DISYUNTOR*"
        # no matchea "CONTROL  MÁSCARA ..." -> NO MATCH
        ("previo CONTROL  MÁSCARA Y/O DISYUNTOR posterior", False),
    ],
)
def test_contains_semantics_exact_transform(rules, data, should_match):
    rec = _rec(PRESTACION=data, SEXO="Mujer")
    expected = "1Mujer" if should_match else "0Mujer"
    assert legacy_derived_value("AJ", rec, rules) == expected


def test_ruleset_and_transform_share_the_same_semantics(rules):
    """LegacyRuleSet.classify y legacy_transform deben coincidir en casing/tildes/espacios."""
    accented = {"TIPO_DE_CITA": "control evolución dentaria", "PRESTACION": ""}
    no_accent = {"TIPO_DE_CITA": "CONTROL EVOLUCION DENTARIA", "PRESTACION": ""}
    padded = {"TIPO_DE_CITA": " CONTROL EVOLUCIÓN DENTARIA", "PRESTACION": ""}

    assert rules.classify(accented) == ["AH"]
    assert legacy_derived_value("AH", _rec(**accented), rules) != LEGACY_NA

    assert rules.classify(no_accent) == []
    assert legacy_derived_value("AH", _rec(**no_accent), rules) == LEGACY_NA

    assert rules.classify(padded) == []  # whitespace inicial -> NO MATCH en ambos
    assert legacy_derived_value("AH", _rec(**padded), rules) == LEGACY_NA

    # contains: misma coherencia (tildes y espacios)
    aj_hit = {"TIPO_DE_CITA": "", "PRESTACION": "x control máscara y/o disyuntor x"}
    aj_miss = {"TIPO_DE_CITA": "", "PRESTACION": "x control mascara y/o disyuntor x"}
    aj_ws = {"TIPO_DE_CITA": "", "PRESTACION": "x control  máscara y/o disyuntor x"}
    assert "AJ" in rules.classify(aj_hit)
    assert legacy_derived_value("AJ", _rec(**aj_hit), rules).startswith("1")
    assert "AJ" not in rules.classify(aj_miss)
    assert legacy_derived_value("AJ", _rec(**aj_miss), rules).startswith("0")
    assert "AJ" not in rules.classify(aj_ws)
    assert legacy_derived_value("AJ", _rec(**aj_ws), rules).startswith("0")
