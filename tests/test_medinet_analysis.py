"""Tests del análisis real de Medinet (Sprint 2.1).

Todos los workbooks son sintéticos (fixture ``make_medinet``). No se usa
data/local. No se escriben datos personales en ninguna aserción.
"""

from __future__ import annotations

from datetime import date

import pytest

from remasep.adapters.medinet import read_medinet
from remasep.core.errors import SourceValidationError
from remasep.services.common import Period
from remasep.services.legacy_rules import load_legacy_rules
from remasep.services.legacy_transform import legacy_derived_value
from remasep.services.medinet_analysis import (
    ANALYSIS_MODE_REAL,
    MedinetAnalysisService,
    legacy_age_years,
)

PERIOD = Period(8, 2026)
SECRET = "SECRET-RUN-9F3K"


@pytest.fixture(scope="module")
def service():
    return MedinetAnalysisService()


def _row(**kw):
    base = {
        "dia": date(2026, 8, 10),
        "nac": date(1990, 1, 1),
        "sexo": "F",
        "sucursal": "Santiago",
        "especialidad": "GENETICA",
        "tipo": "CONSULTA GENERAL",
        "prestacion": "prestacion x",
        "estado": "ATENDIDO",
        "modalidad": "Presencial",
        "prest_real": "",
    }
    base.update(kw)
    return [
        base["dia"], base["nac"], base["sexo"], base["sucursal"], base["especialidad"],
        base["tipo"], base["prestacion"], base["estado"], base["modalidad"], base["prest_real"],
    ]


# --- 1. archivo válido -------------------------------------------------


def test_valid_file(make_medinet, service):
    path = make_medinet([_row(), _row(dia=date(2026, 8, 15)), _row(dia=date(2026, 8, 20))])
    result = service.analyze(path, PERIOD)

    assert result.analysis_mode == ANALYSIS_MODE_REAL
    assert result.source_name == "medinet.xlsx"
    assert result.total_records == 3
    assert result.valid_records == 3
    assert result.invalid_records == 0
    assert result.records_in_period == 3
    assert result.total_records == result.valid_records + result.invalid_records
    assert result.valid_records == result.legacy_matches_total + result.non_target_records
    assert set(result.detected_columns) >= {
        "DIA_CITA", "FECHA_NACIMIENTO", "SEXO", "SUCURSAL", "ESPECIALIDAD",
        "TIPO_DE_CITA", "PRESTACION",
    }


# --- 2. archivo inexistente ---------------------------------------


def test_missing_file(tmp_path, service):
    with pytest.raises(SourceValidationError, match="no existe"):
        service.analyze(tmp_path / "no-existe.xlsx", PERIOD)


# --- 3. columna requerida faltante ------------------------------


def test_missing_required_column(make_medinet, service):
    headers = [
        "DIA CITA", "FECHA NACIMIENTO", "SEXO", "SUCURSAL", "ESPECIALIDAD",
        "PRESTACIÓN",  # falta TIPO DE CITA
    ]
    path = make_medinet([[date(2026, 8, 1), date(1990, 1, 1), "F", "S", "E", "p"]], headers=headers)
    with pytest.raises(SourceValidationError, match="TIPO_DE_CITA"):
        service.analyze(path, PERIOD)


# --- 4. hoja de datos ------------------------------------------------


def test_locates_data_sheet_among_decoys(make_medinet, service):
    path = make_medinet(
        [_row()],
        sheet="Datos Medinet",
        decoy_sheets=[("Portada", [["REMASEP"], ["mes"]]), ("Notas", [["a", "b"]])],
    )
    result = service.analyze(path, PERIOD)
    assert result.sheet_name == "Datos Medinet"
    assert result.total_records == 1


def test_no_sheet_with_headers_raises(make_medinet, service):
    path = make_medinet([["x", "y", "z"]], headers=["col_a", "col_b", "col_c"])
    with pytest.raises(SourceValidationError):
        service.analyze(path, PERIOD)


# --- 5/6/7. validación de registros -----------------------------


def test_invalid_dia_cita(make_medinet, service):
    path = make_medinet([_row(), _row(dia="no-es-fecha")])
    result = service.analyze(path, PERIOD)
    assert result.invalid_records == 1
    codes = {(p.row_number, p.error_code) for p in result.problems}
    assert (3, "DIA_CITA_INVALID") in codes


def test_invalid_fecha_nacimiento_but_blank_is_ok(make_medinet, service):
    path = make_medinet([_row(nac="basura"), _row(nac="")])
    result = service.analyze(path, PERIOD)
    assert result.invalid_records == 1  # solo la que tiene fecha presente e inválida
    assert any(p.error_code == "FECHA_NACIMIENTO_INVALID" for p in result.problems)
    # el desglose de edad se calcula sobre el alcance de procesamiento (1 registro
    # válido en período, sin fecha de nacimiento -> edad no calculable).
    assert result.processing_scope_records == 1
    assert result.age_missing == 1


def test_empty_sexo(make_medinet, service):
    path = make_medinet([_row(), _row(sexo="")])
    result = service.analyze(path, PERIOD)
    assert result.invalid_records == 1
    assert any(p.error_code == "SEXO_EMPTY" and p.field == "SEXO" for p in result.problems)


def test_rows_are_never_dropped(make_medinet, service):
    path = make_medinet([_row(), _row(sexo=""), _row(dia="x"), _row()])
    result = service.analyze(path, PERIOD)
    assert result.total_records == 4  # nada se elimina en silencio


# --- filas estructuralmente vacías -----------------------------


_EMPTY = [None] * 10  # las 10 columnas semánticas vacías


def test_structural_empty_rows_excluded(make_medinet, service):
    path = make_medinet(
        [_row() for _ in range(10)] + [list(_EMPTY) for _ in range(5)],
        extra_columns=[("EDAD", "formula-arrastrada")],
    )
    result = service.analyze(path, PERIOD)

    assert result.physical_rows_examined == 15
    assert result.structural_empty_rows == 5
    assert result.total_records == 10
    assert result.valid_records == 10
    assert result.invalid_records == 0
    assert result.problems == ()  # las filas vacías no generan problemas
    assert result.physical_rows_examined == result.structural_empty_rows + result.total_records
    assert result.total_records == result.valid_records + result.invalid_records
    assert any("estructuralmente vacía" in n for n in result.notes)


def test_row_with_dia_empty_but_other_fields_is_not_structural_empty(make_medinet, service):
    # DIA_CITA vacía pero SEXO y TIPO_DE_CITA presentes -> registro inválido, NO vacío
    path = make_medinet(
        [_row(), _row(dia=None)],
        extra_columns=[("EDAD", "x")],
    )
    result = service.analyze(path, PERIOD)
    assert result.structural_empty_rows == 0
    assert result.total_records == 2
    assert result.invalid_records == 1
    assert any(p.error_code == "DIA_CITA_INVALID" for p in result.problems)


def test_row_with_single_semantic_field_is_not_structural_empty(make_medinet, service):
    only_especialidad = [None, None, None, None, "GENETICA", None, None, None, None, None]
    path = make_medinet(
        [_row(), only_especialidad],
        extra_columns=[("EDAD", "x")],
    )
    result = service.analyze(path, PERIOD)
    assert result.structural_empty_rows == 0
    assert result.total_records == 2
    assert result.invalid_records == 1  # le faltan DIA/SEXO/TIPO


def test_physical_invariant_holds_with_mixed_rows(make_medinet, service):
    rows = (
        [_row() for _ in range(3)]
        + [_row(sexo="")]  # inválida
        + [list(_EMPTY) for _ in range(4)]  # estructuralmente vacías
    )
    result = service.analyze(make_medinet(rows, extra_columns=[("EDAD", "x")]), PERIOD)
    assert result.physical_rows_examined == 8
    assert result.structural_empty_rows == 4
    assert result.total_records == 4
    assert result.valid_records == 3
    assert result.invalid_records == 1


def test_empty_prestacion_alone_is_not_a_problem(make_medinet, service):
    path = make_medinet([_row(prestacion=""), _row(prestacion=""), _row()])
    result = service.analyze(path, PERIOD)
    assert result.total_records == 3
    assert result.invalid_records == 0
    assert result.prestacion_empty_records == 2
    assert all(p.field != "PRESTACION" or p.error_code == "BIRTH_AFTER_SERVICE"
               for p in result.problems)
    # tampoco hay un ValidationResult de warning "global" por PRESTACION
    assert not any("PRESTACION" in v.name.upper() for v in result.validations)


# --- 8/9. período ------------------------------------------------


def test_period_ok(make_medinet, service):
    path = make_medinet([_row(dia=date(2026, 8, d)) for d in (1, 10, 25)])
    result = service.analyze(path, PERIOD)
    assert result.records_outside_period == 0
    period_check = next(v for v in result.validations if v.name == "Período correcto")
    assert period_check.status == "ok"


def test_records_outside_period_not_filtered(make_medinet, service):
    path = make_medinet(
        [_row(dia=date(2026, 8, 5)), _row(dia=date(2026, 9, 3)), _row(dia=date(2026, 7, 30))]
    )
    result = service.analyze(path, PERIOD)
    assert result.total_records == 3
    assert result.records_in_period == 1
    assert result.records_outside_period == 2
    assert result.processing_scope_records == 1
    # los registros de otros períodos NO son un error ni una advertencia
    period_check = next(v for v in result.validations if v.name == "Período correcto")
    assert period_check.status == "ok"
    assert "otros períodos" in period_check.message
    assert "no se incluir" in period_check.message.lower()


# --- alcance de procesamiento (hotfix período) -----------------


def _classified_only_in_scope(result) -> bool:
    return result.legacy_matches_total + result.non_target_records == result.processing_scope_records


def test_processing_scope_is_valid_intersect_in_period(make_medinet, service):
    path = make_medinet([
        _row(dia=date(2026, 8, 5)),          # válido, en período
        _row(dia=date(2026, 8, 20)),         # válido, en período
        _row(dia=date(2026, 9, 3)),          # válido, OTRO período
        _row(dia=date(2026, 7, 30)),         # válido, OTRO período
        _row(dia=date(2026, 8, 9), sexo=""),  # en período pero INVÁLIDO
    ])
    result = service.analyze(path, PERIOD)
    assert result.valid_records == 4
    assert result.records_in_period == 2       # de los válidos
    assert result.processing_scope_records == 2
    assert _classified_only_in_scope(result)


def test_classification_runs_only_in_selected_period(make_medinet, service):
    # 3 en período + 5 en otros meses; sólo los 3 se clasifican / cuentan
    rows = [_row(dia=date(2026, 8, d)) for d in (2, 12, 22)]
    rows += [_row(dia=date(2026, m, 15)) for m in (1, 2, 3, 9, 10)]
    result = service.analyze(make_medinet(rows), PERIOD)
    assert result.total_records == 8
    assert result.valid_records == 8
    assert result.processing_scope_records == 3
    assert result.legacy_matches_total + result.non_target_records == 3
    assert result.age_computed + result.age_legacy_zero + result.age_missing == 3


def test_out_of_period_records_do_not_change_metrics(make_medinet, service):
    base = [_row(dia=date(2026, 8, 3)), _row(dia=date(2026, 8, 17))]
    a = service.analyze(make_medinet(list(base)), PERIOD)
    b = service.analyze(
        make_medinet([*base, _row(dia=date(2026, 5, 1)), _row(dia=date(2026, 11, 30))]),
        PERIOD,
    )
    assert a.processing_scope_records == b.processing_scope_records == 2
    assert a.legacy_counts == b.legacy_counts
    assert a.legacy_matches_total == b.legacy_matches_total
    assert a.non_target_records == b.non_target_records
    assert (a.age_computed, a.age_legacy_zero, a.age_missing) == \
           (b.age_computed, b.age_legacy_zero, b.age_missing)
    # pero el conteo de fuera-de-período sí cambia
    assert b.records_outside_period == 2 and a.records_outside_period == 0


def test_all_valid_is_not_processing_scope_when_multi_month(make_medinet, service):
    rows = [_row(dia=date(2026, 8, 10))] + [_row(dia=date(2026, 4, 10))] * 4
    result = service.analyze(make_medinet(rows), PERIOD)
    assert result.valid_records == 5
    assert result.processing_scope_records == 1
    assert result.valid_records != result.processing_scope_records


# --- 10/11/12. cálculo de edad (semántica legacy DATEDIF "Y") ---


def test_age_exact_birthday():
    assert legacy_age_years(date(2000, 8, 10), date(2026, 8, 10)) == 26


def test_age_before_birthday():
    assert legacy_age_years(date(2000, 8, 11), date(2026, 8, 10)) == 25


def test_age_after_birthday():
    assert legacy_age_years(date(2000, 8, 9), date(2026, 8, 10)) == 26


def test_age_year_change():
    assert legacy_age_years(date(2000, 12, 31), date(2026, 1, 1)) == 25


def test_age_birth_after_service_returns_zero():
    assert legacy_age_years(date(2030, 1, 1), date(2026, 8, 10)) == 0


def test_age_missing_returns_none():
    assert legacy_age_years(None, date(2026, 8, 10)) is None


def test_birth_after_service_is_flagged_but_record_stays_valid(make_medinet, service):
    path = make_medinet([_row(), _row(nac=date(2030, 5, 5))])
    result = service.analyze(path, PERIOD)
    assert result.valid_records == 2  # la fila anómala NO es inválida
    assert result.age_legacy_zero == 1
    assert any(p.error_code == "BIRTH_AFTER_SERVICE" for p in result.problems)


# --- 13..18. reglas legacy AG:AL -------------------------------


@pytest.mark.parametrize(
    ("code", "field", "value"),
    [
        ("AG", "tipo", "CONSULTA MEDICA DE ESPECIALIDAD EN GENETICA CLINICA - COD.0101325"),
        ("AH", "tipo", "CONTROL PERIODONCIA"),
        ("AI", "tipo", "EVALUACIÓN DE ORTODONCIA"),
        ("AJ", "prestacion", "algo CONTROL APARATO REMOVIBLE algo"),
        ("AK", "prestacion", "x CONTROL TAPE LABIAL PREQ. y"),
        ("AL", "prestacion", "INSTALACIÓN PLACA OBTURADORA POSTQ. incluida"),
    ],
)
def test_legacy_rule_matches(make_medinet, service, code, field, value):
    path = make_medinet([_row(**{field: value}), _row()])  # una que matchea + una neutra
    result = service.analyze(path, PERIOD)
    assert result.legacy_counts[code] == 1
    assert result.legacy_matches_total == 1
    assert result.non_target_records == 1


# --- 19. semántica legacy: case-insensitive, CON tildes, whitespace exacto ------


def test_matching_is_case_insensitive(make_medinet, service):
    path = make_medinet(
        [
            # minúsculas (el valor de la regla no lleva tildes) -> match AG
            _row(tipo="consulta medica de especialidad en genetica clinica - cod.0101325"),
            # minúsculas, espacios simples -> match AJ (contains)
            _row(prestacion="previo control aparato removible posterior"),
        ]
    )
    result = service.analyze(path, PERIOD)
    assert result.legacy_counts["AG"] == 1
    assert result.legacy_counts["AJ"] == 1


def test_matching_is_accent_sensitive(make_medinet, service):
    # "CONTROL EVOLUCION DENTARIA" (sin tilde) NO coincide con la regla AH
    # "CONTROL EVOLUCIÓN DENTARIA"; sí coincide en minúsculas con tilde.
    path = make_medinet(
        [
            _row(tipo="CONTROL EVOLUCION DENTARIA"),  # sin tilde -> no match
            _row(tipo="control evolución dentaria"),  # minúsculas + tilde -> match
        ]
    )
    result = service.analyze(path, PERIOD)
    assert result.legacy_counts["AH"] == 1


def test_legacy_matching_is_whitespace_exact_through_pipeline(make_medinet, service):
    # read_medinet ya NO recorta los valores de texto: el pipeline completo
    # (adapter -> service -> LegacyRuleSet) usa el mismo texto que el workbook.
    path = make_medinet(
        [
            _row(tipo=" CONTROL EVOLUCIÓN DENTARIA"),  # espacio inicial -> NO match AH
            _row(tipo="CONTROL EVOLUCIÓN DENTARIA "),  # espacio final -> NO match AH
            _row(tipo="CONTROL  EVOLUCIÓN DENTARIA"),  # doble espacio interno -> NO match AH
            _row(tipo="control evolución dentaria"),  # solo casing -> match AH
        ]
    )
    result = service.analyze(path, PERIOD)
    assert result.legacy_counts["AH"] == 1


def test_countif_pattern_whitespace_is_exact_through_pipeline(make_medinet, service):
    # patrón AJ = "CONTROL APARATO REMOVIBLE" (espacios simples).
    path = make_medinet(
        [
            _row(prestacion="pre control aparato removible post"),  # substring exacto -> match
            _row(prestacion="pre control  aparato removible post"),  # doble espacio -> NO match
            _row(prestacion=" CONTROL APARATO REMOVIBLE "),  # el patrón sí es substring -> match
            _row(prestacion="controlaparatoremovible"),  # sin espacios -> NO match
        ]
    )
    result = service.analyze(path, PERIOD)
    assert result.legacy_counts["AJ"] == 2


def test_legacy_concat_preserves_whitespace_through_read_medinet(make_medinet):
    # AC/AD/AE: la concatenación legacy conserva el whitespace inicial/final
    # tal como sale de read_medinet (sin llamar al service).
    rules = load_legacy_rules()
    path = make_medinet(
        [_row(tipo=" CTRL ", sucursal="  SUC", especialidad="ESP ", sexo="Mujer ", prestacion=" P")]
    )
    frame = read_medinet(path).frame
    record = frame.iloc[0].to_dict()

    assert record["TIPO_DE_CITA"] == " CTRL "
    assert record["SUCURSAL"] == "  SUC"
    assert record["SEXO"] == "Mujer "

    assert legacy_derived_value("AC", record, rules) == " CTRL   SUC"  # TIPO + SUCURSAL
    assert legacy_derived_value("AD", record, rules) == " CTRL Mujer "  # TIPO + SEXO
    assert legacy_derived_value("AE", record, rules) == " PESP Mujer "  # PRESTACION+ESPECIALIDAD+SEXO


# --- 20. no-match válido -----------------------------------------


def test_valid_but_not_targeted(make_medinet, service):
    path = make_medinet([_row(tipo="CONSULTA DESCONOCIDA XYZ", prestacion="OTRA COSA")])
    result = service.analyze(path, PERIOD)
    assert result.valid_records == 1
    assert result.legacy_matches_total == 0
    assert result.non_target_records == 1
    assert result.invalid_records == 0
    assert result.problems == ()  # un no-match NO es un problema


# --- 21/22. diagnósticos ----------------------------------------


def test_diagnostics_estado_distribution(make_medinet, service):
    rows = [_row(estado="ATENDIDO") for _ in range(3)] + [_row(estado="ANULADO")]
    result = service.analyze(make_medinet(rows), PERIOD)
    estado = next(d for d in result.diagnostics if d.field == "ESTADO")
    assert estado.present is True
    assert estado.non_empty == 4
    assert estado.unique_values == 2
    assert dict(estado.distribution) == {"ATENDIDO": 3, "ANULADO": 1}


def test_diagnostics_modalidad_distribution(make_medinet, service):
    rows = [_row(modalidad="Presencial") for _ in range(2)] + [_row(modalidad="Telemedicina")]
    result = service.analyze(make_medinet(rows), PERIOD)
    modalidad = next(d for d in result.diagnostics if d.field == "MODALIDAD")
    assert dict(modalidad.distribution) == {"Presencial": 2, "Telemedicina": 1}


def test_diagnostics_prestacion_realizada_presence_only(make_medinet, service):
    rows = [_row(prest_real="EVAL COMPLETA " + str(i)) for i in range(5)]
    result = service.analyze(make_medinet(rows), PERIOD)
    diag = next(d for d in result.diagnostics if d.field == "PRESTACION_REALIZADA")
    assert diag.present is True
    assert diag.non_empty == 5
    assert diag.distribution == ()  # nunca se listan textos completos
    assert "texto libre" in diag.note


def test_diagnostics_absent_column(make_medinet, service):
    headers = [
        "DIA CITA", "FECHA NACIMIENTO", "SEXO", "SUCURSAL", "ESPECIALIDAD",
        "TIPO DE CITA", "PRESTACIÓN",  # sin ESTADO / MODALIDAD / PRESTACIÓN REALIZADA
    ]
    path = make_medinet(
        [[date(2026, 8, 1), date(1990, 1, 1), "F", "S", "E", "CONSULTA GENERAL", "p"]],
        headers=headers,
    )
    result = service.analyze(path, PERIOD)
    estado = next(d for d in result.diagnostics if d.field == "ESTADO")
    assert estado.present is False
    assert "ESTADO" in result.missing_optional_columns
    assert "MODALIDAD" in result.missing_optional_columns


# --- 23. privacidad de los problemas --------------------------


def test_problems_carry_no_personal_data(make_medinet, service):
    headers = [
        "DIA CITA", "FECHA NACIMIENTO", "SEXO", "SUCURSAL", "ESPECIALIDAD",
        "TIPO DE CITA", "PRESTACIÓN", "ESTADO", "MODALIDAD", "PRESTACIÓN REALIZADA",
        "RUN", "NOMBRE PACIENTE",
    ]
    good = _row() + [SECRET, "Juan Pérez"]
    bad = _row(sexo="", dia="x") + [SECRET, "Ana Díaz"]
    result = service.analyze(make_medinet([good, bad], headers=headers), PERIOD)

    blob = "\n".join(
        [
            *(f"{p.row_number}|{p.field}|{p.error_code}|{p.message}" for p in result.problems),
            *(f"{v.name}|{v.message}" for v in result.validations),
            *result.notes,
            *(f"{d.field}|{d.note}|{d.distribution}" for d in result.diagnostics),
        ]
    )
    assert SECRET not in blob
    assert "Juan" not in blob and "Ana" not in blob
    assert "1990" not in blob  # ninguna fecha de nacimiento individual
    assert "RUN" not in result.detected_columns
    # los problemas solo referencian fila + campo + tipo de problema
    assert all(p.field in {"DIA_CITA", "FECHA_NACIMIENTO", "SEXO", "TIPO_DE_CITA"}
               for p in result.problems)


# --- adapter: alias de encabezados documentados --------------


def test_header_aliases(make_medinet):
    headers = [
        "FECHA CITA", "FECHA DE NACIMIENTO", "SEXO", "SUCURSAL", "ESPECIALIDAD",
        "TIPO CITA", "PRESTACION", "ESTADO CITA", "MODALIDAD", "PRESTACION REALIZADA",
    ]
    path = make_medinet([_row()], headers=headers)
    medinet = read_medinet(path)
    assert "DIA_CITA" in medinet.detected_fields
    assert "TIPO_DE_CITA" in medinet.detected_fields
    assert "ESTADO" in medinet.detected_fields


def test_rejects_non_xlsx(tmp_path, service):
    bogus = tmp_path / "medinet.csv"
    bogus.write_text("a,b,c\n1,2,3\n", encoding="utf-8")
    with pytest.raises(SourceValidationError, match="xlsx"):
        service.analyze(bogus, PERIOD)


def test_legacy_rules_config_is_flagged_pending():
    rules = load_legacy_rules()
    assert rules.status == "pending_functional_validation"
    assert rules.codes == ("AG", "AH", "AI", "AJ", "AK", "AL")
    assert "MINSAL" not in rules.disclaimer or "No son reglas oficiales MINSAL" in rules.disclaimer
