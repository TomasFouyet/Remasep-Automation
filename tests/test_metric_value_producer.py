"""Sprint 3.7A — pruebas del productor de MetricValue (sintéticas, sin datos reales)."""

from __future__ import annotations

import pytest

from remasep.services.common import Period
from remasep.services.legacy_rules import load_legacy_rules
from remasep.services.metric_value_producer import (
    CMP_MATCH,
    CMP_MISMATCH,
    CMP_SOURCE_UNAVAILABLE,
    CMP_ZERO_VS_BLANK,
    EXPECTED_INTEGER_COUNT,
    INPUT_SCOPE_PERIOD_AND_ESTADO_HYPOTHESIS,
    INPUT_SCOPE_PERIOD_ONLY,
    MODE_DIAGNOSTIC,
    MODE_PRODUCTION,
    ZERO_SOURCE_ZERO_TARGET_BLANK,
    ZERO_SOURCE_ZERO_TARGET_ZERO,
    CompletenessReport,
    LegacyFormulaSpec,
    MetricValue,
    MetricValueProducer,
    MetricValueProducerError,
    build_rows,
    check_write_completeness,
    compare_reference_values,
    join_metric_values_with_manifest,
    validate_value_type,
)

DETAIL = "Detalle"
_PERIOD = Period(7, 2026)


class _Instruction:
    def __init__(self, iid: str, mid: str, sheet: str, cell: str, evt: str = "INTEGER_COUNT") -> None:
        self.instruction_id = iid
        self.source_metric_id = mid
        self.target_sheet = sheet
        self.target_cell = cell
        self.expected_value_type = evt


def _no_ref(_ref: str) -> object:
    return None


def _producer(rows, *, mode=MODE_PRODUCTION, period=_PERIOD):
    # m_female: COUNTIF sexo == F  (BASE)
    # m_total : m_female + C1  (DERIVED, C1 = COUNTIF sexo == M en la misma hoja)
    index = {
        "m_female": LegacyFormulaSpec(
            "m_female", "S", "A1", "BASE_AGGREGATION",
            f"=COUNTIF('{DETAIL}'!$H:$H,\"F\")", (),
        ),
        "m_male": LegacyFormulaSpec(
            "m_male", "S", "C1", "BASE_AGGREGATION",
            f"=COUNTIF('{DETAIL}'!$H:$H,\"M\")", (),
        ),
        "m_total": LegacyFormulaSpec(
            "m_total", "S", "B1", "DERIVED_AGGREGATION",
            f"=COUNTIF('{DETAIL}'!$H:$H,\"F\")+C1", ("C1",),
        ),
    }
    return MetricValueProducer(
        formula_index=index,
        detail_sheet=DETAIL,
        resolver_by_sheet={"S": _no_ref},
        rows=rows,
        period=period,
        mode=mode,
    )


ROWS = [
    {"H": "F", "P": "Atendido"},
    {"H": "F", "P": "Atendido"},
    {"H": "M", "P": "Cancelado"},
    {"H": "M", "P": "Atendido"},
]


# ---------------------------------------------------------------------------
# Evaluación básica
# ---------------------------------------------------------------------------


def test_base_and_derived_aggregation_evaluate_over_given_rows() -> None:
    producer = _producer(ROWS)
    assert producer.produce("m_female").value == 2
    assert producer.produce("m_male").value == 2
    # derivada: 2 (F) + C1(=2 M) = 4
    assert producer.produce("m_total").value == 4


def test_period_is_preserved_on_every_metric_value() -> None:
    producer = _producer(ROWS, period=Period(3, 2027))
    metric_value = producer.produce("m_female")
    assert metric_value.period == "Marzo 2027"
    assert metric_value.producer_version.startswith("medinet_value_producer_2026")


def test_unknown_source_metric_id_raises() -> None:
    with pytest.raises(MetricValueProducerError):
        _producer(ROWS).produce("no_such_metric")


def test_rows_outside_scope_do_not_affect_values() -> None:
    # El productor sólo ve las filas que se le pasan: añadir/quitar filas de otro
    # período o estado cambia el conteo -> el alcance lo decide quien construye
    # las filas (processing_scope), no el productor.
    fewer = _producer(ROWS[:2])
    assert fewer.produce("m_female").value == 2
    assert fewer.produce("m_male").value == 0


# ---------------------------------------------------------------------------
# Modos
# ---------------------------------------------------------------------------


def test_production_mode_does_not_apply_estado_hypothesis() -> None:
    producer = _producer(ROWS, mode=MODE_PRODUCTION)
    assert producer.mode == MODE_PRODUCTION
    assert producer.is_diagnostic is False
    assert producer.estado_filter_applied is False
    assert producer.input_scope == INPUT_SCOPE_PERIOD_ONLY
    # ve las 4 filas, incluida la Cancelado
    assert producer.produce("m_male").value == 2


def test_default_mode_is_production_diagnostic_requires_opt_in() -> None:
    default = MetricValueProducer(
        formula_index={}, detail_sheet=DETAIL, resolver_by_sheet={}, rows=[], period=Period(7, 2026)
    )
    assert default.mode == MODE_PRODUCTION
    assert default.estado_filter_applied is False


def test_diagnostic_mode_marks_output_and_scope() -> None:
    producer = _producer(ROWS, mode=MODE_DIAGNOSTIC)
    assert producer.is_diagnostic is True
    assert producer.estado_filter_applied is True
    assert producer.input_scope == INPUT_SCOPE_PERIOD_AND_ESTADO_HYPOTHESIS


def test_diagnostic_mode_reproduces_the_legacy_subset() -> None:
    kept = [row for row in ROWS if row["P"] in {"Atendido"}]
    producer = _producer(kept, mode=MODE_DIAGNOSTIC)
    # sólo 1 hombre "Atendido"
    assert producer.produce("m_male").value == 1
    assert producer.produce("m_female").value == 2


def test_invalid_mode_raises() -> None:
    with pytest.raises(MetricValueProducerError):
        MetricValueProducer(
            formula_index={}, detail_sheet=DETAIL, resolver_by_sheet={},
            rows=[], period=Period(7, 2026), mode="SOMETHING_ELSE",
        )


# ---------------------------------------------------------------------------
# Tipo de valor (§6)
# ---------------------------------------------------------------------------


def test_integer_count_accepted() -> None:
    ok, coerced, reason = validate_value_type(5, EXPECTED_INTEGER_COUNT)
    assert (ok, coerced, reason) == (True, 5, "OK")
    ok, coerced, _ = validate_value_type(4.0, EXPECTED_INTEGER_COUNT)
    assert ok and coerced == 4 and isinstance(coerced, int)


def test_integer_count_rejects_float_string_bool_and_negative() -> None:
    assert validate_value_type(4.5, EXPECTED_INTEGER_COUNT)[0] is False
    assert validate_value_type("7", EXPECTED_INTEGER_COUNT)[0] is False
    assert validate_value_type(True, EXPECTED_INTEGER_COUNT)[0] is False
    assert validate_value_type(-1, EXPECTED_INTEGER_COUNT)[0] is False


def test_produce_run_separates_conflicts_from_values() -> None:
    index = {
        "good": LegacyFormulaSpec("good", "S", "A1", "BASE_AGGREGATION",
                                  f"=COUNTIF('{DETAIL}'!$H:$H,\"F\")", ()),
        "neg": LegacyFormulaSpec("neg", "S", "A2", "DERIVED_AGGREGATION",
                                 f"=COUNTIF('{DETAIL}'!$H:$H,\"F\")-10", ()),
    }
    producer = MetricValueProducer(
        formula_index=index, detail_sheet=DETAIL, resolver_by_sheet={"S": _no_ref},
        rows=ROWS, period=Period(7, 2026),
    )
    run = producer.produce_run(["good", "neg"])
    assert [mv.source_metric_id for mv in run.metric_values] == ["good"]
    assert [c.source_metric_id for c in run.value_type_conflicts] == ["neg"]


# ---------------------------------------------------------------------------
# Equivalencia de valores de referencia (§8 / §9)
# ---------------------------------------------------------------------------


def test_zero_source_zero_target_is_match_not_zero_vs_blank() -> None:
    status, zero_sem = compare_reference_values(0, 0, target_present=True)
    assert status == CMP_MATCH
    assert zero_sem == ZERO_SOURCE_ZERO_TARGET_ZERO


def test_zero_source_blank_target_is_explicit_zero_vs_blank() -> None:
    status, zero_sem = compare_reference_values(0, None, target_present=False)
    assert status == CMP_ZERO_VS_BLANK
    assert zero_sem == ZERO_SOURCE_ZERO_TARGET_BLANK


def test_nonzero_mismatch() -> None:
    status, _ = compare_reference_values(7, 5, target_present=True)
    assert status == CMP_MISMATCH


def test_nonzero_match() -> None:
    status, _ = compare_reference_values(7, 7.0, target_present=True)
    assert status == CMP_MATCH


def test_source_unavailable() -> None:
    status, _ = compare_reference_values(None, 3, target_present=True)
    assert status == CMP_SOURCE_UNAVAILABLE


# ---------------------------------------------------------------------------
# Unión con el manifiesto + completitud (§16 / §17)
# ---------------------------------------------------------------------------


def _values(*pairs) -> list[MetricValue]:
    return [
        MetricValue(source_metric_id=mid, value=val, period="Julio 2026", producer_version="t")
        for mid, val in pairs
    ]


def test_manifest_join_is_exact() -> None:
    instructions = [
        _Instruction("wi:1", "m_a", "S", "A1"),
        _Instruction("wi:2", "m_b", "S", "B1"),
    ]
    pending = join_metric_values_with_manifest(_values(("m_a", 3), ("m_b", 0)), instructions)
    assert [(p.instruction_id, p.value) for p in pending] == [("wi:1", 3), ("wi:2", 0)]
    assert pending[0].target_sheet == "S" and pending[0].target_cell == "A1"


def test_manifest_join_rejects_duplicate_metric_value_no_last_one_wins() -> None:
    instructions = [_Instruction("wi:1", "m_a", "S", "A1")]
    with pytest.raises(MetricValueProducerError):
        join_metric_values_with_manifest(_values(("m_a", 3), ("m_a", 9)), instructions)


def test_completeness_detects_missing_duplicate_and_orphan() -> None:
    instructions = [
        _Instruction("wi:1", "m_a", "S", "A1"),
        _Instruction("wi:2", "m_b", "S", "B1"),
    ]
    report: CompletenessReport = check_write_completeness(
        _values(("m_a", 1), ("m_a", 2), ("m_c", 9)), instructions
    )
    assert report.missing == ("m_b",)
    assert report.duplicate == ("m_a",)
    assert report.orphan == ("m_c",)
    assert report.ok is False


def test_completeness_ok_when_one_value_per_instruction() -> None:
    instructions = [_Instruction("wi:1", "m_a", "S", "A1")]
    report = check_write_completeness(_values(("m_a", 4)), instructions)
    assert report.ok is True


# ---------------------------------------------------------------------------
# Privacidad (§20)
# ---------------------------------------------------------------------------

_PII_SENTINELS = ("12345678-9", "PACIENTE PRUEBA", "correo@ejemplo.cl", "+56911112222")


def test_build_rows_keeps_only_column_letters_and_derived_no_pii() -> None:
    ruleset = load_legacy_rules()
    detail_columns_map = {
        "DIA_CITA": "D", "FECHA_NACIMIENTO": "G", "SEXO": "H", "SUCURSAL": "K",
        "ESPECIALIDAD": "M", "TIPO_DE_CITA": "O", "PRESTACION": "AA",
        "ESTADO": "P", "MODALIDAD": "J", "PRESTACION_REALIZADA": "AB",
    }
    record = {
        "DIA_CITA": "2026-07-10", "FECHA_NACIMIENTO": "1990-01-01", "SEXO": "F",
        "SUCURSAL": "Centro", "ESPECIALIDAD": "Odontología", "TIPO_DE_CITA": "Nueva",
        "PRESTACION": "Consulta", "ESTADO": "Atendido", "MODALIDAD": "Presencial",
        "PRESTACION_REALIZADA": "Sí",
        # columnas PII que jamás deben propagarse:
        "RUN": "12345678-9", "NOMBRE": "PACIENTE PRUEBA",
        "EMAIL": "correo@ejemplo.cl", "TELEFONO": "+56911112222",
    }
    rows = build_rows([record], detail_columns_map, ruleset)
    assert len(rows) == 1
    keys = set(rows[0])
    assert keys == set(detail_columns_map.values()) | {
        "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL"
    }
    blob = repr(rows[0])
    for sentinel in _PII_SENTINELS:
        assert sentinel not in blob
