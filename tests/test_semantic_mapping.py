"""Tests del modelo semántico explícito (Sprint 3.2).

Unidades puras sobre `MetricInput`; sin workbook, sin data/local.
"""

from __future__ import annotations

import pytest

from remasep.services import semantic_mapping as sm

CONFIG = sm.load_mapping_config()


def _inp(**kw) -> sm.MetricInput:
    base = {
        "metric_id": "LEGACY::REMASEP_OD::C6",
        "sheet": "REMASEP_OD",
        "cell": "C6",
        "kind": "BASE_AGGREGATION",
        "value": 1,
        "section_path": ("SECCIÓN A: X",),
        "row_path": ("COD-1", "Prestación uno"),
        "column_path": (),
        "text_criteria": (),
        "formula": "",
        "formula_age": None,
        "label_age": None,
        "label_age_text": "",
        "semantic_signature": "sig",
        "context_status": "COMPLETE",
        "own_row_is_header": False,
    }
    base.update(kw)
    return sm.MetricInput(**base)


# --- FORM --------------------------------------------------------------


@pytest.mark.parametrize(
    ("sheet", "form"),
    [("REMASEP 01", "REMASEP_01"), ("B2 ANEXO", "B2_ANEXO"), ("REMASEP_OD", "REMASEP_OD")],
)
def test_form_mapping(sheet, form):
    assert sm.infer_form(sheet, CONFIG)[0] == form


def test_form_keeps_raw_sheet_name():
    assert sm.infer_form("REMASEP 01", CONFIG) == ("REMASEP_01", "REMASEP 01")


# --- SEX -------------------------------------------------------------


def test_sex_label_and_formula_agree_is_confirmed():
    inp = _inp(column_path=("POR SEXO", "Hombres"), text_criteria=("*consulta*hombre*",))
    value, status, ev, conflict = sm.infer_sex(inp, CONFIG)
    assert (value, status) == (sm.SEX_MALE, sm.CONFIRMED)
    assert conflict is None
    assert {e.evidence_type for e in ev} == {"COLUMN_LABEL", "FORMULA_CRITERION"}


def test_sex_case_normalization():
    inp = _inp(column_path=("hOmBrEs",))
    assert sm.infer_sex(inp, CONFIG)[:2] == (sm.SEX_MALE, sm.LABEL_ONLY)


def test_sex_label_only():
    inp = _inp(column_path=("Mujeres",))
    assert sm.infer_sex(inp, CONFIG)[:2] == (sm.SEX_FEMALE, sm.LABEL_ONLY)


def test_sex_formula_only_from_cellref_concat_literal():
    inp = _inp(column_path=("20-24 años",), formula='=COUNTIFS(D!$AD:$AD,$A8&"Hombre")')
    assert sm.infer_sex(inp, CONFIG)[:2] == (sm.SEX_MALE, sm.FORMULA_ONLY)


def test_sex_conflict_label_female_formula_male():
    inp = _inp(column_path=("50 - 54 años", "Mujeres"), formula='=COUNTIFS(D!$AD:$AD,$A8&"Hombre")')
    value, status, ev, conflict = sm.infer_sex(inp, CONFIG)
    assert status == sm.CONFLICT
    assert value == sm.SEX_FEMALE  # el valor "reclamado" por el rótulo
    assert conflict.dimension == "SEX"
    assert conflict.label_evidence == "FEMALE" and conflict.formula_evidence == "MALE"
    assert any(e.evidence_status == "CONFLICTS" for e in ev)


def test_sex_both_label_is_label_only_without_formula_confirmation():
    # "Ambos sexos" es un rótulo; la ausencia de criterio de sexo NO se toma
    # como confirmación independiente (no convertir heurística en hecho).
    assert sm.infer_sex(_inp(column_path=("TOTAL", "Ambos sexos")), CONFIG)[:2] == (
        sm.SEX_BOTH, sm.LABEL_ONLY,
    )
    assert sm.infer_sex(_inp(column_path=("Ambos sexo",)), CONFIG)[:2] == (
        sm.SEX_BOTH, sm.LABEL_ONLY,
    )


def test_sex_not_applicable():
    inp = _inp(column_path=("20 - 24 años",))
    assert sm.infer_sex(inp, CONFIG)[:2] == ("", sm.NOT_APPLICABLE)


def test_sex_unresolved_when_formula_mentions_both_and_no_label():
    inp = _inp(column_path=("20 - 24 años",), text_criteria=("*hombre*", "*mujer*"))
    assert sm.infer_sex(inp, CONFIG)[:2] == ("", sm.UNRESOLVED)


# --- AGE ------------------------------------------------------------------


def test_age_label_and_formula_agree():
    inp = _inp(label_age=(20, 24), label_age_text="20 A 24 AÑOS", formula_age=(20, 24))
    lo, hi, status, ev, conflict = sm.infer_age(inp)
    assert (lo, hi, status) == (20, 24, sm.CONFIRMED)
    assert conflict is None
    assert {e.evidence_type for e in ev} == {"COLUMN_LABEL", "FORMULA_BOUNDS"}


def test_age_less_than_range():
    inp = _inp(label_age=(0, 1), label_age_text="Menos de 1 año - 1 año", formula_age=(None, 1))
    lo, hi, status, *_ = sm.infer_age(inp)
    assert (lo, hi, status) == (0, 1, sm.CONFIRMED)


def test_age_open_upper_bound():
    inp = _inp(label_age=(75, None), label_age_text="75 Y MÁS", formula_age=(75, None))
    lo, hi, status, *_ = sm.infer_age(inp)
    assert (lo, hi, status) == (75, None, sm.CONFIRMED)


def test_age_label_only():
    inp = _inp(label_age=(20, 24), label_age_text="20 A 24 AÑOS", formula_age=None)
    assert sm.infer_age(inp)[:3] == (20, 24, sm.LABEL_ONLY)


def test_age_formula_only():
    inp = _inp(label_age=None, formula_age=(10, 14))
    assert sm.infer_age(inp)[:3] == (10, 14, sm.FORMULA_ONLY)


def test_age_conflict_disjoint():
    inp = _inp(label_age=(20, 24), label_age_text="20 A 24 AÑOS", formula_age=(25, 29))
    lo, hi, status, _ev, conflict = sm.infer_age(inp)
    assert status == sm.CONFLICT
    assert (lo, hi) == (20, 24)
    assert conflict.conflict_type == "AGE_LABEL_FORMULA_DISJOINT"


def test_age_not_applicable():
    assert sm.infer_age(_inp())[:3] == (None, None, sm.NOT_APPLICABLE)


# --- AGGREGATION SCOPE --------------------------------------------


def test_total_from_explicit_column_label():
    inp = _inp(column_path=("TOTAL", "Ambos sexo"))
    value, status, ev = sm.infer_aggregation_scope(inp, has_age_or_sex=False, has_procedure=False)
    assert (value, status) == (sm.SCOPE_TOTAL, sm.CONFIRMED)
    assert ev and ev[0].evidence_type == "COLUMN_LABEL"


def test_subtotal_from_group_header_row():
    inp = _inp(row_path=("A.4. ACTIVIDADES GENERALES",))
    assert sm.infer_aggregation_scope(inp, has_age_or_sex=True, has_procedure=False)[:2] == (
        sm.SCOPE_SUBTOTAL, sm.CONFIRMED,
    )


def test_subtotal_from_own_header_row_flag():
    inp = _inp(row_path=("PROCEDIMIENTOS DIAGNÓSTICOS",), own_row_is_header=True)
    assert sm.infer_aggregation_scope(inp, has_age_or_sex=False, has_procedure=False)[:2] == (
        sm.SCOPE_SUBTOTAL, sm.CONFIRMED,
    )


def test_detail_confirmed_only_with_a_real_dimension():
    inp = _inp(row_path=("Prestación x",))
    assert sm.infer_aggregation_scope(inp, has_age_or_sex=False, has_procedure=False)[:2] == (
        sm.SCOPE_DETAIL, sm.UNRESOLVED,
    )
    assert sm.infer_aggregation_scope(inp, has_age_or_sex=True, has_procedure=False)[:2] == (
        sm.SCOPE_DETAIL, sm.CONFIRMED,
    )


def test_downstream_total_kind_is_not_auto_total():
    inp = _inp(kind="DOWNSTREAM_TOTAL", column_path=("20 - 24 años",), row_path=("Prestación",))
    value, _status, _ev = sm.infer_aggregation_scope(inp, has_age_or_sex=True, has_procedure=False)
    assert value == sm.SCOPE_DETAIL  # el kind no fuerza TOTAL


# --- PROCEDURE ---------------------------------------------------


def test_procedure_inline_code():
    code, label, _row = sm.extract_procedure(("5010009 - VIDRIO IONÓMERO",), CONFIG)
    assert (code, label) == ("5010009", "VIDRIO IONÓMERO")


def test_procedure_bare_code_then_name():
    code, label, _row = sm.extract_procedure(("0601105", "Atención Kinesiológica Integral"), CONFIG)
    assert (code, label) == ("0601105", "Atención Kinesiológica Integral")


def test_procedure_text_without_code():
    assert sm.extract_procedure(("Pediatría",), CONFIG) == ("", "", "")


def test_procedure_incidental_number_not_parsed():
    assert sm.extract_procedure(("A.4. ACTIVIDADES GENERALES.38:51",), CONFIG) == ("", "", "")
    assert sm.extract_procedure(("20 A 24 AÑOS",), CONFIG) == ("", "", "")


# --- MAPPING STATUS ------------------------------------------------


def test_mapping_status_confirmed():
    assert sm.compute_mapping_status(sm.CONFIRMED, sm.CONFIRMED, sm.CONFIRMED, False) == "CONFIRMED"
    assert sm.compute_mapping_status(
        sm.NOT_APPLICABLE, sm.NOT_APPLICABLE, sm.CONFIRMED, False
    ) == "CONFIRMED"


def test_mapping_status_partial():
    assert sm.compute_mapping_status(sm.LABEL_ONLY, sm.CONFIRMED, sm.CONFIRMED, False) == "PARTIAL"
    assert sm.compute_mapping_status(
        sm.CONFIRMED, sm.CONFIRMED, sm.UNRESOLVED, False
    ) == "PARTIAL"


def test_mapping_status_conflict_wins():
    assert sm.compute_mapping_status(sm.CONFLICT, sm.CONFIRMED, sm.CONFIRMED, True) == "CONFLICT"


# --- EVIDENCE + fin a fin -----------------------------------------


def test_map_metric_end_to_end_and_evidence_preserved():
    inp = _inp(
        column_path=("SEGÚN GRUPOS DE EDAD O DE RIESGO", "20-24 años", "Mujeres"),
        row_path=("5010009 - VIDRIO IONÓMERO", "Obturación"),
        text_criteria=("*5010009 - VIDRIO IONÓMERO*", "*Mujer*"),
        formula='=COUNTIFS(D!$AE:$AE,"*5010009 - VIDRIO IONÓMERO*"&"*Mujer*",D!$AF:$AF,">=20",D!$AF:$AF,"<=24")',
        label_age=(20, 24),
        label_age_text="20-24 años",
        formula_age=(20, 24),
    )
    metric = sm.map_metric(inp, CONFIG)
    assert metric.source == "MEDINET"
    assert metric.form == "REMASEP_OD"
    assert (metric.sex_value, metric.sex_status) == (sm.SEX_FEMALE, sm.CONFIRMED)
    assert (metric.age_min_years, metric.age_max_years, metric.age_status) == (20, 24, sm.CONFIRMED)
    assert metric.aggregation_scope == sm.SCOPE_DETAIL
    assert (metric.procedure_code_raw, metric.procedure_label_raw) == ("5010009", "VIDRIO IONÓMERO")
    assert metric.mapping_status == "CONFIRMED"
    # evidencia: raw + normalized preservados
    ev_types = {(e.dimension, e.evidence_type) for e in metric.evidence}
    assert ("SEX", "COLUMN_LABEL") in ev_types
    assert ("SEX", "FORMULA_CRITERION") in ev_types
    assert ("AGE", "COLUMN_LABEL") in ev_types
    assert ("PROCEDURE", "ROW_LABEL") in ev_types
    sex_col_ev = next(e for e in metric.evidence if e.dimension == "SEX" and e.evidence_type == "COLUMN_LABEL")
    assert sex_col_ev.raw_value == "Mujeres" and sex_col_ev.normalized_value == "MUJERES"


def test_b2_anexo_row_only_metric():
    inp = _inp(
        sheet="B2 ANEXO", cell="D5",
        section_path=("A. KINESIOLOGÍA",),
        row_path=("0601105", "Atención Kinesiológica Integral Ambulatoria"),
        column_path=(),
    )
    metric = sm.map_metric(inp, CONFIG)
    assert metric.form == "B2_ANEXO"
    assert metric.sex_status == sm.NOT_APPLICABLE
    assert metric.age_status == sm.NOT_APPLICABLE
    assert metric.procedure_code_raw == "0601105"
    assert metric.aggregation_scope == sm.SCOPE_DETAIL
    assert metric.mapping_status == "CONFIRMED"


def test_config_is_marked_preliminary():
    import yaml

    from remasep.services.semantic_mapping import _DEFAULT_CONFIG_DIR

    for name in ("sex_labels.yaml", "form_labels.yaml", "procedure_codes.yaml"):
        data = yaml.safe_load((_DEFAULT_CONFIG_DIR / name).read_text(encoding="utf-8"))
        assert data["status"] == "preliminary_pending_functional_validation"
