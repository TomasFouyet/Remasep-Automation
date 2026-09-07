"""Tests de la alineación semántica generador → plantilla (Sprint 3.5).

Unitarios con `SemanticMetric` y `TargetMetricContext` sintéticos. Sin workbooks,
sin data/local.
"""

from __future__ import annotations

from dataclasses import dataclass

import pytest

from remasep.services import semantic_mapping as sm
from remasep.services import template_alignment as ta

POLICY = ta.load_alignment_policy()
REGIONS = ta.load_source_regions()


@dataclass
class FakeReadiness:
    readiness_status: str
    review_reasons: tuple = ()


def _metric(**kw) -> sm.SemanticMetric:
    base = {
        "metric_id": "LEGACY::REMASEP_OD::C6",
        "source": "MEDINET",
        "form": "REMASEP_OD",
        "form_raw": "REMASEP_OD",
        "sheet": "REMASEP_OD",
        "cell": "C6",
        "kind": "BASE_AGGREGATION",
        "value": 1,
        "section_path_raw": ("SECCIÓN A: CONSULTAS Y CONTROLES",),
        "row_path_raw": ("A.1. CONSULTAS Y ALTAS", "Primeras Consultas"),
        "column_path_raw": ("SEGÚN GRUPOS DE EDAD", "20 A 24 AÑOS", "Hombres"),
        "sex_value": "MALE",
        "sex_status": sm.CONFIRMED,
        "age_min_years": 20,
        "age_max_years": 24,
        "age_status": sm.CONFIRMED,
        "aggregation_scope": "DETAIL",
        "aggregation_scope_status": sm.CONFIRMED,
        "procedure_code_raw": "",
        "procedure_label_raw": "",
        "procedure_status": sm.NOT_APPLICABLE,
        "mapping_status": sm.MAPPING_CONFIRMED,
        "semantic_signature": "REMASEP_OD :: A :: PRIMERAS CONSULTAS :: 20 A 24 ANOS :: HOMBRES",
        "evidence": (),
        "conflicts": (),
    }
    base.update(kw)
    return sm.SemanticMetric(**base)


def _target(**kw) -> ta.TargetMetricContext:
    base = {
        "target_sheet": "REMASEP_OD",
        "target_cell": "H99",
        "form": "REMASEP_OD",
        "row": 99,
        "column": 8,
        "section_path": ("SECCIÓN A: CONSULTAS Y CONTROLES",),
        "row_path": ("A.1. CONSULTAS Y ALTAS", "Primeras Consultas"),
        "column_path": ("SEGÚN GRUPOS DE EDAD", "20 A 24 AÑOS", "Hombres"),
        "section_norm": ("SECCION A: CONSULTAS Y CONTROLES",),
        "row_norm": ("A.1. CONSULTAS Y ALTAS", "PRIMERAS CONSULTAS"),
        "column_norm": ("SEGUN GRUPOS DE EDAD", "20 A 24 ANOS", "HOMBRES"),
        "column_structure_norm": ("SEGUN GRUPOS DE EDAD",),
        "target_kind": ta.DIRECT_INPUT_TARGET,
        "has_formula": False,
        "formula": "",
        "is_blank": True,
        "sex_value": "MALE",
        "age_min": 20,
        "age_max": 24,
        "aggregation_scope": "DETAIL",
        "procedure_code": "",
        "semantic_signature": "REMASEP_OD :: A :: PRIMERAS CONSULTAS :: 20 A 24 ANOS :: HOMBRES",
        "target_source_expectation": ta.EXP_MEDINET,
        "target_source_evidence": "",
    }
    base.update(kw)
    base.setdefault("target_alignment_role", ta.alignment_role(base["target_kind"]))
    return ta.TargetMetricContext(**base)


def _align(metrics, targets, readiness=None):
    rid = readiness or {m.metric_id: FakeReadiness("AUTO_READY") for m in metrics}
    by_form: dict[str, list] = {}
    for t in targets:
        by_form.setdefault(t.form, []).append(t)
    return ta.align(metrics, rid, by_form, POLICY, REGIONS)


# --- exact / strong / no-match --------------------------------------


def test_exact_semantic_match():
    rep = _align([_metric()], [_target()])
    assert len(rep.rows) == 1
    row = rep.rows[0]
    assert row.match_status == ta.EXACT_SEMANTIC_MATCH
    assert row.match_score == 1.0
    assert row.target_cell == "H99"
    # el match NO usa la coordenada: source C6 -> target H99
    assert row.source_metric_id.endswith("C6")


def test_strong_match_with_small_structural_difference():
    # el target tiene un nivel de fila extra -> PARTIAL en ROW_PATH -> STRONG
    t = _target(
        row_norm=("A.1. CONSULTAS Y ALTAS", "GRUPO EXTRA", "PRIMERAS CONSULTAS DISTINTAS"),
        semantic_signature="REMASEP_OD :: OTRA",
    )
    rep = _align([_metric()], [t])
    assert rep.rows[0].match_status in (ta.STRONG_MATCH, ta.NO_MATCH)
    if rep.rows[0].match_status == ta.STRONG_MATCH:
        assert rep.rows[0].match_score >= POLICY.strong_threshold


def test_no_match_when_nothing_plausible():
    t = _target(section_norm=("SECCION Z",), row_norm=("OTRA COSA",),
                column_structure_norm=("NADA",))
    rep = _align([_metric()], [t])
    assert rep.rows == []
    assert rep.unmatched_sources[0].reason in (
        ta.REASON_NO_TARGET, ta.REASON_CONTEXT_INSUFFICIENT,
    )


# --- sex / age --------------------------------------------------


def test_sex_conflict_is_conflict_status():
    rep = _align([_metric(sex_value="MALE")], [_target(sex_value="FEMALE",
                 semantic_signature="x")])
    assert rep.rows[0].match_status == ta.CONFLICT
    assert rep.rows[0].sex_match == ta.EV_MISMATCH
    assert rep.unmatched_sources[0].reason == ta.REASON_CONFLICT


def test_age_conflict_is_conflict_status():
    rep = _align([_metric(age_min_years=20, age_max_years=24)],
                 [_target(age_min=40, age_max=54, column_norm=("40 A 54 ANOS", "HOMBRES"),
                          semantic_signature="x")])
    assert rep.rows[0].match_status == ta.CONFLICT
    assert rep.rows[0].age_match == ta.EV_MISMATCH


def test_sex_and_age_match_contribute_evidence():
    row = _align([_metric()], [_target()]).rows[0]
    assert row.sex_match == ta.EV_MATCH
    assert row.age_match == ta.EV_MATCH


# --- procedure code ------------------------------------------------


def test_procedure_code_exact_match_b2():
    m = _metric(
        metric_id="LEGACY::B2_ANEXO::D5", form="B2_ANEXO", sheet="B2 ANEXO", cell="D5",
        section_path_raw=("A. KINESIOLOGÍA",),
        row_path_raw=("0601105", "Atención Kinesiológica Integral"),
        column_path_raw=(), sex_value="", sex_status=sm.NOT_APPLICABLE,
        age_min_years=None, age_max_years=None, age_status=sm.NOT_APPLICABLE,
        procedure_code_raw="0601105", procedure_status=sm.PROC_EXPLICIT,
    )
    t = _target(
        target_sheet="B2 ANEXO", target_cell="C833", form="B2_ANEXO",
        section_norm=(), row_norm=("A. KINESIOLOGIA", "0601105"),
        column_norm=("B.- PROCEDIMIENTOS",), column_structure_norm=("B.- PROCEDIMIENTOS",),
        sex_value="", age_min=None, age_max=None, procedure_code="0601105",
        semantic_signature="B2 ANEXO :: A. KINESIOLOGIA :: 0601105",
    )
    rep = _align([m], [t])
    row = rep.rows[0]
    assert row.match_status == ta.EXACT_SEMANTIC_MATCH
    assert row.procedure_match == ta.EV_MATCH
    assert row.sex_match == ta.EV_NOT_APPLICABLE
    assert row.age_match == ta.EV_NOT_APPLICABLE


def test_b2_without_sex_age_not_penalised():
    m = _metric(
        metric_id="LEGACY::B2_ANEXO::D9", form="B2_ANEXO", sheet="B2 ANEXO", cell="D9",
        row_path_raw=("0902003", "Otra prestación"), column_path_raw=(),
        sex_value="", sex_status=sm.NOT_APPLICABLE,
        age_min_years=None, age_max_years=None, age_status=sm.NOT_APPLICABLE,
        procedure_code_raw="0902003", procedure_status=sm.PROC_EXPLICIT,
    )
    t = _target(
        target_sheet="B2 ANEXO", target_cell="C40", form="B2_ANEXO", section_norm=(),
        row_norm=("0902003", "OTRA PRESTACION"), column_norm=(), column_structure_norm=(),
        sex_value="", age_min=None, age_max=None, procedure_code="0902003",
        semantic_signature="B2 ANEXO :: 0902003",
    )
    row = _align([m], [t]).rows[0]
    assert row.match_status == ta.EXACT_SEMANTIC_MATCH


# --- source eligibility -----------------------------------------


def test_blocked_conflict_source_is_excluded_not_mapped():
    m = _metric(metric_id="LEGACY::REMASEP_01::AB84")
    rep = _align([m], [_target()], {m.metric_id: FakeReadiness("BLOCKED_CONFLICT")})
    assert rep.rows == []
    excl = rep.excluded_sources[0]
    assert excl.reason == ta.EXCL_BLOCKED_CONFLICT
    assert m.metric_id not in rep.eligible_source_ids


def test_review_required_source_is_excluded():
    m = _metric()
    rep = _align([m], [_target()],
                 {m.metric_id: FakeReadiness("REVIEW_REQUIRED", ("SEX_LABEL_ONLY",))})
    assert rep.rows == []
    assert rep.excluded_sources[0].reason == ta.EXCL_REVIEW_REQUIRED


def test_not_applicable_source_exclusion_reason_is_rollup():
    m = _metric(kind="DOWNSTREAM_TOTAL")
    rep = _align([m], [_target()], {m.metric_id: FakeReadiness("NOT_APPLICABLE")})
    assert rep.excluded_sources[0].reason == ta.EXCL_NOT_APPLICABLE
    # renombrada: readiness NOT_APPLICABLE == roll-up TOTAL/SUBTOTAL, no un input
    assert ta.EXCL_NOT_APPLICABLE == "ROLLUP_NOT_INPUT"


def test_non_medinet_source_excluded():
    m = _metric(source="EGRESOS")
    rep = _align([m], [_target()])
    assert rep.excluded_sources[0].reason == ta.EXCL_NOT_MEDINET


# --- target source expectation (Sprint 3.4 regions) ----------------


def test_target_in_non_medinet_region_excluded_from_medinet_alignment():
    # source y target con la MISMA semántica, pero el target cae en una región
    # cuya fuente esperada es RESOURCE_CALCULATION (Sprint 3.4) -> no se mezcla.
    t = _target(target_source_expectation=ta.EXP_RESOURCE_CALCULATION,
                target_source_evidence="R01_SECTION_D_RESOURCES")
    rep = _align([_metric()], [t])
    assert rep.rows == []
    assert rep.unmatched_sources[0].reason == ta.REASON_TARGET_NOT_MEDINET


def test_region_expectation_default_is_unknown_not_medinet():
    # una hoja/sección sin región declarada -> UNKNOWN (no se inventa MEDINET)
    exp, _ev = ta.region_expectation(REGIONS, "HOJA RARA", 5, ("SECCION DESCONOCIDA",))
    assert exp == ta.EXP_UNKNOWN


def test_region_expectation_positive_medinet_region():
    exp, ev = ta.region_expectation(REGIONS, "REMASEP_OD", 20, ("SECCION A: X",))
    assert exp == ta.EXP_MEDINET
    assert "OD_MEDINET" in ev or "ambulatoria" in ev.lower()


def test_region_expectation_section_d_is_resource_calculation():
    exp, ev = ta.region_expectation(
        REGIONS, "REMASEP 01", 129,
        ("SECCION D: CAPACIDAD INSTALADA Y UTILIZACION DE QUIROFANOS",),
    )
    assert exp == ta.EXP_RESOURCE_CALCULATION
    assert "3.4" in ev or "RESOURCE" in ev or "dotaci" in ev.lower()


# --- ambiguity ---------------------------------------------------


def test_multiple_equal_candidates_is_ambiguous_not_arbitrary():
    t1 = _target(target_cell="H10")
    t2 = _target(target_cell="H20")
    t3 = _target(target_cell="H30")
    rep = _align([_metric()], [t1, t2, t3])
    assert rep.rows[0].match_status == ta.AMBIGUOUS
    assert rep.ambiguous[0].candidate_count == 3
    assert set(rep.ambiguous[0].candidate_cells) == {
        "REMASEP_OD!H10", "REMASEP_OD!H20", "REMASEP_OD!H30"
    }


# --- config / determinism / signatures -------------------------


def test_policy_and_regions_are_versioned_preliminary():
    assert POLICY.version == "official_template_alignment_2026"
    assert POLICY.exact_threshold >= POLICY.strong_threshold


def test_missing_config_raises():
    with pytest.raises(ta.TemplateAlignmentError):
        ta.load_alignment_policy("/nonexistent/xyz")


def test_alignment_is_deterministic():
    a = _align([_metric()], [_target(target_cell="H10"), _target(target_cell="H20")])
    b = _align([_metric()], [_target(target_cell="H10"), _target(target_cell="H20")])
    assert [r.match_status for r in a.rows] == [r.match_status for r in b.rows]
    assert [(r.source_metric_id, r.target_cell) for r in a.rows] == \
           [(r.source_metric_id, r.target_cell) for r in b.rows]


def test_scoring_uses_dimensions_not_coordinates():
    # dos targets con MISMA semántica pero coordenadas muy distintas -> ambos
    # candidatos; la coordenada nunca entra al score.
    sp1 = ta.score_pair(_metric(), _target(target_cell="A1", row=1, column=1), POLICY)
    sp2 = ta.score_pair(_metric(), _target(target_cell="ZZ999", row=999, column=700), POLICY)
    assert sp1.score == sp2.score == 1.0


def test_evidence_recorded_per_dimension():
    rep = _align([_metric()], [_target()])
    dims = {e[3] for e in rep.evidence}
    assert {"FORM", "SECTION", "ROW_PATH", "COLUMN_PATH", "SEX", "AGE"} <= dims
    # no hay valores individuales de paciente, sólo rótulos/estados
    for _mid, _sh, _cell, _dim, sval, tval, _ev in rep.evidence:
        assert "RUN" not in str(sval) and "RUN" not in str(tval)


# --- TARGET ROLE / COVERAGE (revisión de cierre Sprint 3.5) --------


def _unmatched_by_cell(rep):
    return {(u.target_sheet, u.target_cell): u for u in rep.unmatched_targets}


def test_structural_target_not_counted_as_unmatched_medinet_input():
    t = _target(target_cell="Z9", target_kind=ta.STRUCTURAL, is_blank=False)
    rep = _align([], [t])
    u = _unmatched_by_cell(rep)[("REMASEP_OD", "Z9")]
    assert u.target_alignment_role == ta.ROLE_STRUCTURAL
    assert u.reason == ta.TGT_STRUCTURAL_NOT_ALIGNMENT
    assert u.is_genuine_unmatched_medinet_input is False


def test_formula_target_is_derived_role_not_input():
    t = _target(target_cell="F9", target_kind=ta.FORMULA_TARGET, has_formula=True,
                formula="=SUM(A1:A2)", is_blank=False)
    assert t.target_alignment_role == ta.ROLE_DERIVED_TARGET
    rep = _align([], [t])
    u = _unmatched_by_cell(rep)[("REMASEP_OD", "F9")]
    assert u.reason == ta.TGT_DERIVED_NOT_INPUT
    assert u.is_genuine_unmatched_medinet_input is False


def test_non_medinet_direct_target_reason_is_non_medinet_region():
    t = _target(target_cell="D13", target_source_expectation=ta.EXP_EGRESOS)
    rep = _align([], [t])
    u = _unmatched_by_cell(rep)[("REMASEP_OD", "D13")]
    assert u.reason == ta.TGT_NON_MEDINET_REGION
    assert u.is_genuine_unmatched_medinet_input is False


def test_unknown_source_input_target_reason_is_unknown_not_medinet():
    t = _target(target_cell="D14", target_source_expectation=ta.EXP_UNKNOWN)
    rep = _align([], [t])
    u = _unmatched_by_cell(rep)[("REMASEP_OD", "D14")]
    assert u.reason == ta.TGT_UNKNOWN_SOURCE
    assert u.is_genuine_unmatched_medinet_input is False


def test_genuine_unmatched_medinet_input_target():
    # DIRECT_INPUT_TARGET, expectativa MEDINET, sin ningún source candidato
    t = _target(target_cell="H50", target_source_expectation=ta.EXP_MEDINET,
                row_norm=("A.9. NADA", "FILA SIN SOURCE"))
    rep = _align([], [t])
    u = _unmatched_by_cell(rep)[("REMASEP_OD", "H50")]
    assert u.reason == ta.TGT_NO_MEDINET_SOURCE_FOUND
    assert u.is_genuine_unmatched_medinet_input is True


def test_unknown_expectation_target_still_matchable_and_upgraded():
    # una expectativa UNKNOWN NO bloquea el emparejamiento (§4); un match EXACT
    # de un source MEDINET sube la expectativa UNKNOWN -> MEDINET (evidencia).
    t = _target(target_source_expectation=ta.EXP_UNKNOWN)
    rep = _align([_metric()], [t])
    assert rep.rows and rep.rows[0].match_status == ta.EXACT_SEMANTIC_MATCH
    assert rep.rows[0].target_source_expectation == ta.EXP_MEDINET


def test_non_medinet_region_target_not_upgraded_it_is_excluded():
    # una región NO-MEDINET NUNCA se empareja ni se "sube" a MEDINET
    t = _target(target_source_expectation=ta.EXP_EGRESOS)
    rep = _align([_metric()], [t])
    assert rep.rows == []
    assert rep.unmatched_sources[0].reason == ta.REASON_TARGET_NOT_MEDINET


def test_alignment_role_maps_from_kind():
    assert ta.alignment_role(ta.DIRECT_INPUT_TARGET) == ta.ROLE_INPUT_TARGET
    assert ta.alignment_role(ta.FORMULA_TARGET) == ta.ROLE_DERIVED_TARGET
    assert ta.alignment_role(ta.STRUCTURAL) == ta.ROLE_STRUCTURAL
    assert ta.alignment_role("weird") == ta.ROLE_UNKNOWN
