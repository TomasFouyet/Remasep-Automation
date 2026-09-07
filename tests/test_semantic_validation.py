"""Tests de la capa de readiness / validación semántica (Sprint 3.3).

Unidades sobre `SemanticMetric` sintéticos; sin workbook, sin data/local.
"""

from __future__ import annotations

import pytest

from remasep.services import semantic_mapping as sm
from remasep.services import semantic_validation as sv

POLICY = sv.load_readiness_policy()


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
        "section_path_raw": ("SECCIÓN A: X",),
        "row_path_raw": ("5010009 - VIDRIO IONÓMERO", "Obturación"),
        "column_path_raw": ("20 A 24 AÑOS", "Hombres"),
        "sex_value": "MALE",
        "sex_status": sm.CONFIRMED,
        "age_min_years": 20,
        "age_max_years": 24,
        "age_status": sm.CONFIRMED,
        "aggregation_scope": "DETAIL",
        "aggregation_scope_status": sm.CONFIRMED,
        "procedure_code_raw": "5010009",
        "procedure_label_raw": "VIDRIO IONÓMERO",
        "procedure_status": sm.PROC_EXPLICIT,
        "mapping_status": sm.MAPPING_CONFIRMED,
        "semantic_signature": "REMASEP_OD :: A :: B :: 20 A 24 ANOS :: HOMBRES",
        "evidence": (),
        "conflicts": (),
    }
    base.update(kw)
    return sm.SemanticMetric(**base)


def _sex_conflict() -> tuple[sm.DimensionConflict, ...]:
    return (
        sm.DimensionConflict(
            dimension="SEX", label_evidence="FEMALE", formula_evidence="MALE",
            conflict_type="SEX_LABEL_FORMULA_MISMATCH",
        ),
    )


# --- AUTO_READY --------------------------------------------------------


def test_confirmed_with_required_dimensions_ok_is_auto_ready():
    r = sv.assess_readiness(_metric(), POLICY)
    assert r.readiness_status == sv.AUTO_READY
    assert r.review_reasons == ()


def test_derived_confirmed_is_auto_ready():
    r = sv.assess_readiness(_metric(kind="DERIVED_AGGREGATION"), POLICY)
    assert r.readiness_status == sv.AUTO_READY


# --- REVIEW_REQUIRED (dimensión requerida sin resolver) ------------


def test_partial_required_dimension_is_review_required():
    # REMASEP_OD DETAIL requiere `sex`; LABEL_ONLY no basta.
    r = sv.assess_readiness(
        _metric(sex_status=sm.LABEL_ONLY, mapping_status=sm.MAPPING_PARTIAL), POLICY
    )
    assert r.readiness_status == sv.REVIEW_REQUIRED
    assert "SEX_LABEL_ONLY" in r.review_reasons


def test_age_unresolved_required_is_review_required():
    r = sv.assess_readiness(
        _metric(age_status=sm.UNRESOLVED, mapping_status=sm.MAPPING_PARTIAL), POLICY
    )
    assert r.readiness_status == sv.REVIEW_REQUIRED
    assert "AGE_UNRESOLVED" in r.review_reasons


def test_aggregation_scope_unresolved_is_review_required():
    r = sv.assess_readiness(
        _metric(aggregation_scope_status=sm.UNRESOLVED, mapping_status=sm.MAPPING_PARTIAL),
        POLICY,
    )
    assert r.readiness_status == sv.REVIEW_REQUIRED
    assert "AGGREGATION_SCOPE_UNRESOLVED" in r.review_reasons


# --- PARTIAL en dimensión NO requerida sigue AUTO_READY -----------


def test_partial_non_required_dimension_can_stay_auto_ready():
    # `procedure` NO es requerida en REMASEP_OD -> UNRESOLVED es sólo INFO.
    r = sv.assess_readiness(
        _metric(
            procedure_code_raw="", procedure_label_raw="", procedure_status="UNRESOLVED",
            mapping_status=sm.MAPPING_PARTIAL,
        ),
        POLICY,
    )
    assert r.readiness_status == sv.AUTO_READY
    assert r.review_reasons == ()  # INFO no cuenta como review reason
    assert any(i.severity == sv.SEV_INFO and i.issue_type == "PROCEDURE_UNRESOLVED" for i in r.issues)


# --- BLOCKED_CONFLICT ------------------------------------------------


def test_conflict_is_blocked_and_never_auto_ready():
    r = sv.assess_readiness(
        _metric(
            sex_value="FEMALE", sex_status=sm.CONFLICT, conflicts=_sex_conflict(),
            mapping_status=sm.MAPPING_CONFLICT,
        ),
        POLICY,
    )
    assert r.readiness_status == sv.BLOCKED_CONFLICT
    assert "SEX_LABEL_FORMULA_MISMATCH" in r.review_reasons
    issue = next(i for i in r.issues if i.severity == sv.SEV_BLOCKING)
    assert issue.recommended_action == "HUMAN_REVIEW"
    assert (issue.label_evidence, issue.formula_evidence) == ("FEMALE", "MALE")


def test_conflict_not_absorbed_even_if_mapping_status_says_confirmed():
    # Regresión: ningún camino (mapping CONFIRMED ni policy) puede absorber un CONFLICT.
    r = sv.assess_readiness(
        _metric(sex_status=sm.CONFLICT, conflicts=_sex_conflict(),
                mapping_status=sm.MAPPING_CONFIRMED),
        POLICY,
    )
    assert r.readiness_status == sv.BLOCKED_CONFLICT


def test_conflict_wins_over_not_applicable_scope():
    r = sv.assess_readiness(
        _metric(aggregation_scope="TOTAL", conflicts=_sex_conflict(), sex_status=sm.CONFLICT),
        POLICY,
    )
    assert r.readiness_status == sv.BLOCKED_CONFLICT


# --- NOT_APPLICABLE (roll-up) --------------------------------------


def test_rollup_scope_is_not_applicable():
    for scope in ("TOTAL", "SUBTOTAL"):
        r = sv.assess_readiness(
            _metric(kind="DOWNSTREAM_TOTAL", aggregation_scope=scope,
                    sex_status=sm.LABEL_ONLY, age_status=sm.NOT_APPLICABLE,
                    mapping_status=sm.MAPPING_PARTIAL),
            POLICY,
        )
        assert r.readiness_status == sv.NOT_APPLICABLE
        assert r.review_reasons == ()  # sexo no es requerido en un roll-up


# --- B2 ANEXO: sex/age N/A no bloquea; procedure sí es requerida --


def test_b2_anexo_sex_age_not_applicable_does_not_block():
    r = sv.assess_readiness(
        _metric(
            form="B2_ANEXO", sheet="B2 ANEXO", cell="D5", column_path_raw=(),
            sex_value="", sex_status=sm.NOT_APPLICABLE,
            age_min_years=None, age_max_years=None, age_status=sm.NOT_APPLICABLE,
            procedure_code_raw="0601105", procedure_label_raw="Atención",
            procedure_status=sm.PROC_EXPLICIT,
        ),
        POLICY,
    )
    assert r.readiness_status == sv.AUTO_READY


def test_b2_anexo_procedure_unresolved_is_review_required():
    r = sv.assess_readiness(
        _metric(
            form="B2_ANEXO", sheet="B2 ANEXO", cell="D22", column_path_raw=(),
            sex_value="", sex_status=sm.NOT_APPLICABLE,
            age_min_years=None, age_max_years=None, age_status=sm.NOT_APPLICABLE,
            procedure_code_raw="", procedure_label_raw="", procedure_status="UNRESOLVED",
            mapping_status=sm.MAPPING_PARTIAL,
        ),
        POLICY,
    )
    assert r.readiness_status == sv.REVIEW_REQUIRED
    assert "PROCEDURE_UNRESOLVED" in r.review_reasons


# --- review reasons / issues -----------------------------------------


def test_review_reasons_are_explicit_list_not_single_string():
    r = sv.assess_readiness(
        _metric(sex_status=sm.LABEL_ONLY, age_status=sm.UNRESOLVED,
                mapping_status=sm.MAPPING_PARTIAL),
        POLICY,
    )
    assert set(r.review_reasons) == {"SEX_LABEL_ONLY", "AGE_UNRESOLVED"}
    assert len(r.review_reasons) == 2  # lista explícita, no un string ambiguo
    assert isinstance(r.review_reasons, tuple)


# --- clusters -------------------------------------------------------


def test_review_clusters_group_equivalent_issues():
    metrics = [
        _metric(metric_id=f"LEGACY::B2_ANEXO::D{n}", form="B2_ANEXO", sheet="B2 ANEXO",
                cell=f"D{n}", column_path_raw=(), sex_status=sm.NOT_APPLICABLE,
                age_min_years=None, age_max_years=None, age_status=sm.NOT_APPLICABLE,
                procedure_code_raw="", procedure_label_raw="", procedure_status="UNRESOLVED",
                aggregation_scope_status=sm.UNRESOLVED, mapping_status=sm.MAPPING_PARTIAL)
        for n in (22, 23, 27)
    ]
    result = sv.assess_all(metrics, POLICY)
    clusters = sv.build_review_clusters(result)
    proc = next(c for c in clusters if c.issue_type == "PROCEDURE_UNRESOLVED")
    assert proc.form == "B2_ANEXO" and proc.metric_count == 3
    assert len(proc.example_metric_ids) <= 5


# --- overrides: arquitectura, sin resolución ---------------------


def test_overrides_empty_by_default():
    assert sv.load_overrides() == {}


def test_override_key_is_semantic_signature_not_cell():
    m = _metric(semantic_signature="REMASEP_OD :: SIG :: X")
    assert sv.override_key(m) == "REMASEP_OD :: SIG :: X"
    assert m.cell not in sv.override_key(m)


def test_override_matched_flag_uses_signature(tmp_path):
    (tmp_path / "policy.yaml").write_text(
        (sv._DEFAULT_CONFIG_DIR / "policy.yaml").read_text(encoding="utf-8"), encoding="utf-8"
    )
    (tmp_path / "overrides.yaml").write_text(
        "version: t\nstatus: preliminary_pending_functional_validation\n"
        "overrides:\n"
        "  - metric_signature: \"REMASEP_OD :: SIG :: X\"\n"
        "    decision: REVIEW_PENDING\n"
        "    reason: test\n    evidence: test\n    approved_by: test\n    approved_date: 2026-01-01\n",
        encoding="utf-8",
    )
    overrides = sv.load_overrides(tmp_path)
    assert "REMASEP_OD :: SIG :: X" in overrides
    m = _metric(semantic_signature="REMASEP_OD :: SIG :: X")
    r = sv.assess_readiness(m, sv.load_readiness_policy(tmp_path), overrides)
    assert r.override_matched is True
    # NO resuelve: sigue evaluando por policy (aquí AUTO_READY), no cambia por el override
    assert r.readiness_status == sv.AUTO_READY


def test_missing_config_raises():
    with pytest.raises(sv.SemanticValidationError):
        sv.load_readiness_policy("/nonexistent/path/xyz")


# --- form / row_path preconditions ----------------------------------


def test_missing_row_path_is_review_required():
    r = sv.assess_readiness(_metric(row_path_raw=(), mapping_status=sm.MAPPING_PARTIAL), POLICY)
    assert r.readiness_status == sv.REVIEW_REQUIRED
    assert "ROW_PATH_MISSING" in r.review_reasons
