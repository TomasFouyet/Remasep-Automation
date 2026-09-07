"""Tests del modelo de provenance de fuente (Sprint 3.4). Puro, sin workbooks."""

from __future__ import annotations

from remasep.services import source_provenance as sp


def _ctx(**kw) -> sp.DeltaContext:
    base = {"sheet": "REMASEP 01"}
    base.update(kw)
    return sp.DeltaContext(**base)


# --- inputs directos: EGRESOS (CLIENT_CONFIRMED) -----------------


def test_b2_anexo_with_procedure_code_is_egresos_client_confirmed():
    prov = sp.classify_direct(_ctx(
        sheet="B2 ANEXO", procedure_code="1302008",
        row_path=("INTERVENCIONES QUIRÚRGICAS", "Rinoplastía"),
    ))
    assert prov.source == sp.SRC_EGRESOS
    assert prov.status == sp.ST_CLIENT_CONFIRMED
    assert prov.change_origin == sp.ORIGIN_DIRECT_INPUT
    assert prov.needs_human_review is False
    assert "1302008" in prov.evidence_summary()


def test_b1_age_sex_grid_is_egresos_client_confirmed():
    prov = sp.classify_direct(_ctx(
        sheet="REMASEP B1",
        section_path=("SECCIÓN A: RESUMEN DE INTERVENCIONES QUIRÚRGICAS ...",),
        column_path=("POR GRUPO DE EDAD (AÑOS)", "Menor de 10 años"),
    ))
    assert (prov.source, prov.status) == (sp.SRC_EGRESOS, sp.ST_CLIENT_CONFIRMED)
    assert prov.needs_human_review is False


def test_b2_anexo_row_without_code_is_unknown():
    prov = sp.classify_direct(_ctx(sheet="B2 ANEXO", row_path=("A.- ATENCIÓN",)))
    assert prov.source == sp.SRC_UNKNOWN_PENDING
    assert prov.needs_human_review is True


# --- CONTROL metadata ------------------------------------------


def test_control_sheet_is_control_metadata():
    prov = sp.classify_direct(_ctx(
        sheet="CONTROL", control_sheet_without_data="URGENCIAS",
        control_reason="No tenemos servicio de Urgencia",
    ))
    assert prov.source == sp.SRC_CONTROL_METADATA
    assert prov.status == sp.ST_STRUCTURALLY_INFERRED
    assert prov.needs_human_review is False
    assert "URGENCIAS" in prov.evidence_summary()


# --- REMASEP 01 Sección D: recursos vs tabla quirúrgica ---------


def test_section_d_surgical_table_column():
    prov = sp.classify_direct(_ctx(
        section_path=("SECCIÓN D: CAPACIDAD INSTALADA Y UTILIZACIÓN DE QUIRÓFANOS",),
        column_path=("QUIRÓFANOS CIRUGIA MAYOR",
                     "Horas Mensuales Ocupadas de Quirófanos en Trabajo"),
    ))
    assert prov.source == sp.SRC_SURGICAL_TABLE
    assert prov.status == sp.ST_STRUCTURALLY_INFERRED
    assert prov.needs_human_review is True  # regla por columna no confirmada


def test_section_d_resource_calculation_column():
    prov = sp.classify_direct(_ctx(
        section_path=("SECCIÓN D: CAPACIDAD INSTALADA Y UTILIZACIÓN DE QUIRÓFANOS",),
        column_path=("NÚMERO DE QUIRÓFANOS EN DOTACIÓN",),
    ))
    assert prov.source == sp.SRC_RESOURCE_CALCULATION
    assert prov.status == sp.ST_STRUCTURALLY_INFERRED
    # nunca CLIENT_CONFIRMED: el cliente no confirmó la regla por columna
    assert prov.status != sp.ST_CLIENT_CONFIRMED


def test_section_d_ambiguous_column_is_unknown():
    prov = sp.classify_direct(_ctx(
        section_path=("SECCIÓN D: CAPACIDAD INSTALADA Y UTILIZACIÓN DE QUIRÓFANOS",),
        column_path=("Total",),
    ))
    assert prov.source == sp.SRC_UNKNOWN_PENDING


# --- REMASEP 01 Sección E: fuente NO confirmada ------------------


def test_section_e_suspensions_is_unknown_pending():
    prov = sp.classify_direct(_ctx(
        section_path=("SECCIÓN E: CAUSAS DE SUSPENSIÓN DE CIRUGÍAS ELECTIVAS",),
        row_path=("Paciente",), column_path=("Nº DE PERSONAS", "Menores de 15 años"),
    ))
    assert prov.source == sp.SRC_UNKNOWN_PENDING
    assert prov.status == sp.ST_UNKNOWN
    assert "suspensión" in prov.evidence_summary() or "suspensi" in prov.evidence_summary().lower()


def test_section_e_not_assumed_surgical_table():
    prov = sp.classify_direct(_ctx(
        section_path=("SECCIÓN E: CAUSAS DE SUSPENSIÓN DE CIRUGÍAS ELECTIVAS",),
    ))
    assert prov.source != sp.SRC_SURGICAL_TABLE  # no asumir tabla quirúrgica


# --- combine_derived --------------------------------------------


def _p(source: str, **kw) -> sp.SourceProvenance:
    return sp.SourceProvenance(
        source=source,
        status=kw.get("status", sp.ST_CLIENT_CONFIRMED),
        change_origin=kw.get("change_origin", sp.ORIGIN_DIRECT_INPUT),
        evidence=kw.get("evidence", ()),
        needs_human_review=kw.get("needs_human_review", False),
    )


def test_combine_single_source_is_derived_from_dependencies():
    prov = sp.combine_derived([_p(sp.SRC_EGRESOS), _p(sp.SRC_EGRESOS)])
    assert prov.source == sp.SRC_EGRESOS
    assert prov.status == sp.ST_DERIVED_FROM_DEPENDENCIES
    assert prov.change_origin == sp.ORIGIN_FORMULA_PROPAGATION


def test_combine_multiple_sources_is_mixed_derived():
    prov = sp.combine_derived([
        _p(sp.SRC_RESOURCE_CALCULATION),
        _p(sp.SRC_SURGICAL_TABLE),
        _p(sp.SRC_UNKNOWN_PENDING, status=sp.ST_UNKNOWN),
    ])
    assert prov.source == sp.SRC_MIXED_DERIVED
    assert prov.status == sp.ST_DERIVED_FROM_DEPENDENCIES
    assert prov.needs_human_review is True
    ev = prov.evidence_summary()
    assert "RESOURCE_CALCULATION" in ev and "SURGICAL_TABLE" in ev


def test_combine_no_upstream_is_unknown_pending():
    prov = sp.combine_derived([])
    assert prov.source == sp.SRC_UNKNOWN_PENDING
    assert prov.status == sp.ST_UNKNOWN
    assert prov.needs_human_review is True


def test_combine_unknown_upstream_flags_review():
    prov = sp.combine_derived([_p(sp.SRC_UNKNOWN_PENDING, status=sp.ST_UNKNOWN)])
    assert prov.needs_human_review is True


def test_combine_nested_mixed_stays_mixed():
    inner = sp.combine_derived([_p(sp.SRC_EGRESOS), _p(sp.SRC_SURGICAL_TABLE)])
    outer = sp.combine_derived([inner, _p(sp.SRC_EGRESOS)])
    assert outer.source == sp.SRC_MIXED_DERIVED


def test_direct_and_propagation_origins_are_distinct_constants():
    assert sp.ORIGIN_DIRECT_INPUT != sp.ORIGIN_FORMULA_PROPAGATION
    assert sp.classify_direct(_ctx(sheet="CONTROL")).change_origin == sp.ORIGIN_DIRECT_INPUT
    assert sp.combine_derived([_p(sp.SRC_EGRESOS)]).change_origin == sp.ORIGIN_FORMULA_PROPAGATION


# --- cambio de EXPRESIÓN de fórmula ----------------------------------


def test_unknown_expression_change_is_unknown_and_needs_review():
    for before, after, kind in [
        ("", "=A1*2", "ADDED"),
        ("=A1*2", "", "REMOVED"),
        ("=A1*2", "=A1*3", "MODIFIED"),
    ]:
        prov = sp.unknown_expression_change(before, after)
        assert prov.source == sp.SRC_UNKNOWN_PENDING
        assert prov.status == sp.ST_UNKNOWN
        assert prov.change_origin == sp.ORIGIN_FORMULA_EXPRESSION_CHANGE
        assert prov.needs_human_review is True
        assert kind in prov.evidence_summary()
        # nunca se afirma input directo ni propagación
        assert prov.change_origin not in (sp.ORIGIN_DIRECT_INPUT, sp.ORIGIN_FORMULA_PROPAGATION)
