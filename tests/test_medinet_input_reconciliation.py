"""Tests de la reconciliación input Medinet de producción vs. referencia legacy.

Archivos sintéticos (fixture ``make_medinet``). Sin data/local, sin PII en
aserciones.
"""

from __future__ import annotations

import json
from datetime import date

import pytest
import validate_medinet_production_input as vmpi

from remasep.services import medinet_input_reconciliation as rec
from remasep.services.common import Period

JULY = Period(7, 2026)


def _row(**kw):
    base = {
        "dia": date(2026, 7, 10), "nac": date(1990, 1, 1), "sexo": "Mujer",
        "sucursal": "Santiago", "especialidad": "GENETICA", "tipo": "CONSULTA GENERAL",
        "prestacion": "prestacion x", "estado": "Atendido", "modalidad": "Fonasa B",
        "prest_real": "",
    }
    base.update(kw)
    return [
        base["dia"], base["nac"], base["sexo"], base["sucursal"], base["especialidad"],
        base["tipo"], base["prestacion"], base["estado"], base["modalidad"],
        base["prest_real"],
    ]


# --- period scope --------------------------------------------------


def test_period_scope_multi_month(make_medinet):
    path = make_medinet([
        _row(dia=date(2026, 7, 1)),
        _row(dia=date(2026, 7, 20)),
        _row(dia=date(2026, 6, 15)),
        _row(dia=date(2026, 8, 3)),
    ])
    scope = rec.period_scope(rec.load_records(path), JULY)
    assert scope.physical_records == 4
    assert scope.structurally_valid_records == 4
    assert scope.in_period_records == 2
    assert scope.out_of_period_records == 2
    assert scope.processing_scope_records == 2


def test_structurally_invalid_rows_excluded_from_scope(make_medinet):
    # filas presentes pero con DIA_CITA no parseable -> no asignables a un mes
    bad = _row(dia="no-es-fecha")
    path = make_medinet([_row(dia=date(2026, 7, 5)), bad, bad])
    loaded = rec.load_records(path)
    scope = rec.period_scope(loaded, JULY)
    assert scope.physical_records == 3
    assert scope.structurally_invalid_records == 2   # DIA_CITA no parseable
    assert scope.structurally_valid_records == 1
    assert scope.processing_scope_records == 1


# --- multiset subset --------------------------------------------


def test_legacy_exact_multiset_subset(make_medinet):
    direct = make_medinet(
        [_row(dia=date(2026, 7, d)) for d in range(1, 11)]      # 10 July
        + [_row(dia=date(2026, 6, 1))],                          # 1 other month
        filename="direct.xlsx",
    )
    legacy = make_medinet([_row(dia=date(2026, 7, d)) for d in (2, 4, 6)],
                          filename="legacy.xlsx")
    d = rec.in_period_frame(rec.load_records(direct), JULY)
    lg = rec.load_records(legacy).frame
    cmp = rec.multiset_subset_check(lg, d, rec.FINGERPRINT_FIELDS)
    assert cmp.subset_records == 3
    assert cmp.superset_records == 10
    assert cmp.matched_subset_records == 3
    assert cmp.unmatched_subset_records == 0
    assert cmp.is_exact_subset is True
    assert cmp.superset_excess_records == 7


def test_legacy_non_exact_subset_when_a_record_missing(make_medinet):
    direct = make_medinet([_row(dia=date(2026, 7, 2)), _row(dia=date(2026, 7, 4))],
                          filename="direct.xlsx")
    legacy = make_medinet(
        [_row(dia=date(2026, 7, 2)), _row(dia=date(2026, 7, 9), prestacion="NO EXISTE EN DIRECTO")],
        filename="legacy.xlsx",
    )
    d = rec.in_period_frame(rec.load_records(direct), JULY)
    lg = rec.load_records(legacy).frame
    cmp = rec.multiset_subset_check(lg, d, rec.FINGERPRINT_FIELDS)
    assert cmp.unmatched_subset_records == 1
    assert cmp.is_exact_subset is False


def test_multiset_handles_duplicate_fingerprints(make_medinet):
    # 3 citas idénticas (huella) en directo, 2 idénticas en legacy -> subset exacto
    direct = make_medinet([_row(dia=date(2026, 7, 7))] * 3, filename="direct.xlsx")
    legacy = make_medinet([_row(dia=date(2026, 7, 7))] * 2, filename="legacy.xlsx")
    d = rec.in_period_frame(rec.load_records(direct), JULY)
    lg = rec.load_records(legacy).frame
    cmp = rec.multiset_subset_check(lg, d, rec.FINGERPRINT_FIELDS)
    assert cmp.matched_subset_records == 2
    assert cmp.unmatched_subset_records == 0
    assert cmp.superset_excess_records == 1
    assert cmp.is_exact_subset is True
    # y una 3ª copia en legacy que directo no tiene -> ya no es subset
    legacy2 = make_medinet([_row(dia=date(2026, 7, 7))] * 4, filename="legacy2.xlsx")
    lg2 = rec.load_records(legacy2).frame
    cmp2 = rec.multiset_subset_check(lg2, d, rec.FINGERPRINT_FIELDS)
    assert cmp2.unmatched_subset_records == 1
    assert cmp2.is_exact_subset is False


# --- estado reconstruction (hipótesis, no regla) ----------------


def test_estado_reconstruction_reproduces_legacy_count(make_medinet):
    direct = make_medinet(
        [_row(estado="Atendido") for _ in range(6)]
        + [_row(estado="Atención Pausada")]
        + [_row(estado="Cancelado") for _ in range(3)]
        + [_row(estado="No Se Presenta") for _ in range(2)],
        filename="direct.xlsx",
    )
    d = rec.in_period_frame(rec.load_records(direct), JULY)
    er = rec.estado_reconstruction(d, legacy_active_records=7)
    assert er.direct_kept_records == 7
    assert er.direct_excluded_records == 5
    assert er.reproduces_legacy_count_exactly is True
    assert er.status == "CURRENT_LEGACY_BEHAVIOR_PENDING_FUNCTIONAL_CONFIRMATION"


def test_estado_reconstruction_non_exact(make_medinet):
    direct = make_medinet(
        [_row(estado="Atendido") for _ in range(5)]
        + [_row(estado="Cancelado")],
        filename="direct.xlsx",
    )
    d = rec.in_period_frame(rec.load_records(direct), JULY)
    er = rec.estado_reconstruction(d, legacy_active_records=99)
    assert er.direct_kept_records == 5
    assert er.reproduces_legacy_count_exactly is False


def test_estado_is_not_applied_as_a_filter():
    # el módulo NO expone ningún filtro que aplique estados al alcance;
    # `estado_reconstruction` sólo cuenta.
    assert not hasattr(rec, "apply_estado_filter")
    assert not hasattr(rec, "filter_by_estado")


# --- extras aggregate --------------------------------------------


def test_extra_records_by_status_aggregate(make_medinet):
    direct = make_medinet(
        [_row(dia=date(2026, 7, 1), estado="Atendido")]
        + [_row(dia=date(2026, 7, 2), estado="Cancelado")]
        + [_row(dia=date(2026, 7, 3), estado="Cancelado")]
        + [_row(dia=date(2026, 7, 4), estado="No Se Presenta")],
        filename="direct.xlsx",
    )
    legacy = make_medinet([_row(dia=date(2026, 7, 1), estado="Atendido")],
                          filename="legacy.xlsx")
    d = rec.in_period_frame(rec.load_records(direct), JULY)
    lg = rec.load_records(legacy).frame
    excess = rec.direct_export_excess_frame(d, lg)
    agg = dict(rec.categorical_aggregate(excess, "ESTADO"))
    assert agg == {"Cancelado": 2, "No Se Presenta": 1}


# --- orquestación + privacidad ---------------------------------


@pytest.fixture
def pair(make_medinet):
    _SECRET = "RUN-9-SECRET"
    direct = make_medinet(
        [_row(dia=date(2026, 7, d), estado="Atendido") for d in range(1, 6)]
        + [_row(dia=date(2026, 7, 6), estado="Atención Pausada")]
        + [_row(dia=date(2026, 7, 7), estado="Cancelado", especialidad=_SECRET)]
        + [_row(dia=date(2026, 7, 8), estado="No Se Presenta")]
        + [_row(dia=date(2026, 6, 1), estado="Atendido")],   # otro mes
        filename="direct.xlsx",
    )
    legacy = make_medinet(
        [_row(dia=date(2026, 7, d), estado="Atendido") for d in range(1, 6)]
        + [_row(dia=date(2026, 7, 6), estado="Atención Pausada")],
        filename="legacy.xlsx",
    )
    return direct, legacy, _SECRET


def test_reconcile_end_to_end(pair):
    direct, legacy, _secret = pair
    result = rec.reconcile(direct, legacy, JULY)
    assert result.direct_scope.physical_records == 9
    assert result.direct_scope.in_period_records == 8
    assert result.direct_scope.processing_scope_records == 8
    assert result.legacy_scope.structurally_valid_records == 6
    assert result.subset_no_estado_fp.is_exact_subset is True
    assert result.estado.direct_kept_records == 6
    assert result.estado.reproduces_legacy_count_exactly is True
    assert result.estado.direct_excluded_records == 2


def test_reconcile_is_privacy_safe(pair, tmp_path):
    direct, legacy, _secret = pair
    out = vmpi.run(str(direct), str(legacy), JULY, str(tmp_path / "out"))
    # el "secreto" que metimos en ESPECIALIDAD SÍ es un campo semántico: aparece
    # sólo como agregado categórico, nunca como fila. Aseguramos que no hay
    # ninguna estructura tipo "fila individual".
    text = (tmp_path / "out" / "summary.json").read_text(encoding="utf-8")
    data = json.loads(text)
    # summary.json no lleva listas de registros
    for value in data.values():
        assert not (isinstance(value, list) and value and isinstance(value[0], dict))
    assert out["summary"]["source_hash_verified_unchanged"] is True


def test_fingerprint_has_no_identifier_fields():
    for f in rec.FINGERPRINT_FIELDS:
        assert f not in ("RUN", "RUN_PACIENTE", "NOMBRE", "NOMBRE_PACIENTE",
                         "TELEFONO", "EMAIL", "DIRECCION", "ID_PACIENTE", "ID_CITA")


def test_period_scope_invariant_raises_on_bad_counts():
    with pytest.raises(ValueError, match="inconsistente"):
        rec.PeriodScopeBreakdown(
            physical_records=10, structurally_valid_records=8,
            structurally_invalid_records=1,  # 8 + 1 != 10
            in_period_records=5, out_of_period_records=3, processing_scope_records=5,
        )
