"""Sprint 3.10 — modelo de datos del resumen mensual Medinet."""

from __future__ import annotations

import json
from datetime import date
from pathlib import Path

import pytest

from remasep.services.common import Period
from remasep.services.medinet_summary import (
    MonthlyMedinetSummary,
    build_monthly_medinet_summary,
)

_MEDINET_REAL = Path("data/local/detalle_citas - 2026-09-07T123630.940.xlsx")

_PII_SENTINELS = ("12345678-9", "PACIENTE DE PRUEBA", "correo@ejemplo.cl", "+56911112222")


def _row(*, estado="Atendido", sexo="Mujer", esp="Ortodoncia", nac=date(2016, 1, 1),
         dia=date(2026, 7, 10)):
    return [dia, nac, sexo, "Centro", esp, "CONSULTA", "prestacion", estado, "Presencial", ""]


def test_summary_totals_and_invariants(make_medinet):
    rows = (
        [_row(estado="Atendido") for _ in range(30)]
        + [_row(estado="En Sala de Espera") for _ in range(5)]
        + [_row(estado="Cancelado") for _ in range(10)]
        + [_row(estado="No Se Presenta") for _ in range(8)]
    )
    s = build_monthly_medinet_summary(make_medinet(rows), Period(7, 2026))

    assert s.period_scope_records == 53
    assert s.included_records == 35            # 30 Atendido + 5 En Sala de Espera
    assert s.excluded_records == 18            # 10 Cancelado + 8 No Se Presenta
    # invariante clave
    assert s.included_records + s.excluded_records == s.period_scope_records
    assert s.included_percentage == pytest.approx(66.0, abs=0.1)
    assert s.considered_ratio_label == "35 / 53"

    # las distribuciones agregadas suman lo que corresponde
    assert sum(e.count for e in s.estado_distribution) == s.period_scope_records
    assert sum(c.count for c in s.sex_distribution) == s.included_records
    assert sum(c.count for c in s.age_distribution) == s.included_records
    assert sum(c.count for c in s.service_distribution) == s.included_records


def test_estado_distribution_marks_included_and_excluded(make_medinet):
    rows = [_row(estado="Atendido"), _row(estado="Cancelado"), _row(estado="En Atención")]
    s = build_monthly_medinet_summary(make_medinet(rows), Period(7, 2026))
    by = {e.label: e.included for e in s.estado_distribution}
    assert by["Atendido"] is True
    assert by["En Atención"] is True
    assert by["Cancelado"] is False


def test_summary_period_matches_request(make_medinet):
    s = build_monthly_medinet_summary(make_medinet([_row()]), Period(7, 2026))
    assert s.period_label == "Julio 2026"
    assert (s.period_month, s.period_year) == (7, 2026)


def test_out_of_period_rows_excluded_from_scope(make_medinet):
    rows = [_row(dia=date(2026, 7, 5)), _row(dia=date(2026, 8, 5)), _row(dia=date(2025, 7, 5))]
    s = build_monthly_medinet_summary(make_medinet(rows), Period(7, 2026))
    assert s.period_scope_records == 1
    assert s.included_records == 1


def test_summary_has_no_pii(make_medinet):
    rows = [_row(sexo="Hombre"), _row(sexo="Mujer")]
    s = build_monthly_medinet_summary(make_medinet(rows), Period(7, 2026))
    blob = json.dumps(s.as_dict(), ensure_ascii=False) + repr(s)
    for pii in _PII_SENTINELS:
        assert pii not in blob
    # el dict sólo tiene claves agregadas, ninguna "rut"/"nombre"/"fecha_nacimiento"
    flat = json.dumps(s.as_dict()).lower()
    for forbidden in ("rut", "run", "nombre", "apellido", "telefono", "email", "nacimiento"):
        assert forbidden not in flat


def test_pending_sources_present(make_medinet):
    s = build_monthly_medinet_summary(make_medinet([_row()]), Period(7, 2026))
    assert "Tabla quirúrgica" in s.pending_sources
    assert any("Egresos" in p for p in s.pending_sources)


# ------------------------------------------------------------------
# Escenario real Julio 2026 — coincide con el pipeline (gated)
# ------------------------------------------------------------------


@pytest.mark.skipif(not _MEDINET_REAL.is_file(), reason="requiere data/local de desarrollo")
def test_july_2026_summary_matches_production_pipeline():
    from remasep.services.production_pipeline import build_production_pending_writes

    period = Period(7, 2026)
    s = build_monthly_medinet_summary(_MEDINET_REAL, period)
    prod = build_production_pending_writes(_MEDINET_REAL, period)

    assert s.period_scope_records == prod.scope.period_scope_records == 2114
    assert s.included_records == prod.scope.processing_scope_records == 1364
    assert s.excluded_records == prod.scope.estado_excluded_records == 750
    assert s.included_percentage == pytest.approx(64.5, abs=0.1)
    assert isinstance(s, MonthlyMedinetSummary)
