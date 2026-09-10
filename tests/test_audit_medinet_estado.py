"""Sprint 3.9 — test diagnóstico (sólo lectura) de la auditoría de ESTADO.

Fija los números que se llevan al cliente para que no deriven en silencio. No
ejercita ningún filtro por ESTADO en producción: `scripts/audit_medinet_estado`
sólo *evalúa* la hipótesis legacy.

Gated: requiere el export Medinet de desarrollo bajo `data/local/`.
"""

from __future__ import annotations

from pathlib import Path

import pytest

_MEDINET = Path("data/local/detalle_citas - 2026-09-07T123630.940.xlsx")

pytestmark = pytest.mark.skipif(
    not _MEDINET.is_file(), reason="requiere el export Medinet de desarrollo (data/local)"
)


@pytest.fixture(scope="module")
def result():
    import audit_medinet_estado as audit

    from remasep.services.common import Period

    return audit.audit(_MEDINET, Period(7, 2026))


def test_total_scope_is_2114(result):
    assert result["total_records"] == 2114
    assert result["total_is_2114"] is True
    assert sum(r["cantidad"] for r in result["estado_distribution"]) == 2114


def test_estado_distribution_matches_audit(result):
    dist = {r["estado"]: r["cantidad"] for r in result["estado_distribution"]}
    assert dist == {
        "Atendido": 1337,
        "Cancelado": 545,
        "No Se Presenta": 182,
        "Atención Pausada": 16,
        "Agendado": 11,
        "Confirmado": 7,
        "En Sala de Espera": 7,
        "Re-Agendado": 5,
        "En Atención": 4,
    }


def test_legacy_hypothesis_yields_exactly_1364(result):
    h = result["hypothesis"]
    assert h["included_count"] == 1364
    assert h["excluded_count"] == 750
    assert h["included_equals_1364"] is True
    assert h["states_not_covered_by_hypothesis"] == []


def test_scope_comparison_impact_is_locked(result):
    sc = result["scope_comparison"]
    a, b = sc["A_PERIOD_ONLY"], sc["B_LEGACY_STATE_HYPOTHESIS"]
    assert (a["scope_records"], b["scope_records"]) == (2114, 1364)
    # ambas producen el contrato completo
    assert a["metric_values"] == b["metric_values"] == 1122
    assert a["pending_writes"] == b["pending_writes"] == 1122
    assert a["completeness_ok"] is b["completeness_ok"] is True
    assert (a["nonzero_metric_values"], b["nonzero_metric_values"]) == (210, 187)
    assert (a["sum_of_all_counts"], b["sum_of_all_counts"]) == (1444, 928)
    # impacto
    assert sc["metrics_changed_A_to_B"] == 122
    assert sc["all_changes_decrease_A_gt_B"] is True
    assert sc["sum_abs_delta"] == 516
    assert sc["sum_of_all_counts_delta_A_minus_B"] == 516
    assert sc["by_target_sheet"] == {
        "B2 ANEXO": {"metrics_changed": 3, "sum_abs_delta": 87},
        "REMASEP 01": {"metrics_changed": 37, "sum_abs_delta": 230},
        "REMASEP_OD": {"metrics_changed": 82, "sum_abs_delta": 199},
    }


def test_audit_does_not_touch_production_config(tmp_path):
    """El script diagnóstico no escribe si se le pide --no-artifacts y nunca
    toca config/ ni los runtime assets."""
    import audit_medinet_estado as audit

    before = {
        p: p.read_bytes()
        for p in Path("config/runtime_2026").glob("*")
    }
    code = audit.main(["--medinet", str(_MEDINET), "--period", "2026-07", "--no-artifacts"])
    assert code == 0
    after = {p: p.read_bytes() for p in Path("config/runtime_2026").glob("*")}
    assert before == after
