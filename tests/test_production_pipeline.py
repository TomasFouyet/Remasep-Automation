"""Sprint 3.8 / 3.9 — path de PRODUCCIÓN Medinet → PendingWrite.

Construye 1122 MetricValues/PendingWrites usando **sólo** el export Medinet + los
assets de runtime versionados (no abre el workbook legacy ni lee ``artifacts/``),
y aplica el **filtro ESTADO confirmado** (Sprint 3.9 fase 2).
"""

from __future__ import annotations

import builtins
from datetime import date
from pathlib import Path

import openpyxl
import pytest

from remasep.core.errors import RemasepError
from remasep.services.common import Period
from remasep.services.production_pipeline import (
    apply_estado_filter,
    build_production_metric_values,
    build_production_pending_writes,
)
from remasep.services.runtime_assets import RuntimeAssetError, load_runtime_bundle

_REAL_ROOT = Path(__file__).resolve().parent.parent / "config" / "runtime_2026"
_LEGACY = Path("data/local/GENERACION DATOS REMASEP.xlsx")
_MEDINET_REAL = Path("data/local/detalle_citas - 2026-09-07T123630.940.xlsx")

_CONFIRMED_INCLUDED = ("Atendido", "En Sala de Espera", "Atención Pausada", "En Atención")
_CONFIRMED_EXCLUDED = ("Cancelado", "No Se Presenta", "Agendado", "Confirmado", "Re-Agendado")


def _row(*, estado="Atendido", dia=date(2026, 7, 10)):
    return [
        dia, date(1990, 1, 1), "Mujer", "Santiago", "ODONTOLOGIA",
        "CONSULTA GENERAL", "EVALUACIÓN ODONTOLÓGICA", estado, "Presencial", "",
    ]


@pytest.fixture
def _forbid_legacy_and_artifacts(monkeypatch):
    """Cualquier intento de abrir el workbook legacy o algo bajo artifacts/ falla."""
    forbidden = ("GENERACION DATOS REMASEP.xlsx",)
    art = f"{Path('artifacts').resolve()}"
    real_open = builtins.open
    real_load = openpyxl.load_workbook

    def guarded_open(file, *a, **k):
        s = str(Path(file).resolve()) if not str(file).startswith("<") else str(file)
        if any(f in str(file) for f in forbidden) or s.startswith(art):
            raise AssertionError(f"producción intentó abrir {file!r}")
        return real_open(file, *a, **k)

    def guarded_load(fn, *a, **k):
        if any(f in str(fn) for f in forbidden):
            raise AssertionError(f"producción intentó openpyxl.load_workbook({fn!r})")
        return real_load(fn, *a, **k)

    monkeypatch.setattr(builtins, "open", guarded_open)
    monkeypatch.setattr(openpyxl, "load_workbook", guarded_load)


# ---------------------------------------------------------------------------
# Regla ESTADO versionada
# ---------------------------------------------------------------------------


def test_runtime_bundle_declares_confirmed_estado_rule():
    bundle = load_runtime_bundle(_REAL_ROOT)
    ef = bundle.estado_filter
    assert bundle.estado_filter_status == "CONFIRMED"
    assert ef.status == "CONFIRMED"
    assert ef.confirmed is True
    assert ef.included_states == _CONFIRMED_INCLUDED
    assert ef.excluded_states == _CONFIRMED_EXCLUDED
    assert "Gantz" in ef.confirmed_source


def test_runtime_rule_matches_legacy_kept_states_constant():
    """La regla versionada y el constante dev (`LEGACY_KEPT_STATES`) deben
    describir el mismo conjunto normalizado."""
    from remasep.services.medinet_input_reconciliation import (
        LEGACY_EXCLUDED_STATES,
        LEGACY_KEPT_STATES,
    )

    ef = load_runtime_bundle(_REAL_ROOT).estado_filter
    norm = lambda xs: frozenset(s.strip().casefold() for s in xs)
    assert norm(ef.included_states) == norm(LEGACY_KEPT_STATES)
    assert norm(ef.excluded_states) == norm(LEGACY_EXCLUDED_STATES)


@pytest.mark.parametrize("estado", _CONFIRMED_INCLUDED)
def test_confirmed_included_state_passes_the_filter(estado):
    ef = load_runtime_bundle(_REAL_ROOT).estado_filter
    assert ef.includes(estado) is True
    assert ef.includes(f"  {estado}  ") is True          # espacios externos
    assert ef.includes(estado.upper()) is True           # mayúsculas


@pytest.mark.parametrize("estado", _CONFIRMED_EXCLUDED)
def test_confirmed_excluded_state_does_not_pass_the_filter(estado):
    ef = load_runtime_bundle(_REAL_ROOT).estado_filter
    assert ef.includes(estado) is False


def test_apply_estado_filter_keeps_only_included(make_medinet):
    ef = load_runtime_bundle(_REAL_ROOT).estado_filter
    rows = (
        [_row(estado=s) for s in _CONFIRMED_INCLUDED]   # 4 dentro
        + [_row(estado=s) for s in _CONFIRMED_EXCLUDED]  # 5 fuera
    )
    medinet = make_medinet(rows)
    from remasep.services.medinet_analysis import processing_scope_frame

    frame = processing_scope_frame(medinet, Period(7, 2026))
    assert len(frame) == 9
    filtered, excluded = apply_estado_filter(frame, ef)
    assert len(filtered) == 4
    assert excluded == 5
    assert set(filtered["ESTADO"]) == set(_CONFIRMED_INCLUDED)


# ---------------------------------------------------------------------------
# Contrato de runtime + filtro (sintético)
# ---------------------------------------------------------------------------


def test_production_builds_1122_pending_writes_from_runtime_only(
    make_medinet, _forbid_legacy_and_artifacts
):
    medinet = make_medinet([_row(), _row(estado="En Sala de Espera"), _row()])
    result = build_production_pending_writes(medinet, Period(7, 2026))

    assert result.bundle_version == "runtime_2026"
    assert result.template_fingerprint_id == "stf:dc624775927d4d4d"
    assert result.zero_write_policy == "WRITE_ZERO"
    assert result.estado_filter_status == "CONFIRMED"
    assert result.scope.estado_filter_applied is True
    assert result.scope.processing_scope_records == 3
    assert len(result.run.metric_values) == 1122
    assert len(result.pending_writes) == 1122
    assert result.completeness.ok is True
    assert not result.completeness.missing
    assert not result.completeness.duplicate
    assert not result.completeness.orphan
    assert result.run.value_type_conflicts == []
    assert result.run.unsupported == []
    for pw in result.pending_writes:
        assert isinstance(pw.value, int) and not isinstance(pw.value, bool)
        assert pw.value >= 0
    cells = [(pw.target_sheet, pw.target_cell) for pw in result.pending_writes]
    assert len(cells) == len(set(cells))


def test_production_applies_confirmed_estado_filter(make_medinet, _forbid_legacy_and_artifacts):
    """Producción SÍ filtra por ESTADO (regla confirmada): `Cancelado` no entra."""
    medinet = make_medinet([
        _row(estado="Atendido"),
        _row(estado="Atención Pausada"),
        _row(estado="Cancelado"),
        _row(estado="No Se Presenta"),
        _row(estado="Agendado"),
    ])
    run, _bundle, scope = build_production_metric_values(medinet, Period(7, 2026))
    assert run.mode == "PRODUCTION_PERIOD_SCOPE"
    assert scope.period_scope_records == 5
    assert scope.estado_filter_status == "CONFIRMED"
    assert scope.estado_filter_applied is True
    assert scope.estado_excluded_records == 3
    assert scope.processing_scope_records == 2


def test_out_of_period_medinet_rows_do_not_count(make_medinet, _forbid_legacy_and_artifacts):
    medinet = make_medinet([_row(), _row(dia=date(2026, 8, 10))])
    _run, _b, scope = build_production_metric_values(medinet, Period(7, 2026))
    assert scope.period_scope_records == 1
    assert scope.processing_scope_records == 1


def test_missing_runtime_bundle_is_controlled_error(monkeypatch, make_medinet):
    import remasep.services.runtime_assets as ra

    def _boom(explicit=None):
        raise RuntimeAssetError("RUNTIME_ASSET_MISSING", "simulado")

    monkeypatch.setattr(ra, "resolve_runtime_root", _boom)
    with pytest.raises(RuntimeAssetError) as exc:
        build_production_pending_writes(make_medinet([_row()]), Period(7, 2026))
    assert exc.value.code == "RUNTIME_ASSET_MISSING"
    assert isinstance(exc.value, RemasepError)
    assert "instalación no contiene" in exc.value.user_message


# ---------------------------------------------------------------------------
# Escenario real Julio 2026 — números exactos de la auditoría 3.9 fase 1
# ---------------------------------------------------------------------------

_local_available = _MEDINET_REAL.is_file()


@pytest.mark.skipif(not _local_available, reason="requiere data/local de desarrollo")
def test_july_2026_production_scope_and_metrics_match_audit():
    result = build_production_pending_writes(_MEDINET_REAL, Period(7, 2026))
    s = result.scope
    assert s.period_scope_records == 2114
    assert s.estado_excluded_records == 750
    assert s.processing_scope_records == 1364          # 2114 -> 1364
    assert s.estado_filter_status == "CONFIRMED"
    assert s.estado_filter_applied is True

    mv = result.run.metric_values
    assert len(mv) == 1122
    assert len(result.pending_writes) == 1122
    assert sum(1 for x in mv if x.value) == 187        # no-cero
    assert sum(x.value for x in mv) == 928             # suma de conteos
    assert result.completeness.ok is True
    assert result.run.value_type_conflicts == []
    assert result.run.unsupported == []


@pytest.mark.skipif(
    not (_LEGACY.is_file() and _MEDINET_REAL.is_file()),
    reason="requiere data/local de desarrollo",
)
def test_production_is_exactly_the_phase1_legacy_state_hypothesis():
    """Equivalencia exacta: producción (filtro ESTADO) == escenario B de la
    auditoría (`LEGACY_STATE_HYPOTHESIS`), métrica por métrica."""
    import sys

    sys.path.insert(0, str(Path(__file__).resolve().parent.parent / "scripts"))
    import audit_medinet_estado as audit  # type: ignore[import-not-found]
    from build_metric_values import (  # type: ignore[import-not-found]
        _diagnostic_frame,
        build_legacy_reference,
        load_manifest,
        run_producer,
    )

    from remasep.services.medinet_analysis import processing_scope_frame
    from remasep.services.metric_value_producer import join_metric_values_with_manifest

    period = Period(7, 2026)

    # A) auditoría fase 1, escenario B
    audit_result = audit.audit(_MEDINET_REAL, period)
    audit_b = audit_result["scope_comparison"]["B_LEGACY_STATE_HYPOTHESIS"]

    # B) OLD (workbook legacy) filtrado a los estados confirmados
    reference = build_legacy_reference(_LEGACY)
    old_manifest = load_manifest(
        Path("artifacts/writable_target_mapping/write_manifest_ready.csv")
    )
    old_frame = _diagnostic_frame(
        processing_scope_frame(_MEDINET_REAL, period), _CONFIRMED_INCLUDED
    )
    old_run = run_producer(
        old_frame, reference, old_manifest, period, mode="LEGACY_EQUIVALENCE_DIAGNOSTIC"
    )
    old_pending = join_metric_values_with_manifest(old_run.metric_values, old_manifest)

    # C) NEW: producción con filtro ESTADO confirmado
    new = build_production_pending_writes(_MEDINET_REAL, period)

    assert new.scope.processing_scope_records == audit_b["scope_records"] == 1364
    assert len(new.run.metric_values) == audit_b["metric_values"] == 1122
    assert len(new.pending_writes) == audit_b["pending_writes"] == 1122
    assert sum(1 for x in new.run.metric_values if x.value) == audit_b["nonzero_metric_values"]
    assert sum(x.value for x in new.run.metric_values) == audit_b["sum_of_all_counts"] == 928

    old_map = {
        p.instruction_id: (p.source_metric_id, p.target_sheet, p.target_cell, p.value)
        for p in old_pending
    }
    new_map = {
        p.instruction_id: (p.source_metric_id, p.target_sheet, p.target_cell, p.value)
        for p in new.pending_writes
    }
    assert old_map == new_map
    old_mv = {mv.source_metric_id: mv.value for mv in old_run.metric_values}
    new_mv = {mv.source_metric_id: mv.value for mv in new.run.metric_values}
    assert old_mv == new_mv
