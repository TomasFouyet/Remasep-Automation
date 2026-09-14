"""Regresión F01: los inválidos relevantes no desaparecen entre ingesta y UI."""

from __future__ import annotations

import json
from datetime import date

import pytest
from PySide6.QtWidgets import QApplication

from remasep.services.common import Period
from remasep.services.medinet_summary import build_monthly_medinet_summary
from remasep.services.production_pipeline import (
    InvalidPeriodRecordsError,
    build_production_pending_writes,
)
from remasep.ui.main_window import MainWindow
from remasep.ui.workers import run_generation

PERIOD = Period(7, 2026)


def _row(*, dia=date(2026, 7, 10), nac=date(1990, 1, 1), sexo="Mujer"):
    return [
        dia,
        nac,
        sexo,
        "Centro",
        "KINESIOLOGIA",
        "CONSULTA",
        "prestacion",
        "Atendido",
        "Presencial",
        "",
    ]


class _WriterSpy:
    def __init__(self) -> None:
        self.called = False

    def generate(self, **_kwargs):
        self.called = True
        raise AssertionError("el writer no debe invocarse con registros inválidos relevantes")


@pytest.mark.parametrize(
    ("invalid_row", "expected_code"),
    [
        (_row(nac="fecha-invalida"), "FECHA_NACIMIENTO_INVALID"),
        (_row(sexo=""), "SEXO_EMPTY"),
    ],
)
def test_same_period_invalid_record_is_visible_and_blocks_before_writer(
    make_medinet, tmp_path, invalid_row, expected_code
):
    medinet = make_medinet([_row(), invalid_row])

    summary = build_monthly_medinet_summary(medinet, PERIOD)
    assert summary.period_scope_records == 2
    assert summary.valid_records == 1
    assert summary.invalid_records == 1
    assert summary.validation_blocked is True
    assert summary.problem_counts == ((expected_code, 1),)

    # El diagnóstico serializable/repr sólo contiene agregados, nunca valores de fila.
    diagnostic_blob = json.dumps(summary.as_dict(), ensure_ascii=False) + repr(summary)
    assert "fecha-invalida" not in diagnostic_blob

    with pytest.raises(InvalidPeriodRecordsError) as exc:
        build_production_pending_writes(medinet, PERIOD)
    assert exc.value.invalid_records == 1
    assert exc.value.problem_counts == ((expected_code, 1),)

    writer = _WriterSpy()
    outcome = run_generation(
        medinet,
        PERIOD,
        tmp_path / "template.xlsm",
        tmp_path / "output.xlsm",
        service=writer,
    )
    assert outcome.ok is False
    assert outcome.status == "ABORTED_INVALID_SOURCE_RECORDS"
    assert writer.called is False
    assert outcome.human_error is not None
    assert "1 registro" in outcome.human_error.detail


def test_invalid_record_from_another_month_does_not_block_selected_period(make_medinet):
    medinet = make_medinet(
        [_row(), _row(dia=date(2026, 8, 10), nac="fecha-invalida")]
    )

    summary = build_monthly_medinet_summary(medinet, PERIOD)
    assert summary.period_scope_records == 1
    assert summary.valid_records == 1
    assert summary.invalid_records == 0
    assert summary.validation_blocked is False

    production = build_production_pending_writes(medinet, PERIOD)
    assert len(production.pending_writes) == 1122
    assert production.completeness.ok is True


def test_invalid_service_date_is_unassignable_and_blocks_conservatively(make_medinet):
    medinet = make_medinet([_row(), _row(dia="fecha-invalida")])

    summary = build_monthly_medinet_summary(medinet, PERIOD)
    assert summary.period_scope_records == 2
    assert summary.invalid_records == 1
    assert summary.problem_counts == (("DIA_CITA_INVALID", 1),)
    with pytest.raises(InvalidPeriodRecordsError) as exc:
        build_production_pending_writes(medinet, PERIOD)
    assert exc.value.unassigned_records == 1


def test_fully_valid_input_keeps_current_production_behavior(make_medinet):
    medinet = make_medinet([_row(), _row()])

    summary = build_monthly_medinet_summary(medinet, PERIOD)
    assert summary.period_scope_records == 2
    assert summary.valid_records == 2
    assert summary.invalid_records == 0
    assert summary.included_records == 2
    assert summary.excluded_records == 0
    assert summary.considered_ratio_label == "2 / 2"

    production = build_production_pending_writes(medinet, PERIOD)
    assert production.scope.period_scope_records == 2
    assert production.scope.processing_scope_records == 2
    assert len(production.pending_writes) == 1122
    assert production.completeness.ok is True


@pytest.mark.usefixtures("qapp")
@pytest.mark.parametrize("invalid_row", [_row(nac="fecha-invalida"), _row(sexo="")])
def test_dashboard_shows_invalid_diagnostic_and_disables_generation(
    make_medinet, invalid_row
):
    medinet = make_medinet([_row(), invalid_row])
    window = MainWindow()
    try:
        window.show()
        window.state.summary = build_monthly_medinet_summary(medinet, PERIOD)
        window.navigate("dashboard")
        QApplication.processEvents()
        dashboard = window.screens["dashboard"]
        assert dashboard._validation_warning.isVisible()
        assert "1 registro inválido" in dashboard._validation_warning.text()
        assert "Ningún registro inválido fue contado" in dashboard._validation_warning.text()
        assert dashboard.generate_button.isEnabled() is False
    finally:
        window.close()
