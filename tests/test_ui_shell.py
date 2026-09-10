"""Sprint 3.10 — flujo de UI MEDINET (Qt offscreen, sin Excel real).

Inicio → Nuevo informe → Análisis → Resumen mensual → Generar → Resultado.

Los tests usan las rutas síncronas (`run_now`) para ser deterministas; el hilo
de fondo que arranca cada pantalla se detiene con ``_teardown`` antes.
"""

from __future__ import annotations

from datetime import date
from pathlib import Path

import pytest
from PySide6.QtWidgets import QApplication

from remasep.services.common import Period
from remasep.ui.errors import HumanError
from remasep.ui.workers import GenerationOutcome

pytestmark = pytest.mark.usefixtures("qapp")

_MEDINET_REAL = Path("data/local/detalle_citas - 2026-09-07T123630.940.xlsx")


def _row(*, estado="Atendido", sexo="Mujer", dia=date(2026, 7, 10)):
    return [
        dia, date(2016, 1, 1), sexo, "Centro", "Ortodoncia",
        "CONSULTA", "prestacion", estado, "Presencial", "",
    ]


@pytest.fixture
def window():
    from remasep.ui.main_window import MainWindow

    win = MainWindow()
    win.show()  # con QT_QPA_PLATFORM=offscreen es seguro y hace fiable isVisible()
    yield win
    win.close()


@pytest.fixture
def medinet_file(make_medinet):
    return make_medinet(
        [_row(estado="Atendido") for _ in range(20)]
        + [_row(estado="En Sala de Espera") for _ in range(4)]
        + [_row(estado="Cancelado") for _ in range(6)]
        + [_row(estado="No Se Presenta") for _ in range(3)]
    )


@pytest.fixture
def template_file(tmp_path):
    p = tmp_path / "REMASEP_plantilla.xlsm"
    p.write_bytes(b"PK\x03\x04 placeholder")
    return p


# --- helpers ---------------------------------------------------------------


def _drain() -> None:
    QApplication.instance().processEvents()


def _fill_new_report(window, medinet, template, *, month=7, year=2026):
    nr = window.screens["new_report"]
    nr.month_combo.setCurrentIndex(month - 1)
    nr.year_spin.setValue(year)
    nr.medinet_selector.set_selection(Path(medinet))
    nr.template_selector.set_selection(Path(template))
    nr._sync_state()
    return nr


def _run_analysis(window) -> None:
    """Lanza el análisis y lo resuelve de forma determinista."""
    window.screens["new_report"]._analyze()      # navega a 'analysis' + hilo
    window.screens["analysis"]._teardown()        # detiene el hilo de fondo
    _drain()                                      # entrega su señal encolada
    window.screens["analysis"].run_now()          # ejecución síncrona (idempotente)


def _run_generate(window) -> None:
    window.navigate("generate")                    # on_enter arranca el hilo
    window.screens["generate"]._teardown()
    _drain()
    window.screens["generate"].run_now()


def _all_label_text(widget) -> str:
    from PySide6.QtWidgets import QLabel

    return " \n ".join(child.text() for child in widget.findChildren(QLabel))


# ---------------------------------------------------------------------------
# Navegación / arranque
# ---------------------------------------------------------------------------


def test_starts_on_home(window):
    assert window.current_screen_name == "home"
    assert window.screens["home"].start_button.text() == "Nuevo informe"


def test_home_goes_to_new_report(window):
    window.screens["home"].start_button.click()
    assert window.current_screen_name == "new_report"


def test_default_period_is_previous_month(window):
    from remasep.ui.main_window import _default_period

    assert (window.state.month, window.state.year) == _default_period()


def test_step_indicator_advances_through_flow(window, medinet_file, template_file):
    window.navigate("new_report")
    assert window.screens["new_report"].steps.current == 0

    _fill_new_report(window, medinet_file, template_file)
    window.screens["new_report"]._analyze()
    assert window.current_screen_name == "analysis"
    assert window.screens["analysis"].steps.current == 1

    window.screens["analysis"]._teardown()
    _drain()
    window.screens["analysis"].run_now()
    assert window.current_screen_name == "dashboard"
    assert window.screens["dashboard"].steps.current == 2


# ---------------------------------------------------------------------------
# Nuevo informe: selección de archivos + validación
# ---------------------------------------------------------------------------


def test_analyze_disabled_until_both_files_selected(window, medinet_file, template_file):
    window.navigate("new_report")
    nr = window.screens["new_report"]
    assert not nr.analyze_button.isEnabled()

    nr.medinet_selector.set_selection(Path(medinet_file))
    nr._sync_state()
    assert not nr.analyze_button.isEnabled()      # falta la plantilla

    nr.template_selector.set_selection(Path(template_file))
    nr._sync_state()
    assert nr.analyze_button.isEnabled()


def test_file_selector_shows_name_and_ok_state(window, medinet_file):
    window.navigate("new_report")
    sel = window.screens["new_report"].medinet_selector
    sel.set_selection(Path(medinet_file))
    assert sel._name_label.text() == Path(medinet_file).name
    assert sel.has_valid_extension
    assert "✓" in sel._status_label.text()


def test_file_selector_flags_wrong_extension(window, tmp_path):
    window.navigate("new_report")
    bad = tmp_path / "cosa.txt"
    bad.write_text("x")
    nr = window.screens["new_report"]
    nr.medinet_selector.set_selection(bad)
    assert not nr.medinet_selector.has_valid_extension
    nr._sync_state()
    assert not nr.analyze_button.isEnabled()


def test_period_selection_persists_into_state(window, medinet_file, template_file):
    window.navigate("new_report")
    _fill_new_report(window, medinet_file, template_file, month=3, year=2027)
    assert (window.state.month, window.state.year) == (3, 2027)
    assert window.state.period == Period(3, 2027)


# ---------------------------------------------------------------------------
# Análisis -> dashboard
# ---------------------------------------------------------------------------


def test_analysis_produces_summary_that_matches_backend(window, medinet_file, template_file):
    window.navigate("new_report")
    _fill_new_report(window, medinet_file, template_file)
    _run_analysis(window)

    s = window.state.summary
    assert s is not None
    assert window.current_screen_name == "dashboard"
    assert s.period_scope_records == 33          # 20 + 4 + 6 + 3
    assert s.included_records == 24              # 20 Atendido + 4 En Sala de Espera
    assert s.excluded_records == 9
    assert s.included_records + s.excluded_records == s.period_scope_records

    db = window.screens["dashboard"]
    assert db.kpi_period._value.text() == "33"
    assert db.kpi_included._value.text() == "24"
    assert db.kpi_excluded._value.text() == "9"
    assert len(db.chart_estado._bars) == len(s.estado_distribution)
    assert len(db.chart_sex._bars) == len(s.sex_distribution)


def test_analysis_error_shows_human_message_not_crash(window, tmp_path, template_file):
    window.navigate("new_report")
    broken = tmp_path / "roto.xlsx"
    broken.write_text("no soy un excel")
    _fill_new_report(window, broken, template_file)
    _run_analysis(window)

    # no navega al dashboard; muestra un banner de error humano
    assert window.current_screen_name == "analysis"
    err = window.screens["analysis"]._error
    assert err.isVisible()
    assert err._title.text()
    assert "Traceback" not in err._title.text()
    assert "Traceback" not in err._detail.text()
    assert isinstance(window.state.analysis_error, HumanError)


def test_dashboard_declares_medinet_source_and_disclaimer(window, medinet_file, template_file):
    window.navigate("new_report")
    _fill_new_report(window, medinet_file, template_file)
    _run_analysis(window)

    text = _all_label_text(window.screens["dashboard"])
    assert "MEDINET" in text.upper()
    assert "no representa todavía la totalidad" in text
    assert "Tabla quirúrgica" in text  # card "Pendientes de otras fuentes"


def test_dashboard_without_summary_returns_to_new_report(window):
    window.state.summary = None
    window.navigate("dashboard")
    assert window.current_screen_name == "new_report"


# ---------------------------------------------------------------------------
# Generación
# ---------------------------------------------------------------------------


def test_generate_calls_the_generation_service(window, medinet_file, template_file, monkeypatch):
    window.navigate("new_report")
    _fill_new_report(window, medinet_file, template_file)
    _run_analysis(window)

    calls = {}

    def fake_run_generation(medinet, period, template, output, **_kw):
        calls["args"] = (Path(medinet).name, period, Path(template).name, output)
        return GenerationOutcome(
            ok=True, status="GENERATED_DRAFT", submission_label="NOT_FOR_SUBMISSION",
            output_path="outputs/REMASEP_2026_07_DRAFT.xlsm", written_cells=1122,
            considered_records=24, integrity_ok=True,
            control_status="PASS_INTERNAL_VALIDATION", warnings=(),
        )

    monkeypatch.setattr("remasep.ui.workers.run_generation", fake_run_generation)
    monkeypatch.setattr("remasep.ui.screens.generate.run_generation", fake_run_generation)
    _run_generate(window)

    assert calls["args"][0] == Path(medinet_file).name
    assert calls["args"][1] == Period(7, 2026)
    assert calls["args"][2] == Path(template_file).name

    gen = window.screens["generate"]
    assert gen._result_card.isVisible()
    assert not gen._error.isVisible()
    txt = _all_label_text(gen)
    assert "REMASEP_2026_07_DRAFT.xlsm" in txt
    assert "Correcta" in txt
    assert "Julio 2026" in txt
    assert "requerir información adicional" in txt  # aviso NOT_FOR_SUBMISSION


def test_generate_backend_error_shows_human_message(window, medinet_file, template_file, monkeypatch):
    window.navigate("new_report")
    _fill_new_report(window, medinet_file, template_file)
    _run_analysis(window)

    def failing(*_a, **_kw):
        return GenerationOutcome(
            ok=False, status="EXCEL_UNAVAILABLE", submission_label="",
            output_path=None, written_cells=0, considered_records=24,
            integrity_ok=False, control_status="", warnings=(),
            human_error=HumanError(
                "Microsoft Excel Desktop es necesario para generar el informe.",
                "Requiere Windows con Excel de escritorio.", "Volver",
            ),
        )

    monkeypatch.setattr("remasep.ui.workers.run_generation", failing)
    monkeypatch.setattr("remasep.ui.screens.generate.run_generation", failing)
    _run_generate(window)

    gen = window.screens["generate"]
    assert not gen._result_card.isVisible()
    assert gen._error.isVisible()
    assert "Microsoft Excel" in gen._error._title.text()
    assert "Traceback" not in gen._error._detail.text()
    assert window.state.generation is not None and window.state.generation.ok is False


def test_run_generation_is_graceful_without_excel(medinet_file):
    """En Linux / sin Excel: no lanza; devuelve ok=False + mensaje humano."""
    from remasep.ui.workers import run_generation

    outcome = run_generation(medinet_file, Period(7, 2026), "x.xlsm", None)
    assert outcome.ok is False
    assert outcome.status == "EXCEL_UNAVAILABLE"
    assert outcome.human_error is not None
    assert "Excel" in outcome.human_error.title


# ---------------------------------------------------------------------------
# reset_flow
# ---------------------------------------------------------------------------


def test_reset_flow_returns_home_and_clears_selectors(window, medinet_file, template_file):
    window.navigate("new_report")
    _fill_new_report(window, medinet_file, template_file)
    window.reset_flow()

    assert window.current_screen_name == "home"
    nr = window.screens["new_report"]
    assert nr.medinet_selector.path is None
    assert nr.template_selector.path is None
    assert window.state.summary is None
    # conserva el período elegido
    assert (window.state.month, window.state.year) == (7, 2026)


# ---------------------------------------------------------------------------
# Escenario real Julio 2026 (gated)
# ---------------------------------------------------------------------------


@pytest.mark.skipif(not _MEDINET_REAL.is_file(), reason="requiere data/local de desarrollo")
def test_real_july_flow_dashboard_matches_pipeline(window, template_file):
    window.navigate("new_report")
    _fill_new_report(window, _MEDINET_REAL, template_file)
    _run_analysis(window)

    s = window.state.summary
    assert (s.period_scope_records, s.included_records, s.excluded_records) == (2114, 1364, 750)
    db = window.screens["dashboard"]
    assert db.kpi_period._value.text() == "2.114"
    assert db.kpi_included._value.text() == "1.364"
