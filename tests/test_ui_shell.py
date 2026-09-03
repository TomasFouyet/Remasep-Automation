"""Tests de la maqueta de UI (UI Sprint 1).

Qt corre en modo offscreen (ver conftest.py); no se requiere display físico.
No se leen archivos reales ni datos de pacientes.
"""

from __future__ import annotations

import pytest

from remasep.services.mock_remasep import (
    CLASSIFICATION_OPTIONS,
    IGNORE_OPTION,
    AnalysisResult,
    MockRemasepService,
    Period,
    ReviewOutcome,
    project_review,
)
from remasep.ui.styles import format_int

pytestmark = pytest.mark.usefixtures("qapp")


# --- Mock service --------------------------------------------------------


def test_mock_service_analyze_shapes():
    result = MockRemasepService().analyze(Period(8, 2026), demo=True)

    assert isinstance(result, AnalysisResult)
    assert result.total_records == 2006
    # Invariante: total = clasificadas + ignoradas + sin_clasificar.
    assert (
        result.classified_records
        + result.ignored_records
        + result.unclassified_records
        == result.total_records
    )
    assert result.ignored_records == 8
    assert result.unclassified_records == sum(e.count for e in result.exceptions)
    assert result.review_groups == len(result.exceptions) == 3
    # sin_clasificar (N registros) y grupos por revisar (M tipos) son distintos.
    assert result.unclassified_records != result.review_groups

    assert [v.status for v in result.validations] == ["ok"] * 5
    assert [m.name for m in result.modules] == ["REMASEP 01", "B2 ANEXO", "REMASEP OD"]
    assert all(m.detected for m in result.modules)
    assert [s.available for s in result.other_sources] == [False, False]

    for exception in result.exceptions:
        assert exception.possible_categories == CLASSIFICATION_OPTIONS
        assert exception.count >= 1


def test_analysis_result_pending_logic():
    result = MockRemasepService().analyze(Period(8, 2026), demo=True)
    assert result.pending() == 3
    assert not result.can_continue()

    resolved = {e.value for e in result.exceptions}
    assert result.pending(resolved) == 0
    assert result.can_continue(resolved)


def test_analysis_result_invariant_is_enforced():
    from remasep.services.mock_remasep import ExceptionItem

    base = {
        "period": Period(8, 2026),
        "validations": [],
        "modules": [],
        "other_sources": [],
        "exceptions": [ExceptionItem("X", 4, "sin clasificación")],
    }

    # total != clasificadas + ignoradas + sin_clasificar
    with pytest.raises(ValueError, match="inconsistente"):
        AnalysisResult(
            total_records=100, classified_records=90, ignored_records=3,
            unclassified_records=4, **base,
        )

    # suma correcta, pero los ExceptionItem (4) no suman unclassified_records (5)
    with pytest.raises(ValueError, match="ExceptionItem"):
        AnalysisResult(
            total_records=100, classified_records=90, ignored_records=5,
            unclassified_records=5, **base,
        )

    # combinación coherente: no lanza
    ok = AnalysisResult(
        total_records=100, classified_records=91, ignored_records=5,
        unclassified_records=4, **base,
    )
    assert ok.review_groups == 1


def test_period_label():
    assert Period(8, 2026).label == "Agosto 2026"
    assert Period(12, 2025).label == "Diciembre 2025"


# --- Proyección tras revisión ---------------------------------------


def _analysis() -> AnalysisResult:
    return MockRemasepService().analyze(Period(8, 2026), demo=True)


def test_project_review_no_decisions_matches_analysis():
    result = _analysis()
    outcome = project_review(result, {})

    assert isinstance(outcome, ReviewOutcome)
    assert outcome.final_classified == result.classified_records
    assert outcome.final_ignored == result.ignored_records
    assert outcome.final_unclassified == result.unclassified_records
    assert outcome.pending_groups == result.review_groups
    assert outcome.resolved_groups == 0
    assert not outcome.fully_reviewed


def test_project_review_all_classified():
    result = _analysis()
    decisions = {e.value: CLASSIFICATION_OPTIONS[0] for e in result.exceptions}
    outcome = project_review(result, decisions)

    assert outcome.final_unclassified == 0
    assert outcome.pending_groups == 0
    assert outcome.fully_reviewed
    assert outcome.final_ignored == result.ignored_records
    assert outcome.final_classified == result.classified_records + result.unclassified_records
    assert (
        outcome.final_classified + outcome.final_ignored + outcome.final_unclassified
        == result.total_records
    )


def test_project_review_mixed_decisions():
    result = _analysis()
    first, second, third = result.exceptions
    decisions = {
        first.value: CLASSIFICATION_OPTIONS[0],  # clasifica
        third.value: IGNORE_OPTION,  # ignora
        # 'second' queda sin decisión -> pendiente
    }
    outcome = project_review(result, decisions)

    assert outcome.pending_groups == 1
    assert outcome.resolved_groups == 2
    assert outcome.final_unclassified == second.count
    assert outcome.final_ignored == result.ignored_records + third.count
    assert outcome.final_classified == result.classified_records + first.count
    assert (
        outcome.final_classified + outcome.final_ignored + outcome.final_unclassified
        == result.total_records
    )


def test_review_outcome_invariant_enforced():
    with pytest.raises(ValueError, match="ReviewOutcome inconsistente"):
        ReviewOutcome(
            total_records=100,
            final_classified=50,
            final_ignored=10,
            final_unclassified=10,
            resolved_groups=0,
            pending_groups=0,
        )


def test_appstate_review_outcome_none_without_analysis(window):
    assert window.state.analysis is None
    assert window.state.review_outcome() is None


# --- MainWindow / navegación --------------------------------------


@pytest.fixture
def window():
    from remasep.ui.main_window import MainWindow

    return MainWindow()


def test_main_window_starts_on_home(window):
    assert window.current_screen_name == "home"
    assert set(window.screens) == {"home", "new_report", "analysis", "exceptions", "summary"}


def test_home_button_navigates_to_new_report(window):
    window.screens["home"].start_button.click()
    assert window.current_screen_name == "new_report"


def test_default_period_is_previous_month(window):
    # conftest no congela la fecha; solo comprobamos coherencia del estado.
    assert 1 <= window.state.month <= 12
    assert window.state.year >= 2024


def test_default_period_uses_local_previous_month():
    import datetime as dt

    from remasep.ui.main_window import _default_period

    today = dt.date.today()  # noqa: DTZ011 - se compara contra la misma base local
    expected = today.replace(day=1) - dt.timedelta(days=1)
    assert _default_period() == (expected.month, expected.year)


def test_period_selection_persists_through_flow(window):
    new_report = window.screens["home"].start_button.click() or window.screens["new_report"]
    new_report.month_combo.setCurrentIndex(7)  # Agosto
    new_report.year_spin.setValue(2026)

    assert window.state.month == 8
    assert window.state.year == 2026

    new_report.demo_button.click()
    assert window.current_screen_name == "analysis"
    assert window.state.analysis.period_label == "Agosto 2026"
    assert "Agosto 2026" in window.screens["analysis"].period_label.text()


def test_demo_mode_runs_analysis(window):
    window.navigate("new_report")
    window.screens["new_report"].demo_button.click()

    assert window.state.demo is True
    assert window.state.analysis is not None
    assert window.current_screen_name == "analysis"


def test_default_period_survives_new_report_screen(window):
    # Entrar a 'Nuevo reporte' no debe pisar el año/mes por defecto del estado.
    expected_month, expected_year = window.state.month, window.state.year
    window.navigate("new_report")
    window.screens["new_report"].demo_button.click()

    assert (window.state.month, window.state.year) == (expected_month, expected_year)
    assert window.state.analysis.period_label.endswith(str(expected_year))


def test_analyze_button_requires_medinet_file(window, tmp_path):
    window.navigate("new_report")
    new_report = window.screens["new_report"]
    assert new_report.analyze_button.isEnabled() is False

    sample = tmp_path / "medinet.xlsx"
    sample.write_bytes(b"not a real workbook")
    new_report.medinet_selector.set_selection(sample)
    assert new_report.analyze_button.isEnabled() is True


# --- Componentes -------------------------------------------------------


def test_file_selector_shows_name_and_size(window, tmp_path):
    from remasep.ui.components.file_selector import FileSelector, human_size

    sample = tmp_path / "medinet.xlsx"
    sample.write_bytes(b"x" * 2048)

    selector = FileSelector("Medinet", extensions=["xlsx", "csv"])
    captured = []
    selector.fileSelected.connect(lambda p: captured.append(p))
    selector.set_selection(sample)

    assert selector.path == sample
    assert "medinet.xlsx" in selector._name_label.text()
    assert human_size(2048) == "2,0 KB"
    assert captured == [sample]

    selector.clear()
    assert selector.path is None


def test_step_indicator_marks_current(window):
    from remasep.ui.components.step_indicator import StepIndicator

    steps = StepIndicator(["Archivos", "Análisis", "Revisión", "Resultado"])
    steps.set_current(2)
    assert steps.current == 2
    assert steps._labels[1].objectName() == "stepDone"
    assert steps._labels[2].objectName() == "stepActive"
    assert steps._labels[3].objectName() == "stepInactive"


# --- Flujo completo -------------------------------------------------


def test_full_flow_home_to_summary(window):
    # HOME -> NUEVO REPORTE
    window.screens["home"].start_button.click()
    new_report = window.screens["new_report"]
    new_report.month_combo.setCurrentIndex(7)
    new_report.year_spin.setValue(2026)

    # -> ANÁLISIS (modo demo)
    new_report.demo_button.click()
    assert window.current_screen_name == "analysis"

    # -> EXCEPCIONES
    window.screens["analysis"].review_button.click()
    exceptions = window.screens["exceptions"]
    assert window.current_screen_name == "exceptions"
    assert len(exceptions.rows) == 3
    assert exceptions.continue_button.isEnabled() is False  # bloqueado por pendientes

    # resolver: 2 excepciones a una categoría, 1 como "Ignorar justificadamente"
    ignore_index = 1 + CLASSIFICATION_OPTIONS.index(IGNORE_OPTION)
    exceptions.rows[0].combo.setCurrentIndex(1)
    exceptions.rows[1].combo.setCurrentIndex(1)
    exceptions.rows[2].combo.setCurrentIndex(ignore_index)
    assert exceptions.continue_button.isEnabled() is True

    # -> RESUMEN
    exceptions.continue_button.click()
    summary = window.screens["summary"]
    assert window.current_screen_name == "summary"
    assert "AGOSTO 2026" in summary.title_label.text()

    # generación deshabilitada + tooltip
    assert summary.generate_button.isEnabled() is False
    assert "Microsoft Excel" in summary.generate_button.toolTip()

    result = window.state.analysis
    outcome = window.state.review_outcome()

    # RESULTADO DEL ANÁLISIS: intacto, muestra las cifras originales.
    assert _find_row(summary._analysis_layout, "Sin clasificar").value == format_int(
        result.unclassified_records
    )
    assert _find_row(summary._analysis_layout, "Clasificadas").value == format_int(
        result.classified_records
    )

    # ESTADO TRAS LA REVISIÓN: sin pendientes y sin registros sin clasificar.
    assert _find_row(summary._review_layout, "Grupos pendientes de revisión").value == "0"
    assert _find_row(summary._review_layout, "Sin clasificar").value == "0"
    assert _find_row(summary._review_layout, "Clasificadas").value == format_int(
        outcome.final_classified
    )
    assert _find_row(summary._review_layout, "Ignoradas").value == format_int(outcome.final_ignored)

    # Invariante: total = final_classified + final_ignored + final_unclassified.
    assert (
        outcome.final_classified + outcome.final_ignored + outcome.final_unclassified
        == result.total_records
    )
    # 2 grupos clasificados + 1 ignorado respecto de las cifras del análisis.
    ignored_group_count = result.exceptions[2].count
    assert outcome.final_ignored == result.ignored_records + ignored_group_count
    assert outcome.final_classified == result.total_records - outcome.final_ignored

    # -> HOME conservando período
    summary.home_button.click()
    assert window.current_screen_name == "home"
    assert window.state.month == 8
    assert window.state.year == 2026
    assert window.state.analysis is None
    assert window.state.demo is False


def test_back_navigation(window):
    window.navigate("new_report")
    window.screens["new_report"].demo_button.click()
    window.screens["analysis"].review_button.click()
    assert window.current_screen_name == "exceptions"

    window.screens["exceptions"].back_button.click()
    assert window.current_screen_name == "analysis"

    window.screens["analysis"].back_button.click()
    assert window.current_screen_name == "new_report"


def _find_row(layout, title: str):
    for i in range(layout.count()):
        widget = layout.itemAt(i).widget()
        if widget is not None and getattr(widget, "title", None) == title:
            return widget
    raise AssertionError(f"fila {title!r} no encontrada")
