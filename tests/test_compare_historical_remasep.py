"""Tests del diagnóstico de validación histórica REMASEP (sólo lectura).

`scripts/compare_historical_remasep.py` reutiliza `build_production_pending_writes`
(pipeline productivo) y `compare_reference_values` (Sprint 3.7A) — no reimplementa
ningún motor de cálculo. Estos tests cubren la clasificación celda a celda, la
resolución del archivo histórico (sin matching ambiguo silencioso), que no se
muta ningún archivo histórico, determinismo, ausencia de PII en las filas de
salida y que el script efectivamente reutiliza el pipeline productivo.
"""

from __future__ import annotations

import hashlib
from dataclasses import dataclass
from pathlib import Path
from types import SimpleNamespace

import compare_historical_remasep as chr_script
import openpyxl
import pytest

from remasep.services import production_pipeline
from remasep.services.common import Period
from remasep.services.metric_value_producer import PendingWrite


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


# ---------------------------------------------------------------------------
# classify_comparison — las 4 categorías
# ---------------------------------------------------------------------------


def test_classify_comparison_exact_match():
    assert (
        chr_script.classify_comparison(5, 5, reference_available=True)
        == chr_script.STATUS_EXACT_MATCH
    )


def test_classify_comparison_reference_blank_and_calculated_zero_is_zero_vs_blank():
    """Caso (1) pedido: referencia vacía/None (leída con éxito) + calculado 0."""
    assert (
        chr_script.classify_comparison(0, None, reference_available=True)
        == chr_script.STATUS_ZERO_VS_BLANK
    )
    # simétrico: calculado 0, histórico también 0 -> EXACT_MATCH (no ambiguo)
    assert (
        chr_script.classify_comparison(0, 0, reference_available=True)
        == chr_script.STATUS_EXACT_MATCH
    )


def test_classify_comparison_real_mismatch():
    assert (
        chr_script.classify_comparison(5, 3, reference_available=True)
        == chr_script.STATUS_REAL_MISMATCH
    )


def test_classify_comparison_reference_blank_and_calculated_nonzero_is_real_mismatch():
    """Caso (2) pedido: referencia vacía/None LEÍDA CON ÉXITO (archivo, hoja y
    celda disponibles) + calculado no-cero -> REAL_MISMATCH, NO
    REFERENCE_UNAVAILABLE. Una celda vacía frente a un conteo esperado no-cero
    es una discrepancia real, no una referencia indisponible."""
    assert (
        chr_script.classify_comparison(5, None, reference_available=True)
        == chr_script.STATUS_REAL_MISMATCH
    )


def test_classify_comparison_reference_unavailable_when_file_missing():
    """Caso (3) pedido: sin archivo histórico disponible -> SIEMPRE
    REFERENCE_UNAVAILABLE, incluso si el valor calculado es 0 (no se asume
    equivalencia con un archivo ausente; no se confunde con una celda vacía)."""
    assert (
        chr_script.classify_comparison(0, None, reference_available=False)
        == chr_script.STATUS_REFERENCE_UNAVAILABLE
    )
    assert (
        chr_script.classify_comparison(5, None, reference_available=False)
        == chr_script.STATUS_REFERENCE_UNAVAILABLE
    )


# ---------------------------------------------------------------------------
# resolve_reference_file — sin matching ambiguo silencioso
# ---------------------------------------------------------------------------


def test_resolve_reference_file_exact_name_match(tmp_path):
    requested = "REMASEP_V1.4 Julio 2026.xlsm"
    (tmp_path / requested).write_bytes(b"fake")
    result = chr_script.resolve_reference_file(tmp_path, 7, requested)
    assert result.available
    assert result.match_kind == "EXACT_NAME_MATCH"
    assert result.path == tmp_path / requested


def test_resolve_reference_file_not_found_when_directory_empty(tmp_path):
    result = chr_script.resolve_reference_file(tmp_path, 4, "REMASEP 2026 V1.4 Abril 2026.xlsm")
    assert not result.available
    assert result.match_kind == "NOT_FOUND"


def test_resolve_reference_file_case_insensitive_single_candidate(tmp_path):
    requested = "REMASEP_V1.4 Julio 2026.xlsm"
    (tmp_path / requested.upper()).write_bytes(b"fake")
    result = chr_script.resolve_reference_file(tmp_path, 7, requested)
    assert result.available
    assert result.match_kind == "CASE_INSENSITIVE_MATCH"


def test_resolve_reference_file_ambiguous_does_not_pick_silently(tmp_path):
    (tmp_path / "REMASEP_V1.4 Junio_2026_v2.xlsm").write_bytes(b"fake")
    (tmp_path / "REMASEP Junio 2026 copia.xlsm").write_bytes(b"fake")
    result = chr_script.resolve_reference_file(tmp_path, 6, "REMASEP 2026 V1.4 Junio 2026.xlsm")
    assert not result.available
    assert result.match_kind == "AMBIGUOUS"
    assert len(result.ignored_candidates) == 2


def test_resolve_reference_file_exact_match_wins_over_extra_candidates(tmp_path):
    """Julio real: existe el nombre exacto solicitado Y un archivo extra
    parecido (`2026-7 REMASEP_V1.4.xlsm`). Debe usarse el exacto, sin ambigüedad."""
    requested = "REMASEP_V1.4 Julio 2026.xlsm"
    (tmp_path / requested).write_bytes(b"exact")
    (tmp_path / "2026-7 REMASEP_V1.4.xlsm").write_bytes(b"extra")
    result = chr_script.resolve_reference_file(tmp_path, 7, requested)
    assert result.match_kind == "EXACT_NAME_MATCH"
    assert result.path.read_bytes() == b"exact"


# ---------------------------------------------------------------------------
# read_target_cells — distingue "vacío pero leído" de "no se pudo obtener"
# ---------------------------------------------------------------------------


def test_read_target_cells_blank_cell_is_available_with_none_value(tmp_path):
    path = tmp_path / "ref.xlsx"
    _build_reference_workbook(path, "S", {"A1": 5})  # A2 nunca se escribe -> blank
    reads = chr_script.read_target_cells(path, {("S", "A1"), ("S", "A2")})
    assert reads[("S", "A1")] == chr_script.CellRead(True, 5)
    assert reads[("S", "A2")].available is True
    assert reads[("S", "A2")].value is None


def test_read_target_cells_workbook_missing_is_unavailable(tmp_path):
    """Caso (3) pedido: archivo ausente -> CellRead no disponible, con motivo,
    nunca una excepción sin controlar ni un valor fabricado."""
    missing = tmp_path / "no_existe.xlsm"
    reads = chr_script.read_target_cells(missing, {("S", "A1")})
    read = reads[("S", "A1")]
    assert read.available is False
    assert read.value is None
    assert "WORKBOOK_OPEN_ERROR" in read.reason


def test_read_target_cells_sheet_missing_is_unavailable(tmp_path):
    """Caso (4) pedido: la hoja no existe en el workbook -> no disponible."""
    path = tmp_path / "ref.xlsx"
    _build_reference_workbook(path, "OTHER_SHEET", {"A1": 5})
    reads = chr_script.read_target_cells(path, {("S", "A1")})
    read = reads[("S", "A1")]
    assert read.available is False
    assert read.reason == "SHEET_NOT_FOUND"


def test_read_target_cells_unreadable_cell_reference_is_unavailable(tmp_path):
    """Caso (4) pedido: la celda en sí no se puede leer (referencia inválida)."""
    path = tmp_path / "ref.xlsx"
    _build_reference_workbook(path, "S", {"A1": 5})
    reads = chr_script.read_target_cells(path, {("S", "not-a-cell-ref")})
    read = reads[("S", "not-a-cell-ref")]
    assert read.available is False
    assert "CELL_READ_ERROR" in read.reason


def test_run_month_blocked_status_never_calls_read_target_cells(tmp_path, monkeypatch):
    """Si el mes queda BLOCKED por invariantes, no se intenta leer ningún
    histórico (no se fabrica ninguna comparación)."""
    calls = []
    monkeypatch.setattr(
        chr_script, "read_target_cells", lambda *a, **k: calls.append((a, k)) or {}
    )
    pending = [PendingWrite("wi:1", "m1", "S", "A1", 5, "INTEGER_COUNT", "Julio 2026")]

    def fake_pipeline(export_path, period, *, bundle=None):
        return _fake_result(pending)

    chr_script.run_month(
        tmp_path / "export.xlsx", tmp_path, 7, 2026,
        _FakeBundle(expected_instruction_count=1122), {},
        pipeline=fake_pipeline,
    )
    assert calls == []


# ---------------------------------------------------------------------------
# run_month — invariantes / BLOCKED / no mutación / determinismo
# ---------------------------------------------------------------------------


@dataclass
class _FakeBundle:
    expected_instruction_count: int


def _fake_scope(**overrides):
    base = {
        "processing_scope_records": 10,
        "period_scope_records": 12,
        "estado_excluded_records": 2,
        "estado_filter_status": "CONFIRMED",
        "estado_filter_applied": True,
    }
    base.update(overrides)
    return SimpleNamespace(**base)


def _fake_completeness(ok=True):
    return SimpleNamespace(ok=ok, missing=(), duplicate=(), orphan=())


def _fake_result(pending_writes, *, scope=None, completeness=None):
    return SimpleNamespace(
        pending_writes=pending_writes,
        scope=scope or _fake_scope(),
        completeness=completeness or _fake_completeness(),
    )


def test_run_month_blocked_when_pending_write_count_is_wrong(tmp_path):
    pending = [
        PendingWrite("wi:1", "m1", "S", "A1", 5, "INTEGER_COUNT", "Julio 2026"),
    ]

    def fake_pipeline(export_path, period, *, bundle=None):
        return _fake_result(pending)

    outcome = chr_script.run_month(
        tmp_path / "export.xlsx", tmp_path, 7, 2026,
        _FakeBundle(expected_instruction_count=1122), {},
        pipeline=fake_pipeline,
    )
    assert outcome.status == "BLOCKED"
    assert "pending_writes=1 != esperado 1122" in outcome.block_reason
    assert outcome.detail_rows == []


def test_run_month_blocked_when_estado_filter_not_confirmed(tmp_path):
    pending = [PendingWrite("wi:1", "m1", "S", "A1", 5, "INTEGER_COUNT", "Julio 2026")]

    def fake_pipeline(export_path, period, *, bundle=None):
        return _fake_result(pending, scope=_fake_scope(estado_filter_status="PENDING_FUNCTIONAL_CONFIRMATION", estado_filter_applied=False))

    outcome = chr_script.run_month(
        tmp_path / "export.xlsx", tmp_path, 7, 2026,
        _FakeBundle(expected_instruction_count=1), {},
        pipeline=fake_pipeline,
    )
    assert outcome.status == "BLOCKED"
    assert "estado_filter_status" in outcome.block_reason
    assert outcome.detail_rows == []


def _build_reference_workbook(path: Path, sheet: str, cells: dict[str, object]) -> None:
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = sheet
    for cell, value in cells.items():
        ws[cell] = value
    wb.save(path)


def _five_scenarios_pending(sheet: str) -> list[PendingWrite]:
    return [
        PendingWrite("wi:1", "m1", sheet, "A1", 5, "INTEGER_COUNT", "Julio 2026"),  # exact match
        PendingWrite("wi:2", "m2", sheet, "A2", 0, "INTEGER_COUNT", "Julio 2026"),  # blank + calc 0
        PendingWrite("wi:3", "m3", sheet, "A3", 5, "INTEGER_COUNT", "Julio 2026"),  # numeric mismatch
        PendingWrite("wi:4", "m4", sheet, "A4", 7, "INTEGER_COUNT", "Julio 2026"),  # blank + calc != 0
        PendingWrite("wi:5", "m5", "MISSING_SHEET", "A1", 9, "INTEGER_COUNT", "Julio 2026"),  # hoja ausente
    ]


def test_run_month_classifies_all_four_categories_end_to_end(tmp_path):
    """Cubre las 4 categorías, incluyendo los dos casos REAL_MISMATCH pedidos:
    numérico (wi:3) y referencia vacía frente a calculado no-cero (wi:4)."""
    requested_name = chr_script.REQUESTED_REFERENCE_FILES[7]
    ref_path = tmp_path / requested_name
    # A2 y A4 nunca se escriben -> quedan blank (None), leídas con éxito.
    _build_reference_workbook(ref_path, "S", {"A1": 5, "A3": 3})

    pending = _five_scenarios_pending("S")

    def fake_pipeline(export_path, period, *, bundle=None):
        assert period == Period(7, 2026)
        return _fake_result(pending)

    outcome = chr_script.run_month(
        tmp_path / "export.xlsx", tmp_path, 7, 2026,
        _FakeBundle(expected_instruction_count=5), {},
        pipeline=fake_pipeline,
    )
    assert outcome.status == "COMPARED"
    statuses = {row["instruction_id"]: row["comparison_status"] for row in outcome.detail_rows}
    assert statuses["wi:1"] == chr_script.STATUS_EXACT_MATCH
    assert statuses["wi:2"] == chr_script.STATUS_ZERO_VS_BLANK
    assert statuses["wi:3"] == chr_script.STATUS_REAL_MISMATCH
    assert statuses["wi:4"] == chr_script.STATUS_REAL_MISMATCH  # blank + no-cero
    assert statuses["wi:5"] == chr_script.STATUS_REFERENCE_UNAVAILABLE  # hoja ausente

    rows_by_id = {row["instruction_id"]: row for row in outcome.detail_rows}
    assert rows_by_id["wi:3"]["delta"] == 3 - 5
    assert rows_by_id["wi:4"]["delta"] is None  # no se fabrica un 0 para la celda vacía
    assert rows_by_id["wi:4"]["reference_available"] is True
    assert rows_by_id["wi:5"]["reference_available"] is False
    assert "SHEET_NOT_FOUND" in rows_by_id["wi:5"]["reference_unavailable_reason"]


def test_run_month_does_not_mutate_reference_workbook(tmp_path):
    requested_name = chr_script.REQUESTED_REFERENCE_FILES[7]
    ref_path = tmp_path / requested_name
    _build_reference_workbook(ref_path, "S", {"A1": 5})
    sha_before = _sha256(ref_path)

    def fake_pipeline(export_path, period, *, bundle=None):
        return _fake_result([PendingWrite("wi:1", "m1", "S", "A1", 5, "INTEGER_COUNT", "Julio 2026")])

    chr_script.run_month(
        tmp_path / "export.xlsx", tmp_path, 7, 2026,
        _FakeBundle(expected_instruction_count=1), {},
        pipeline=fake_pipeline,
    )
    assert _sha256(ref_path) == sha_before


def test_run_month_is_deterministic(tmp_path):
    requested_name = chr_script.REQUESTED_REFERENCE_FILES[7]
    ref_path = tmp_path / requested_name
    _build_reference_workbook(ref_path, "S", {"A1": 5, "A3": 3})
    pending = _five_scenarios_pending("S")

    def fake_pipeline(export_path, period, *, bundle=None):
        return _fake_result(pending)

    args = (
        tmp_path / "export.xlsx", tmp_path, 7, 2026,
        _FakeBundle(expected_instruction_count=5), {},
    )
    first = chr_script.run_month(*args, pipeline=fake_pipeline)
    second = chr_script.run_month(*args, pipeline=fake_pipeline)
    assert first.detail_rows == second.detail_rows
    assert first.status == second.status == "COMPARED"


def test_run_month_reference_unavailable_when_file_not_found(tmp_path):
    pending = [PendingWrite("wi:1", "m1", "S", "A1", 5, "INTEGER_COUNT", "Abril 2026")]

    def fake_pipeline(export_path, period, *, bundle=None):
        return _fake_result(pending)

    outcome = chr_script.run_month(
        tmp_path / "export.xlsx", tmp_path, 4, 2026,
        _FakeBundle(expected_instruction_count=1), {},
        pipeline=fake_pipeline,
    )
    assert outcome.status == "COMPARED"
    assert not outcome.reference.available
    assert outcome.detail_rows[0]["comparison_status"] == chr_script.STATUS_REFERENCE_UNAVAILABLE
    assert outcome.detail_rows[0]["reference_file"] == ""


# ---------------------------------------------------------------------------
# Privacidad: nunca se exponen columnas de datos individuales
# ---------------------------------------------------------------------------

_FORBIDDEN_FIELD_SUBSTRINGS = (
    "RUN", "RUT", "NOMBRE", "APELLIDO", "TELEFONO", "DIRECCION",
    "FECHA_NACIMIENTO", "DIA_CITA", "PACIENTE",
)


def test_detail_columns_never_expose_individual_record_fields():
    for name in chr_script._DETAIL_FIELDS:
        upper = name.upper()
        for forbidden in _FORBIDDEN_FIELD_SUBSTRINGS:
            assert forbidden not in upper, f"{name!r} parece exponer datos individuales"


def test_detail_rows_are_aggregate_counts_not_raw_records(tmp_path):
    requested_name = chr_script.REQUESTED_REFERENCE_FILES[7]
    ref_path = tmp_path / requested_name
    _build_reference_workbook(ref_path, "S", {"A1": 5})
    pending = [PendingWrite("wi:1", "m1", "S", "A1", 5, "INTEGER_COUNT", "Julio 2026")]

    def fake_pipeline(export_path, period, *, bundle=None):
        return _fake_result(pending)

    outcome = chr_script.run_month(
        tmp_path / "export.xlsx", tmp_path, 7, 2026,
        _FakeBundle(expected_instruction_count=1), {},
        pipeline=fake_pipeline,
    )
    row = outcome.detail_rows[0]
    assert set(row) == set(chr_script._DETAIL_FIELDS)
    assert isinstance(row["calculated_value"], int)


# ---------------------------------------------------------------------------
# El script reutiliza el pipeline productivo: no hay un segundo motor
# ---------------------------------------------------------------------------


def test_script_reuses_production_pipeline_function_directly():
    assert chr_script.build_production_pending_writes is production_pipeline.build_production_pending_writes
    import inspect

    default_pipeline = inspect.signature(chr_script.run_month).parameters["pipeline"].default
    assert default_pipeline is production_pipeline.build_production_pending_writes


def test_historical_regression_uses_real_pipeline_end_to_end(make_medinet):
    """Sin override de `pipeline`, `historical_regression` corre el pipeline
    productivo real (no un motor de cálculo paralelo) sobre un export sintético."""
    from datetime import date

    rows = [
        [date(2026, 7, 10), date(1990, 1, 1), "Mujer", "Santiago", "ODONTOLOGIA",
         "CONSULTA GENERAL", "EVALUACIÓN ODONTOLÓGICA", "Atendido", "Presencial", ""],
    ]
    medinet = make_medinet(rows)
    outcomes = chr_script.historical_regression(medinet, "data/local", months=[7])
    assert len(outcomes) == 1
    assert outcomes[0].status == "COMPARED"
    assert outcomes[0].pending_write_count == 1122


# ---------------------------------------------------------------------------
# Escenario real Julio 2026 (sólo si data/local está disponible)
# ---------------------------------------------------------------------------

_MEDINET_REAL = Path("data/local/detalle_citas - 2026-09-07T123630.940.xlsx")
_DATA_LOCAL = Path("data/local")
_local_available = _MEDINET_REAL.is_file()


_local_all_references_available = _local_available and all(
    chr_script.resolve_reference_file(_DATA_LOCAL, month, name).available
    for month, name in chr_script.REQUESTED_REFERENCE_FILES.items()
)

# Números conocidos (Sprint de validación histórica Abril-Julio 2026, con las
# 3 referencias históricas agregadas a data/local). Si `data/local` cambia
# estos números pueden dejar de aplicar; el test se salta si faltan archivos.
_KNOWN_MONTH_RESULTS = {
    4: {"exact_match": 994, "zero_vs_blank": 1, "real_mismatch": 127, "unavailable": 0},
    5: {"exact_match": 984, "zero_vs_blank": 2, "real_mismatch": 136, "unavailable": 0},
    6: {"exact_match": 974, "zero_vs_blank": 1, "real_mismatch": 147, "unavailable": 0},
    7: {"exact_match": 980, "zero_vs_blank": 0, "real_mismatch": 142, "unavailable": 0},
}


@pytest.mark.skipif(not _local_available, reason="requiere data/local de desarrollo")
def test_july_2026_historical_regression_matches_known_numbers():
    outcomes = chr_script.historical_regression(_MEDINET_REAL, _DATA_LOCAL, months=[7])
    assert len(outcomes) == 1
    outcome = outcomes[0]
    assert outcome.status == "COMPARED"
    assert outcome.reference.available
    assert outcome.reference.match_kind == "EXACT_NAME_MATCH"
    row = chr_script.month_summary_row(outcome)
    expected = _KNOWN_MONTH_RESULTS[7]
    assert row["compared"] == 1122
    assert row["exact_match"] == expected["exact_match"]
    assert row["zero_vs_blank"] == expected["zero_vs_blank"]
    assert row["real_mismatch"] == expected["real_mismatch"]
    assert row["unavailable"] == expected["unavailable"]
    assert row["equivalence_rate"] == pytest.approx(
        (expected["exact_match"] + expected["zero_vs_blank"]) / 1122
    )


@pytest.mark.skipif(
    not _local_all_references_available,
    reason="requiere las 4 referencias históricas en data/local",
)
def test_april_to_july_2026_historical_regression_matches_known_numbers():
    """Con las 4 referencias históricas presentes, ningún mes debe quedar
    BLOCKED ni con UNAVAILABLE espurio: una celda vacía leída con éxito es una
    referencia válida (ZERO_VS_BLANK o REAL_MISMATCH según el calculado), nunca
    REFERENCE_UNAVAILABLE. No se fabrica ninguna coincidencia."""
    outcomes = chr_script.historical_regression(_MEDINET_REAL, _DATA_LOCAL, months=[4, 5, 6, 7])
    assert len(outcomes) == 4
    for outcome in outcomes:
        expected = _KNOWN_MONTH_RESULTS[outcome.month]
        assert outcome.status == "COMPARED"
        assert outcome.reference.available
        assert outcome.completeness_ok
        assert outcome.pending_write_count == 1122
        row = chr_script.month_summary_row(outcome)
        assert row["compared"] == 1122
        assert row["exact_match"] == expected["exact_match"]
        assert row["zero_vs_blank"] == expected["zero_vs_blank"]
        assert row["real_mismatch"] == expected["real_mismatch"]
        assert row["unavailable"] == expected["unavailable"]
        assert (
            row["exact_match"] + row["zero_vs_blank"] + row["real_mismatch"] + row["unavailable"]
            == 1122
        )
