"""Sprint 3.7B — pruebas del núcleo del Excel writer.

Sin Excel real, sin ``data/local``. Se usa ``FakeWriterHarness`` para ejercitar
copy-first → escritura → recálculo → verificación → promoción atómica.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from remasep.services import excel_writer as w
from remasep.services.metric_value_producer import PendingWrite
from remasep.testing import FakeWorkbookModel, FakeWriterHarness

FP = "stf:dc624775927d4d4d"
SHEETS = ["REMASEP_OD", "REMASEP 01", "B2 ANEXO", "CONTROL"]

_CONTROL_MAP = w.load_control_map(
    {
        "sheet": "CONTROL",
        "total_errors_cell": "E16",
        "data_status_ok_value": "OK",
        "data_status_missing_value": "SIN DATOS",
        "modules_not_yet_produced": ["URGENCIAS", "EyP_ET"],
        "rows": [
            {"sheet": "REMASEP 01", "label_cell": "D7", "errors_cell": "E7",
             "data_status_cell": "F7", "causal_cell": "G7"},
            {"sheet": "URGENCIAS", "label_cell": "D8", "errors_cell": "E8",
             "data_status_cell": "F8", "causal_cell": "G8"},
        ],
    }
)


def _pw(pairs) -> list[PendingWrite]:
    return [
        PendingWrite(f"wi:{i}", f"LEGACY::X::{i}", sheet, cell, value, "INTEGER_COUNT", "Julio 2026")
        for i, (sheet, cell, value) in enumerate(pairs)
    ]


def _model(targets, **kw) -> FakeWorkbookModel:
    base = {
        "sheet_names": list(SHEETS),
        "formula_map": {("CONTROL", "E16"): "=E7+E8"},
        "values": {("CONTROL", "E16"): 0, ("CONTROL", "E7"): 0, ("CONTROL", "E8"): 0},
        "writable_cells": {(s, c) for s, c, _ in targets},
        "structural_fingerprint_id": FP,
    }
    base.update(kw)
    return FakeWorkbookModel(**base)


def _request(tmp_path, pending, *, mode=w.MODE_DIAGNOSTIC_REFERENCE, fp=FP, output=None):
    template = tmp_path / "tpl.xlsm"
    template.write_bytes(b"PK\x03\x04 placeholder xlsm")
    out = output or (tmp_path / "REMASEP_2026_07_DRAFT.xlsm")
    return w.GenerationRequest(
        mode=mode,
        template_path=template,
        output_path=out,
        period_year=2026,
        period_month=7,
        expected_fingerprint_id=fp,
        zero_write_policy="WRITE_ZERO",
        zero_write_policy_accepted=("WRITE_ZERO", "FORM_SPECIFIC"),
        manifest_instruction_ids=tuple(p.instruction_id for p in pending),
        forbidden_dirs=(Path("data/local"), Path("data")),
    )


def _generate(tmp_path, pending, model, **req_kw):
    harness = FakeWriterHarness(model)
    request = _request(tmp_path, pending, **req_kw)
    result = w.generate(
        request, pending, open_writer=harness.open_writer,
        control_map=_CONTROL_MAP, inspect=harness.inspect,
    )
    return result, request, harness


# ---------------------------------------------------------------------------
# Camino feliz + cero explícito + atómico
# ---------------------------------------------------------------------------


def test_happy_path_writes_all_targets_and_promotes_atomically(tmp_path):
    targets = [("REMASEP_OD", "B10", 7), ("REMASEP_OD", "B11", 0), ("REMASEP 01", "C5", 3)]
    pending = _pw(targets)
    result, request, _ = _generate(tmp_path, pending, _model(targets))

    assert result.status == w.STATUS_GENERATED_DIAGNOSTIC
    assert result.writer_integrity_status == w.WRITER_INTEGRITY_PASS
    assert result.written_cells == 3
    assert result.output_path == str(request.output_path)
    assert request.output_path.exists()
    working = request.output_path.with_name("REMASEP_2026_07_DRAFT.__working__.xlsm")
    assert not working.exists()
    assert result.template_unchanged is True
    assert result.submission_label == w.SUBMISSION_LABEL


def test_zero_is_written_explicitly_not_as_blank(tmp_path):
    targets = [("REMASEP_OD", "B10", 0)]
    pending = _pw(targets)
    _, _, harness = _generate(tmp_path, pending, _model(targets))
    assert ("REMASEP_OD", "B10", 0) in harness.last_writer.writes
    assert harness.last_writer.snapshot().values[("REMASEP_OD", "B10")] == 0


def test_production_mode_status_is_draft(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    result, _, _ = _generate(tmp_path, pending, _model(targets), mode=w.MODE_PRODUCTION)
    assert result.status == w.STATUS_GENERATED_DRAFT


# ---------------------------------------------------------------------------
# Preflight (§4)
# ---------------------------------------------------------------------------


def test_preflight_rejects_duplicate_target_cell(tmp_path):
    targets = [("REMASEP_OD", "B10", 1), ("REMASEP_OD", "B10", 2)]
    pending = _pw(targets)
    result, request, _ = _generate(tmp_path, pending, _model(targets))
    assert result.status == w.STATUS_ABORTED_PREFLIGHT
    assert "DUPLICATE_TARGET_CELL" in result.errors
    assert not request.output_path.exists()


def test_preflight_detects_missing_write_instructions(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    request = _request(tmp_path, pending)
    request = w.GenerationRequest(
        **{**request.__dict__, "manifest_instruction_ids": ("wi:0", "wi:999")}
    )
    harness = FakeWriterHarness(_model(targets))
    result = w.generate(request, pending, open_writer=harness.open_writer,
                        control_map=_CONTROL_MAP, inspect=harness.inspect)
    assert result.status == w.STATUS_ABORTED_PREFLIGHT
    assert "MISSING_WRITE_INSTRUCTIONS" in result.errors


def test_preflight_rejects_wrong_value_type(tmp_path):
    targets = [("REMASEP_OD", "B10", "3")]
    pending = _pw(targets)
    result, _, _ = _generate(tmp_path, pending, _model(targets))
    assert result.status == w.STATUS_ABORTED_PREFLIGHT
    assert "VALUE_TYPE_MISMATCH" in result.errors


def test_preflight_rejects_negative_count(tmp_path):
    targets = [("REMASEP_OD", "B10", -1)]
    pending = _pw(targets)
    result, _, _ = _generate(tmp_path, pending, _model(targets))
    assert result.status == w.STATUS_ABORTED_PREFLIGHT
    assert "NEGATIVE_COUNT" in result.errors


def test_preflight_rejects_incompatible_zero_policy(tmp_path):
    targets = [("REMASEP_OD", "B10", 0)]
    pending = _pw(targets)
    request = _request(tmp_path, pending)
    request = w.GenerationRequest(**{**request.__dict__, "zero_write_policy": "UNRESOLVED"})
    harness = FakeWriterHarness(_model(targets))
    result = w.generate(request, pending, open_writer=harness.open_writer,
                        control_map=_CONTROL_MAP, inspect=harness.inspect)
    assert result.status == w.STATUS_ABORTED_PREFLIGHT
    assert "ZERO_WRITE_POLICY_INCOMPATIBLE" in result.errors


def test_preflight_refuses_output_inside_data_local(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    result, _, _ = _generate(
        tmp_path, pending, _model(targets), output=Path("data/local/x.xlsm")
    )
    assert result.status == w.STATUS_ABORTED_PREFLIGHT
    assert "OUTPUT_IN_FORBIDDEN_DIR" in result.errors


def test_existing_output_is_not_overwritten(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    out = tmp_path / "REMASEP_2026_07_DRAFT.xlsm"
    out.write_bytes(b"previous")
    result, _, _ = _generate(tmp_path, pending, _model(targets), output=out)
    assert result.status == w.STATUS_OUTPUT_EXISTS
    assert out.read_bytes() == b"previous"


# ---------------------------------------------------------------------------
# Template compatibility (§5)
# ---------------------------------------------------------------------------


def test_fingerprint_mismatch_blocks_generation(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    model = _model(targets, structural_fingerprint_id="stf:deadbeefdeadbeef")
    result, request, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_TEMPLATE_INCOMPATIBLE
    assert not request.output_path.exists()
    assert result.template_compatibility.actual_fingerprint_id == "stf:deadbeefdeadbeef"


# ---------------------------------------------------------------------------
# Defense-in-depth por celda (§9) + integridad (§18/§19)
# ---------------------------------------------------------------------------


def test_target_with_formula_aborts_and_keeps_template(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    model = _model(targets)
    model.formula_map[("REMASEP_OD", "B10")] = "=1+1"
    result, request, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_FAILED_INTEGRITY_CHECK
    assert w.TARGET_FORMULA_CONFLICT in result.errors[0]
    assert not request.output_path.exists()
    assert result.template_unchanged is True
    working = request.output_path.with_name("REMASEP_2026_07_DRAFT.__working__.xlsm")
    assert not working.exists()


def test_target_not_writable_aborts(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    model = _model(targets, writable_cells=set())
    result, _, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_GENERATION_FAILED
    assert "protegida" in result.errors[0]


def test_formula_expression_change_fails_integrity(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)

    def mutate(m: FakeWorkbookModel) -> None:
        m.formula_map[("CONTROL", "E16")] = "=E7+E8+999"

    model = _model(targets, mutate_after_save=mutate)
    result, request, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_FAILED_INTEGRITY_CHECK
    assert result.writer_integrity_status == w.WRITER_INTEGRITY_FAIL
    assert not request.output_path.exists()
    assert result.formula_integrity.ok is False


def test_vba_lost_fails_integrity(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)

    def mutate(m: FakeWorkbookModel) -> None:
        m.vba_present = False

    model = _model(targets, mutate_after_save=mutate)
    result, _, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_FAILED_INTEGRITY_CHECK
    assert result.vba_integrity.ok is False


def test_vba_present_before_and_after_on_success(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    result, _, _ = _generate(tmp_path, pending, _model(targets))
    assert result.vba_integrity.present_before is True
    assert result.vba_integrity.present_after is True


def test_writer_failure_cleans_working_copy_and_keeps_template(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    harness = FakeWriterHarness(_model(targets), fail_during_write=True)
    request = _request(tmp_path, pending)
    result = w.generate(request, pending, open_writer=harness.open_writer,
                        control_map=_CONTROL_MAP, inspect=harness.inspect)
    assert result.status == w.STATUS_GENERATION_FAILED
    assert not request.output_path.exists()
    working = request.output_path.with_name("REMASEP_2026_07_DRAFT.__working__.xlsm")
    assert not working.exists()
    assert result.template_unchanged is True


# ---------------------------------------------------------------------------
# CONTROL (§21/§22)
# ---------------------------------------------------------------------------


def test_control_pass_when_total_errors_zero(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    result, _, _ = _generate(tmp_path, pending, _model(targets))
    assert result.control_result.status == w.CONTROL_PASS
    assert result.control_status == w.CONTROL_PASS


def test_writer_integrity_pass_is_independent_of_control_errors(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    model = _model(targets)
    model.values[("CONTROL", "E16")] = 4  # p.ej. URGENCIAS/EyP_ET sin datos
    model.values[("CONTROL", "E8")] = 4
    result, _request, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_GENERATED_DIAGNOSTIC
    assert result.writer_integrity_status == w.WRITER_INTEGRITY_PASS
    assert result.control_status == w.CONTROL_FAIL
    assert any("CONTROL reporta errores" in x for x in result.warnings)


def test_parse_control_unavailable_without_sheet():
    snap = w.WorkbookSnapshot(
        path="x", sheet_names=("REMASEP_OD",), formula_map={}, values={},
        vba_present=True, vba_payload_sha256=None, structural_fingerprint_id=FP,
    )
    res = w.parse_control(snap, _CONTROL_MAP)
    assert res.status == w.CONTROL_UNAVAILABLE


def test_parse_control_counts_pending_module_errors():
    snap = w.WorkbookSnapshot(
        path="x", sheet_names=("CONTROL",), formula_map={},
        values={("CONTROL", "E16"): 2, ("CONTROL", "E8"): 2, ("CONTROL", "F8"): "SIN DATOS"},
        vba_present=True, vba_payload_sha256=None, structural_fingerprint_id=FP,
    )
    res = w.parse_control(snap, _CONTROL_MAP)
    assert res.status == w.CONTROL_FAIL
    assert res.errors_from_pending_modules == 2


# ---------------------------------------------------------------------------
# Naming + audit + privacidad (§16/§23/§32)
# ---------------------------------------------------------------------------


def test_plan_output_paths_deterministic(tmp_path):
    plan = w.plan_output_paths(
        tmp_path, year=2026, month=7, draft=True, working_suffix=".__working__",
        draft_name_template="REMASEP_{year}_{month:02d}_DRAFT.xlsm",
        final_name_template="REMASEP_{year}_{month:02d}.xlsm",
    )
    assert plan.final_path.name == "REMASEP_2026_07_DRAFT.xlsm"
    assert plan.working_path.name == "REMASEP_2026_07_DRAFT.__working__.xlsm"


def test_write_audit_has_no_pii(tmp_path):
    targets = [("REMASEP_OD", "B10", 5)]
    pending = [
        PendingWrite("wi:0", "LEGACY::REMASEP_OD::B10", "REMASEP_OD", "B10", 5,
                     "INTEGER_COUNT", "Julio 2026")
    ]
    result, _, _ = _generate(tmp_path, pending, _model(targets))
    audit = w.build_write_audit(pending, result.target_verification)
    blob = repr(audit)
    for pii in ("RUN", "nombre", "apellido", "@", "teléfono", "telefono", "dirección"):
        assert pii not in blob
    assert audit[0]["write_status"] == "OK"
    assert set(audit[0]) == {
        "instruction_id", "source_metric_id", "target_sheet", "target_cell",
        "written_value", "write_status",
    }


def test_invalid_mode_raises(tmp_path):
    pending = _pw([("REMASEP_OD", "B10", 1)])
    request = _request(tmp_path, pending)
    request = w.GenerationRequest(**{**request.__dict__, "mode": "NOPE"})
    with pytest.raises(w.ExcelWriterError):
        w.generate(request, pending, open_writer=FakeWriterHarness(_model([])).open_writer,
                   control_map=_CONTROL_MAP)
