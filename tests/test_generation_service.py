"""Sprint 3.7B — GenerationService (sin Excel real; writer inyectado)."""

from __future__ import annotations

import json

from remasep.services import excel_writer as w
from remasep.services.generation_service import GenerationService
from remasep.services.metric_value_producer import PendingWrite
from remasep.testing import FakeWorkbookModel, FakeWriterHarness

FP = "stf:dc624775927d4d4d"
SHEETS = ["REMASEP_OD", "REMASEP 01", "B2 ANEXO", "CONTROL"]


def _pending():
    return [
        PendingWrite("wi:0", "LEGACY::REMASEP_OD::B10", "REMASEP_OD", "B10", 4,
                     "INTEGER_COUNT", "Julio 2026"),
        PendingWrite("wi:1", "LEGACY::REMASEP_OD::B11", "REMASEP_OD", "B11", 0,
                     "INTEGER_COUNT", "Julio 2026"),
    ]


def _model(pending):
    return FakeWorkbookModel(
        sheet_names=list(SHEETS),
        formula_map={("CONTROL", "E16"): "=E6"},
        values={("CONTROL", "E16"): 0},
        writable_cells={(p.target_sheet, p.target_cell) for p in pending},
        structural_fingerprint_id=FP,
    )


def test_generate_with_injected_writer_produces_artifacts(tmp_path):
    pending = _pending()
    harness = FakeWriterHarness(_model(pending))
    template = tmp_path / "tpl.xlsm"
    template.write_bytes(b"PK\x03\x04 placeholder")
    out = tmp_path / "out" / "REMASEP_2026_07_DRAFT.xlsm"

    service = GenerationService(artifacts_dir=tmp_path / "artifacts")
    sr = service.generate(
        mode=w.MODE_DIAGNOSTIC_REFERENCE,
        template_path=template,
        output_path=out,
        pending_writes=pending,
        manifest_instruction_ids=["wi:0", "wi:1"],
        period_year=2026,
        period_month=7,
        zero_write_policy="WRITE_ZERO",
        expected_fingerprint_id=FP,
        open_writer=harness.open_writer,
        inspect=harness.inspect,
        probe_com=False,
    )

    assert sr.status == w.STATUS_GENERATED_DIAGNOSTIC
    assert sr.submission_label == w.SUBMISSION_LABEL
    assert sr.written_cells == 2
    assert out.exists()
    for name in (
        "generation_summary.json", "write_audit.csv", "target_verification.csv",
        "formula_integrity.csv", "control_result.csv", "template_verification.json",
        "README.md",
    ):
        assert (tmp_path / "artifacts" / name).is_file(), name

    summary = json.loads((tmp_path / "artifacts" / "generation_summary.json").read_text())
    assert summary["status"] == w.STATUS_GENERATED_DIAGNOSTIC
    assert summary["submission_label"] == w.SUBMISSION_LABEL
    assert summary["writer_integrity_status"] == w.WRITER_INTEGRITY_PASS
    assert summary["capability"]["platform"]  # presente


def test_generate_without_excel_returns_unavailable(tmp_path, monkeypatch):
    import remasep.services.generation_service as gs

    monkeypatch.setattr(
        gs, "detect_excel_capability",
        lambda **_: gs.ExcelCapability(
            platform="Linux", pywin32_available=False, excel_com_available=False,
            excel_version=None, can_generate=False, reason="WINDOWS_EXCEL_REQUIRED",
        ),
    )
    template = tmp_path / "tpl.xlsm"
    template.write_bytes(b"x")
    service = GenerationService(artifacts_dir=tmp_path / "artifacts")
    sr = service.generate(
        mode=w.MODE_PRODUCTION,
        template_path=template,
        pending_writes=_pending(),
        manifest_instruction_ids=["wi:0", "wi:1"],
        period_year=2026,
        period_month=7,
        zero_write_policy="WRITE_ZERO",
    )
    assert sr.status == w.STATUS_EXCEL_UNAVAILABLE
    assert sr.result.errors == ["WINDOWS_EXCEL_REQUIRED"]
    assert not (tmp_path / "artifacts").exists()


def test_service_reads_versioned_config(tmp_path):
    """La política y el mapa de CONTROL se leen de config/excel_writer_2026/."""
    service = GenerationService()
    policy = service._policy()
    assert policy["version"] == "excel_writer_2026"
    assert policy["final_ready_allowed"] is False
    control_map = service._control_map(policy)
    assert control_map.sheet == "CONTROL"
    assert control_map.total_errors_cell == "E16"
    assert any(r.sheet == "REMASEP_OD" for r in control_map.rows)
