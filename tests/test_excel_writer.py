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
    assert result.run_id  # workspace único por corrida
    assert result.cleanup_status == w.CLEANUP_OK
    assert result.workspace_path is None
    assert not (tmp_path / ".remasep-tmp").exists()  # workspace borrado
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
    assert result.cleanup_status == w.CLEANUP_OK
    assert not (tmp_path / ".remasep-tmp").exists()


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
    assert any("VBA_LOST" in e for e in result.errors)


def test_vba_code_change_fails_integrity(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)

    def mutate(m: FakeWorkbookModel) -> None:
        m.vba_modules = {"Módulo1": 'Sub PROTEGER()\n  Shell "x"\nEnd Sub', "ThisWorkbook": ""}

    model = _model(targets, mutate_after_save=mutate)
    result, _, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_FAILED_INTEGRITY_CHECK
    assert any("VBA_MODULE_SOURCE_CHANGED" in e for e in result.errors)


def test_vba_binary_payload_change_only_is_pass_with_warning(tmp_path):
    """El caso real reportado: Save de Excel reescribe vbaProject.bin sin tocar
    el código -> WRITER_INTEGRITY_PASS + warning, no FAIL."""
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)

    def mutate(m: FakeWorkbookModel) -> None:
        m.vba_payload_sha256 = "post-save-binary-sha"  # sólo el binario

    model = _model(targets, mutate_after_save=mutate)
    result, _, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_GENERATED_DIAGNOSTIC
    assert result.writer_integrity_status == w.WRITER_INTEGRITY_PASS
    assert result.vba_integrity.payload_stable is False
    assert any("SEMANTIC_EQUIVALENT" in x for x in result.warnings)


def test_vba_semantic_unavailable_is_fail_closed_and_blocks_output(tmp_path):
    """Si el VBA existe pero no se puede verificar el código -> FAIL, no se
    promueve la salida, se limpia el workspace, plantilla intacta."""
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    model = _model(targets, vba_semantic_available=False)
    result, request, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_FAILED_INTEGRITY_CHECK
    assert result.writer_integrity_status == w.WRITER_INTEGRITY_FAIL
    assert result.vba_integrity.status == "VBA_INTEGRITY_FAIL"
    assert result.vba_integrity.ok is False
    assert any("VBA_SEMANTIC_CHECK_UNAVAILABLE" in e for e in result.errors)
    assert not request.output_path.exists()
    assert result.cleanup_status == w.CLEANUP_OK
    assert result.template_unchanged is True


def test_vba_present_before_and_after_on_success(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    result, _, _ = _generate(tmp_path, pending, _model(targets))
    assert result.vba_integrity.present_before is True
    assert result.vba_integrity.present_after is True


def test_writer_failure_cleans_workspace_and_keeps_template(tmp_path):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    harness = FakeWriterHarness(_model(targets), fail_during_write=True)
    request = _request(tmp_path, pending)
    result = w.generate(request, pending, open_writer=harness.open_writer,
                        control_map=_CONTROL_MAP, inspect=harness.inspect)
    assert result.status == w.STATUS_GENERATION_FAILED
    assert not request.output_path.exists()
    assert result.cleanup_status == w.CLEANUP_OK
    assert not (tmp_path / ".remasep-tmp").exists()
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


# ---------------------------------------------------------------------------
# Workspace temporal por corrida (§11 — patch final)
# ---------------------------------------------------------------------------


def test_two_runs_do_not_share_a_temp_path(tmp_path):
    out = tmp_path / "outputs"
    ws1 = w.create_run_workspace(out / "R1.xlsm")
    ws2 = w.create_run_workspace(out / "R2.xlsm")
    assert ws1.run_id != ws2.run_id
    assert ws1.working_path != ws2.working_path
    assert ws1.root.is_dir() and ws2.root.is_dir()
    assert w._is_within(ws1.root, ws1.tmp_base)


def test_create_run_workspace_never_reuses_an_existing_dir(tmp_path):
    out = tmp_path / "outputs"
    ws = w.create_run_workspace(out / "R.xlsm", run_id="fixed")
    assert ws.root.is_dir()
    with pytest.raises(FileExistsError):  # exist_ok=False
        w.create_run_workspace(out / "R.xlsm", run_id="fixed")


def test_stale_temp_from_another_run_does_not_block_generation(tmp_path):
    targets = [("REMASEP_OD", "B10", 3)]
    pending = _pw(targets)
    out = tmp_path / "REMASEP_2026_07_DRAFT.xlsm"
    # simula un temporal huérfano de una corrida previa que crasheó, con un
    # archivo "bloqueado" adentro
    stale = tmp_path / ".remasep-tmp" / "20200101T000000-deadbeefcafe"
    stale.mkdir(parents=True)
    (stale / "working.xlsm").write_bytes(b"stale, locked")
    result, request, _ = _generate(tmp_path, pending, _model(targets), output=out)
    assert result.status == w.STATUS_GENERATED_DIAGNOSTIC
    assert request.output_path.exists()
    assert stale.exists()  # NO se toca el temporal ajeno


def test_cleanup_pending_when_workspace_locked_does_not_hide_original_error(
    tmp_path, monkeypatch
):
    targets = [("REMASEP_OD", "B10", 1)]
    pending = _pw(targets)
    model = _model(targets)
    model.formula_map[("REMASEP_OD", "B10")] = "=1+1"  # -> TARGET_FORMULA_CONFLICT

    def _boom(_path):
        raise PermissionError("WinError 32: archivo en uso")

    monkeypatch.setattr(w.shutil, "rmtree", _boom)
    result, _, _ = _generate(tmp_path, pending, model)
    assert result.status == w.STATUS_FAILED_INTEGRITY_CHECK
    assert w.TARGET_FORMULA_CONFLICT in result.errors[0]  # error original intacto
    assert result.cleanup_status == w.CLEANUP_PENDING
    assert result.workspace_path is not None
    assert any("CLEANUP_PENDING" in x for x in result.warnings)


def test_cleanup_refuses_to_delete_outside_its_workspace(tmp_path):
    outside = tmp_path / "not-a-workspace"
    outside.mkdir()
    bogus = w.RunWorkspace(
        run_id="x", tmp_base=tmp_path / ".remasep-tmp", root=outside,
        working_path=outside / "working.xlsm",
    )
    status, _diag = w.cleanup_run_workspace(bogus)
    assert status == w.CLEANUP_REFUSED
    assert outside.exists()  # no se borró nada ajeno


def test_atomic_promote_retries_then_raises(monkeypatch):
    calls = {"n": 0}

    class _P:
        def replace(self, _dst):
            calls["n"] += 1
            raise PermissionError("bloqueado")
        name = "x.xlsm"

    monkeypatch.setattr(w, "_sleep", lambda _s: None)
    with pytest.raises(w.ExcelWriterError):
        w._atomic_promote(_P(), Path("x.xlsm"))
    assert calls["n"] == w._PROMOTE_ATTEMPTS  # acotado, sin loop infinito


def test_atomic_promote_succeeds_after_transient_lock(tmp_path, monkeypatch):
    src = tmp_path / "src.xlsm"
    src.write_bytes(b"data")
    dst = tmp_path / "dst.xlsm"
    real_replace = Path.replace
    state = {"fails": 2}

    def flaky(self, target):
        if state["fails"] > 0:
            state["fails"] -= 1
            raise PermissionError("transient")
        return real_replace(self, target)

    monkeypatch.setattr(w, "_sleep", lambda _s: None)
    monkeypatch.setattr(Path, "replace", flaky)
    w._atomic_promote(src, dst)
    assert dst.read_bytes() == b"data"


def test_no_process_killing_in_codebase():
    """El writer nunca mata procesos ni fuerza borrados globales: no hay
    ``subprocess`` / ``os.system`` / ``os.kill`` / ``psutil`` ni comandos de
    kill como literales de string. (Las MENCIONES en docstrings de que NO se
    hace esto están permitidas y se ignoran.)"""
    import ast

    root = Path(__file__).resolve().parent.parent
    banned_strings = ("taskkill", "stop-process", "pkill", "killall", "/f /im")
    banned_calls = {("os", "system"), ("os", "kill"), ("os", "popen")}
    hits: list[str] = []
    for py in list((root / "src").rglob("*.py")) + list((root / "scripts").rglob("*.py")):
        tree = ast.parse(py.read_text(encoding="utf-8"), filename=str(py))
        # ids de constantes que son docstrings / strings sueltos (inertes)
        inert = {
            id(n.value)
            for n in ast.walk(tree)
            if isinstance(n, ast.Expr) and isinstance(n.value, ast.Constant)
        }
        for node in ast.walk(tree):
            if isinstance(node, ast.Import):
                for alias in node.names:
                    if alias.name.split(".")[0] in {"subprocess", "psutil"}:
                        hits.append(f"{py.name}: import {alias.name}")
            elif isinstance(node, ast.ImportFrom) and (node.module or "").split(".")[0] in {
                "subprocess",
                "psutil",
            }:
                hits.append(f"{py.name}: from {node.module}")
            elif (
                isinstance(node, ast.Constant)
                and isinstance(node.value, str)
                and id(node) not in inert
            ):
                low = node.value.lower()
                hits += [f"{py.name}: string {b!r}" for b in banned_strings if b in low]
            elif isinstance(node, ast.Call) and isinstance(node.func, ast.Attribute):
                value = node.func.value
                if isinstance(value, ast.Name) and (value.id, node.func.attr) in banned_calls:
                    hits.append(f"{py.name}: {value.id}.{node.func.attr}()")
    assert not hits, hits
