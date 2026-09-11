"""Sprint 3.11 — lógica del smoke test estructural sobre un dist/ construido.

No requiere Windows ni un build real de PyInstaller: arma un árbol sintético
con la misma forma (``dist/REMASEP/REMASEP.exe`` + ``_internal/config/...``)
para probar `check_dist` en ambos sentidos (build correcto / build con fugas
de desarrollo).
"""

from __future__ import annotations

import build_smoke_test as bst


def _make_good_dist(tmp_path):
    dist = tmp_path / "REMASEP"
    internal = dist / "_internal"
    dist.mkdir(parents=True)
    (dist / "REMASEP.exe").write_bytes(b"MZ-fake-exe")

    runtime = internal / "config" / "runtime_2026"
    runtime.mkdir(parents=True)
    (runtime / "bundle.yaml").write_text("version: fake\n", encoding="utf-8")
    (runtime / "write_manifest.csv").write_text("instruction_id\n", encoding="utf-8")
    (runtime / "metric_catalog.csv").write_text("metric_id\n", encoding="utf-8")

    writer = internal / "config" / "excel_writer_2026"
    writer.mkdir(parents=True)
    (writer / "policy.yaml").write_text("version: fake\n", encoding="utf-8")
    (writer / "control_map.yaml").write_text("sheet: CONTROL\n", encoding="utf-8")

    # algo de "peso" normal de un build real (dlls, pyd, etc.)
    (internal / "PySide6" / "QtCore.pyd").parent.mkdir(parents=True, exist_ok=True)
    (internal / "PySide6" / "QtCore.pyd").write_bytes(b"\x00")
    return dist


def test_check_dist_passes_on_a_correct_build(tmp_path):
    dist = _make_good_dist(tmp_path)
    result = bst.check_dist(dist)
    assert result.ok, [c for c in result.checks if not c[1]]


def test_check_dist_fails_when_exe_missing(tmp_path):
    dist = _make_good_dist(tmp_path)
    (dist / "REMASEP.exe").unlink()
    result = bst.check_dist(dist)
    assert not result.ok
    names_failed = {name for name, ok, _ in result.checks if not ok}
    assert "presente: REMASEP.exe" in names_failed


def test_check_dist_fails_when_runtime_config_missing(tmp_path):
    dist = _make_good_dist(tmp_path)
    (dist / "_internal" / "config" / "runtime_2026" / "bundle.yaml").unlink()
    result = bst.check_dist(dist)
    assert not result.ok


def test_check_dist_fails_when_excel_writer_config_missing(tmp_path):
    dist = _make_good_dist(tmp_path)
    (dist / "_internal" / "config" / "excel_writer_2026" / "policy.yaml").unlink()
    result = bst.check_dist(dist)
    assert not result.ok


def test_check_dist_fails_when_dev_data_leaked_into_the_build(tmp_path):
    dist = _make_good_dist(tmp_path)
    leak = dist / "_internal" / "data" / "local"
    leak.mkdir(parents=True)
    (leak / "detalle_citas - 2026-09-07T123630.940.xlsx").write_bytes(b"fake")
    result = bst.check_dist(dist)
    assert not result.ok


def test_check_dist_fails_when_generator_legacy_workbook_leaked(tmp_path):
    dist = _make_good_dist(tmp_path)
    (dist / "_internal" / "GENERACION DATOS REMASEP.xlsx").write_bytes(b"fake")
    result = bst.check_dist(dist)
    assert not result.ok


def test_check_dist_fails_when_dev_artifacts_dir_present(tmp_path):
    dist = _make_good_dist(tmp_path)
    (dist / "_internal" / "artifacts" / "excel_writer").mkdir(parents=True)
    result = bst.check_dist(dist)
    assert not result.ok


def test_check_dist_fails_when_dist_directory_does_not_exist(tmp_path):
    result = bst.check_dist(tmp_path / "no_existe")
    assert not result.ok
    assert result.checks[0][0] == "el directorio dist existe"


def test_main_returns_nonzero_exit_code_on_failure(tmp_path, capsys):
    code = bst.main(["--dist", str(tmp_path / "no_existe")])
    assert code == 1
    out = capsys.readouterr().out
    assert "RESULTADO: FALLA" in out


def test_main_returns_zero_on_a_correct_build(tmp_path, capsys):
    dist = _make_good_dist(tmp_path)
    code = bst.main(["--dist", str(dist)])
    assert code == 0
    out = capsys.readouterr().out
    assert "RESULTADO: OK" in out
