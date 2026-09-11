"""Sprint 3.11 — resolución de paths de la app empaquetada.

Nada aquí depende del *cwd* del proceso ni asume una estructura de repo en
tiempo de ejecución: se simula ``sys._MEIPASS`` (PyInstaller onedir) para
probar el camino "empaquetado" sin tener que construir un ``.exe`` real.
"""

from __future__ import annotations

from pathlib import Path

import pytest

from remasep import app_paths
from remasep.services.runtime_assets import resolve_runtime_root

_REPO_ROOT = Path(__file__).resolve().parent.parent


# ---------------------------------------------------------------------------
# resource_root() / find_config_dir() — INTERNAL_READ_ONLY
# ---------------------------------------------------------------------------


def test_is_frozen_is_false_in_dev(monkeypatch):
    monkeypatch.delattr("sys.frozen", raising=False)
    assert app_paths.is_frozen() is False


def test_resource_root_is_not_cwd_dependent(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)  # carpeta vacía, sin config/
    root = app_paths.resource_root()
    assert root == _REPO_ROOT
    assert (root / "config" / "runtime_2026").is_dir()
    assert (root / "config" / "excel_writer_2026").is_dir()
    assert root != tmp_path


def test_resource_root_ignores_unrelated_config_named_directories(tmp_path, monkeypatch):
    """Regresión: src/remasep/config/ (subpaquete Python no relacionado, con
    sólo loader.py) no debe confundirse con el config/ de datos de la raíz
    del repo — resource_root() debe seguir subiendo hasta encontrar uno que
    tenga los dos árboles reales (runtime_2026 / excel_writer_2026)."""
    decoy_repo = tmp_path / "fake_repo"
    decoy_pkg_config = decoy_repo / "src" / "remasep" / "config"
    decoy_pkg_config.mkdir(parents=True)
    (decoy_pkg_config / "loader.py").write_text("# no es config de datos\n", encoding="utf-8")

    real_config = decoy_repo / "config"
    (real_config / "runtime_2026").mkdir(parents=True)
    (real_config / "excel_writer_2026").mkdir(parents=True)

    fake_module_file = decoy_repo / "src" / "remasep" / "app_paths.py"
    monkeypatch.setattr(app_paths, "__file__", str(fake_module_file))
    assert app_paths.resource_root() == decoy_repo


def test_resource_root_uses_meipass_when_frozen(tmp_path, monkeypatch):
    fake_bundle = tmp_path / "_internal"
    fake_bundle.mkdir()
    monkeypatch.setattr("sys._MEIPASS", str(fake_bundle), raising=False)
    assert app_paths.resource_root() == fake_bundle


def test_find_config_dir_finds_real_runtime_2026_from_repo(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)  # cwd distinto: no debe importar
    found = app_paths.find_config_dir("runtime_2026", required_file="bundle.yaml")
    assert (found / "bundle.yaml").is_file()
    assert found.name == "runtime_2026"


def test_find_config_dir_prefers_meipass_bundle_with_spaces_and_unicode(tmp_path, monkeypatch):
    # simula una instalación en una ruta realista de Windows: espacios + tildes/ñ
    bundle_root = tmp_path / "Carpeta con espacios y ñ" / "_internal"
    target = bundle_root / "config" / "excel_writer_2026"
    target.mkdir(parents=True)
    (target / "policy.yaml").write_text("version: fake\n", encoding="utf-8")

    monkeypatch.setattr("sys._MEIPASS", str(bundle_root), raising=False)
    found = app_paths.find_config_dir("excel_writer_2026", required_file="policy.yaml")
    assert found == target
    assert "espacios" in str(found)


def test_find_config_dir_missing_returns_a_path_without_crashing(tmp_path, monkeypatch):
    monkeypatch.setattr("sys._MEIPASS", str(tmp_path), raising=False)
    found = app_paths.find_config_dir("no_existe_2099", required_file="nada.yaml")
    assert isinstance(found, Path)
    assert not found.is_dir()


def test_runtime_config_dir_matches_resolve_runtime_root():
    assert app_paths.runtime_config_dir() == resolve_runtime_root()


def test_excel_writer_config_dir_finds_the_real_policy(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)
    found = app_paths.excel_writer_config_dir()
    assert (found / "policy.yaml").is_file()
    assert (found / "control_map.yaml").is_file()
    assert found == _REPO_ROOT / "config" / "excel_writer_2026"


# ---------------------------------------------------------------------------
# user_data_dir() / logs_dir() / diagnostics_dir() — APP_WRITABLE_DATA
# ---------------------------------------------------------------------------


def test_user_data_dir_is_created_and_writable(tmp_path, monkeypatch):
    target = tmp_path / "appdata"
    monkeypatch.setenv("REMASEP_APP_DATA_DIR", str(target))
    root = app_paths.user_data_dir()
    assert root == target
    assert root.is_dir()
    probe = root / "probe.txt"
    probe.write_text("ok", encoding="utf-8")
    assert probe.read_text(encoding="utf-8") == "ok"


def test_user_data_dir_supports_spaces_and_unicode(tmp_path, monkeypatch):
    target = tmp_path / "Usuarios" / "María José Ñúñez" / "REMASEP data"
    monkeypatch.setenv("REMASEP_APP_DATA_DIR", str(target))
    assert app_paths.user_data_dir() == target
    assert app_paths.logs_dir() == target / "logs"
    assert app_paths.diagnostics_dir() == target / "diagnostics"
    assert (target / "logs").is_dir()
    assert (target / "diagnostics").is_dir()


def test_logs_and_diagnostics_are_subdirs_of_user_data_dir(tmp_path, monkeypatch):
    monkeypatch.setenv("REMASEP_APP_DATA_DIR", str(tmp_path / "appdata"))
    root = app_paths.user_data_dir()
    assert app_paths.logs_dir().parent == root
    assert app_paths.diagnostics_dir().parent == root


def test_writable_data_never_lands_inside_the_frozen_bundle(tmp_path, monkeypatch):
    bundle_root = tmp_path / "_internal"
    bundle_root.mkdir()
    monkeypatch.setattr("sys._MEIPASS", str(bundle_root), raising=False)
    app_data = tmp_path / "appdata"
    monkeypatch.setenv("REMASEP_APP_DATA_DIR", str(app_data))

    resource = app_paths.resource_root()
    data = app_paths.user_data_dir()
    assert resource == bundle_root
    assert not str(data).startswith(str(resource))


@pytest.mark.parametrize("dirname", ["logs", "diagnostics"])
def test_writable_dirs_are_not_read_only_resource_paths(tmp_path, monkeypatch, dirname):
    monkeypatch.setenv("REMASEP_APP_DATA_DIR", str(tmp_path / "appdata"))
    getter = getattr(app_paths, f"{dirname}_dir")
    d = getter()
    assert d.is_dir()
    (d / "write_probe.txt").write_text("x", encoding="utf-8")  # no lanza: es escribible
