"""Sprint 3.8 — carga y validación de los assets de runtime versionados."""

from __future__ import annotations

import hashlib
import shutil
from pathlib import Path

import pytest
import yaml

from remasep.services.runtime_assets import (
    RUNTIME_ASSET_INCOMPATIBLE,
    RUNTIME_ASSET_INVALID,
    RUNTIME_ASSET_MISSING,
    RuntimeAssetError,
    load_runtime_bundle,
    resolve_runtime_root,
)

_REAL_ROOT = Path(__file__).resolve().parent.parent / "config" / "runtime_2026"


# ---------------------------------------------------------------------------
# El bundle real versionado
# ---------------------------------------------------------------------------


def test_committed_bundle_loads_and_validates():
    bundle = load_runtime_bundle(_REAL_ROOT)
    assert bundle.version == "runtime_2026"
    assert bundle.template_fingerprint_id == "stf:dc624775927d4d4d"
    assert bundle.expected_instruction_count == 1122
    assert len(bundle.write_instructions) == 1122
    assert len(set(bundle.instruction_ids)) == 1122
    assert bundle.detail_sheet == "Atenciones - Detalles de citas"
    assert set(bundle.detail_columns_map) >= {"DIA_CITA", "SEXO", "PRESTACION", "ESTADO"}
    assert bundle.zero_write_policy == "WRITE_ZERO"
    # Sprint 3.9 fase 2: regla ESTADO confirmada funcionalmente
    assert bundle.estado_filter_status == "CONFIRMED"
    assert bundle.estado_filter.confirmed is True
    assert bundle.estado_filter.included_states == (
        "Atendido", "En Sala de Espera", "Atención Pausada", "En Atención",
    )
    assert bundle.estado_filter.excluded_states == (
        "Cancelado", "No Se Presenta", "Agendado", "Confirmado", "Re-Agendado",
    )
    # el catálogo cubre las 1122 + dependencias transitivas
    assert len(bundle.formula_index) >= 1122
    for w in bundle.write_instructions:
        assert w.source_metric_id in bundle.formula_index


def test_resolve_runtime_root_is_not_cwd_dependent(tmp_path, monkeypatch):
    monkeypatch.chdir(tmp_path)  # cwd irrelevante
    root = resolve_runtime_root()
    assert (root / "bundle.yaml").is_file()
    assert root.name == "runtime_2026"


def test_resolve_runtime_root_uses_meipass_bundle_when_frozen(tmp_path, monkeypatch):
    """Simula un build PyInstaller: sys._MEIPASS apunta al bundle onedir."""
    bundle_root = tmp_path / "_internal"
    target = bundle_root / "config" / "runtime_2026"
    target.mkdir(parents=True)
    (target / "bundle.yaml").write_text("version: fake\n", encoding="utf-8")

    monkeypatch.setattr("sys._MEIPASS", str(bundle_root), raising=False)
    monkeypatch.chdir(tmp_path)
    root = resolve_runtime_root()
    assert root == target


def test_bundle_exposes_producer_interfaces():
    bundle = load_runtime_bundle(_REAL_ROOT)
    resolvers = bundle.resolver_by_sheet()
    assert "REMASEP_OD" in resolvers
    # las 3 refs de criterio congeladas resuelven a su literal
    assert resolvers["REMASEP_OD"]("A8") == "EVALUACIÓN ODONTOLOGÍA GENERAL"
    assert resolvers["REMASEP_OD"]("$A$9") == "CONTROL ODONTOLOGÍA GENERAL"
    types = bundle.expected_value_types()
    assert set(types.values()) == {"INTEGER_COUNT"}


# ---------------------------------------------------------------------------
# Errores controlados (nunca FileNotFoundError crudo)
# ---------------------------------------------------------------------------


@pytest.fixture
def sandbox(tmp_path):
    dst = tmp_path / "runtime_2026"
    shutil.copytree(_REAL_ROOT, dst)
    return dst


def _rewrite_bundle(root: Path, **changes):
    doc = yaml.safe_load((root / "bundle.yaml").read_text())
    doc.update(changes)
    (root / "bundle.yaml").write_text(yaml.safe_dump(doc, allow_unicode=True))


def _rehash(root: Path, name: str) -> None:
    doc = yaml.safe_load((root / "bundle.yaml").read_text())
    doc["assets"][name] = hashlib.sha256((root / name).read_bytes()).hexdigest()
    (root / "bundle.yaml").write_text(yaml.safe_dump(doc, allow_unicode=True))


def test_missing_root_raises_controlled_error(tmp_path):
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(tmp_path / "nope")
    assert exc.value.code == RUNTIME_ASSET_MISSING
    assert "instalación no contiene" in exc.value.user_message


def test_missing_asset_file_raises_controlled_error(sandbox):
    (sandbox / "metric_catalog.csv").unlink()
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_MISSING


def test_corrupt_asset_sha_mismatch_is_invalid(sandbox):
    path = sandbox / "detail_contract.yaml"
    path.write_text(path.read_text() + "\n# tampered\n")
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID


def test_manifest_schema_change_is_invalid(sandbox):
    lines = (sandbox / "write_manifest.csv").read_text().splitlines()
    lines[0] = lines[0].replace("instruction_id", "iid")
    (sandbox / "write_manifest.csv").write_text("\n".join(lines) + "\n")
    _rehash(sandbox, "write_manifest.csv")
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID


def test_duplicate_instruction_id_is_invalid(sandbox):
    lines = (sandbox / "write_manifest.csv").read_text().splitlines()
    lines.append(lines[1])  # duplica la primera fila
    (sandbox / "write_manifest.csv").write_text("\n".join(lines) + "\n")
    _rehash(sandbox, "write_manifest.csv")
    _rewrite_bundle(sandbox, expected_instruction_count=1123)
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID
    assert "duplicad" in str(exc.value)


def test_count_mismatch_is_invalid(sandbox):
    _rewrite_bundle(sandbox, expected_instruction_count=999)
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID


def test_fingerprint_mismatch_is_incompatible(sandbox):
    _rewrite_bundle(sandbox, template_fingerprint_id="stf:deadbeefdeadbeef")
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INCOMPATIBLE


def test_catalog_gap_is_invalid(sandbox):
    lines = (sandbox / "metric_catalog.csv").read_text().splitlines()
    (sandbox / "metric_catalog.csv").write_text(lines[0] + "\n" + lines[1] + "\n")  # sólo 1 métrica
    _rehash(sandbox, "metric_catalog.csv")
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID


def test_zero_policy_incompatible(sandbox):
    (sandbox / "zero_policy.yaml").write_text(
        yaml.safe_dump({"version": "runtime_2026", "resolution": "UNRESOLVED"})
    )
    _rehash(sandbox, "zero_policy.yaml")
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INCOMPATIBLE


def _rewrite_estado(sandbox, doc):
    (sandbox / "estado_filter.yaml").write_text(yaml.safe_dump(doc, allow_unicode=True))
    _rehash(sandbox, "estado_filter.yaml")


def test_estado_filter_unknown_status_is_invalid(sandbox):
    _rewrite_estado(sandbox, {
        "version": "runtime_2026", "status": "MAYBE",
        "included_states": ["Atendido"], "excluded_states": [],
    })
    _rewrite_bundle(sandbox, estado_filter_status="MAYBE")
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID
    assert "estado_filter.status" in str(exc.value)


def test_estado_filter_confirmed_without_states_is_invalid(sandbox):
    _rewrite_estado(sandbox, {
        "version": "runtime_2026", "status": "CONFIRMED",
        "included_states": [], "excluded_states": [],
    })
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID
    assert "included_states" in str(exc.value)


def test_estado_filter_overlapping_states_is_invalid(sandbox):
    _rewrite_estado(sandbox, {
        "version": "runtime_2026", "status": "CONFIRMED",
        "included_states": ["Atendido", "Cancelado"],
        "excluded_states": ["Cancelado"],
    })
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID
    assert "included y excluded" in str(exc.value)


def test_estado_filter_top_level_mismatch_is_invalid(sandbox):
    _rewrite_bundle(sandbox, estado_filter_status="PENDING_FUNCTIONAL_CONFIRMATION")
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_INVALID
    assert "estado_filter_status" in str(exc.value)


def test_estado_filter_missing_asset_is_missing_error(sandbox):
    (sandbox / "estado_filter.yaml").unlink()
    with pytest.raises(RuntimeAssetError) as exc:
        load_runtime_bundle(sandbox)
    assert exc.value.code == RUNTIME_ASSET_MISSING
