"""Regresión F03: el layout empaquetado aislado contiene todo el cálculo."""

from __future__ import annotations

import importlib.util
import os
import shutil
import subprocess
import sys
from datetime import date
from pathlib import Path

import build_smoke_test
from openpyxl import Workbook

ROOT = Path(__file__).resolve().parent.parent


def _packaging_datas():
    module_path = ROOT / "packaging" / "_spec_common.py"
    spec = importlib.util.spec_from_file_location("remasep_spec_common", module_path)
    assert spec is not None and spec.loader is not None
    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)
    return module.common_paths(str(ROOT / "packaging"))["datas"]


def _make_medinet(path: Path) -> None:
    wb = Workbook()
    ws = wb.active
    ws.append(
        [
            "DIA CITA",
            "FECHA NACIMIENTO",
            "SEXO",
            "SUCURSAL",
            "ESPECIALIDAD",
            "TIPO DE CITA",
            "PRESTACIÓN",
            "ESTADO",
            "MODALIDAD",
            "PRESTACIÓN REALIZADA",
        ]
    )
    ws.append(
        [
            date(2026, 7, 10),
            date(1990, 1, 1),
            "Mujer",
            "Centro",
            "ODONTOLOGIA",
            "CONSULTA GENERAL",
            "EVALUACIÓN ODONTOLÓGICA",
            "Atendido",
            "Presencial",
            "",
        ]
    )
    wb.save(path)


def test_isolated_distribution_resolves_all_assets_and_calculates_1122(tmp_path):
    dist = tmp_path / "dist" / "REMASEP"
    internal = dist / "_internal"
    shutil.copytree(ROOT / "src" / "remasep", internal / "remasep")
    for source, destination in _packaging_datas():
        shutil.copytree(source, internal / destination)
    (dist / "REMASEP.exe").write_bytes(b"MZ-test")

    smoke = build_smoke_test.check_dist(dist)
    assert smoke.ok, [check for check in smoke.checks if not check[1]]

    medinet = tmp_path / "fixture.xlsx"
    _make_medinet(medinet)
    outside_cwd = tmp_path / "outside"
    outside_cwd.mkdir()
    result_file = tmp_path / "result.txt"
    script = tmp_path / "isolated_check.py"
    script.write_text(
        "\n".join(
            [
                "import sys",
                "from pathlib import Path",
                f"internal = Path({str(internal)!r})",
                "sys.path.insert(0, str(internal))",
                "sys._MEIPASS = str(internal)",
                "from remasep.services.common import Period",
                "from remasep.services.generation_service import GenerationService",
                "from remasep.services.legacy_rules import load_legacy_rules",
                "from remasep.services.production_pipeline import build_production_pending_writes",
                "from remasep.services.runtime_assets import load_runtime_bundle",
                "from remasep.main import _smoke_check",
                "assert _smoke_check() == 0",
                "bundle = load_runtime_bundle()",
                "rules = load_legacy_rules()",
                "service = GenerationService()",
                "policy = service._policy()",
                "control_map = service._control_map(policy)",
                f"result = build_production_pending_writes(Path({str(medinet)!r}), Period(7, 2026), bundle=bundle)",
                f"Path({str(result_file)!r}).write_text(f'{{len(result.pending_writes)}}|{{result.completeness.ok}}|{{len(rules.categories)}}|{{policy[\"version\"]}}|{{control_map.sheet}}')",
            ]
        ),
        encoding="utf-8",
    )

    env = os.environ.copy()
    env.pop("PYTHONPATH", None)
    completed = subprocess.run(
        [sys.executable, "-I", str(script)],
        cwd=outside_cwd,
        env=env,
        capture_output=True,
        text=True,
        check=False,
    )
    assert completed.returncode == 0, completed.stderr
    assert result_file.read_text(encoding="utf-8") == (
        "1122|True|6|excel_writer_2026|CONTROL"
    )


def test_smoke_rejects_distribution_without_productive_rules(tmp_path):
    dist = tmp_path / "REMASEP"
    (dist / "_internal" / "config" / "runtime_2026").mkdir(parents=True)
    (dist / "_internal" / "config" / "excel_writer_2026").mkdir(parents=True)
    (dist / "REMASEP.exe").write_bytes(b"MZ-test")
    for rel in build_smoke_test.REQUIRED_RELATIVE_PATHS:
        if rel == "REMASEP.exe" or "legacy_current_logic_2026" in rel:
            continue
        target = dist / rel
        target.parent.mkdir(parents=True, exist_ok=True)
        target.write_text("test", encoding="utf-8")

    result = build_smoke_test.check_dist(dist)
    assert result.ok is False
    assert any(
        "legacy_current_logic_2026/rules.yaml" in name and not ok
        for name, ok, _detail in result.checks
    )
