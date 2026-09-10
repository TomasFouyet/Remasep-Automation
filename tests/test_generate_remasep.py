"""Sprint 3.7B / 3.8 — CLI generate_remasep + desacople de runtime."""

from __future__ import annotations

import ast
from pathlib import Path

import generate_remasep as cli

_SRC = Path(__file__).resolve().parent.parent


def _module_level_imports(path: Path) -> set[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    names: set[str] = set()
    for node in tree.body:  # sólo top-level
        if isinstance(node, ast.Import):
            names |= {a.name for a in node.names}
        elif isinstance(node, ast.ImportFrom):
            names.add(node.module or "")
    return names


def _non_docstring_strings(path: Path) -> list[str]:
    tree = ast.parse(path.read_text(encoding="utf-8"), filename=str(path))
    inert = {
        id(n.value)
        for n in ast.walk(tree)
        if isinstance(n, ast.Expr) and isinstance(n.value, ast.Constant)
    }
    return [
        n.value
        for n in ast.walk(tree)
        if isinstance(n, ast.Constant) and isinstance(n.value, str) and id(n) not in inert
    ]


def test_module_imports_without_win32com() -> None:
    import remasep.adapters.excel_com  # noqa: F401

    assert hasattr(cli, "main")


def test_production_modules_do_not_import_legacy_at_module_level() -> None:
    """El path de producción no importa el motor legacy ni lee artifacts/."""
    for rel in (
        "src/remasep/services/production_pipeline.py",
        "src/remasep/services/runtime_assets.py",
    ):
        imports = _module_level_imports(_SRC / rel)
        for token in ("build_metric_values", "close_legacy_aggregations",
                      "compare_legacy_aggregations"):
            assert not any(token in imp for imp in imports), (rel, token)

    # en el CLI, el import legacy sólo puede vivir DENTRO de _diagnostic_inputs
    cli_imports = _module_level_imports(_SRC / "scripts/generate_remasep.py")
    assert not any("build_metric_values" in imp for imp in cli_imports)
    src = (_SRC / "scripts/generate_remasep.py").read_text(encoding="utf-8")
    diag = src.split("def _diagnostic_inputs")[1].split("\ndef ")[0]
    assert "from build_metric_values import" in diag  # perezoso, dentro de la función


def test_production_modules_have_no_legacy_or_artifacts_string_literals() -> None:
    """Ninguna ruta a artifacts/ ni al workbook legacy como literal ejecutable
    (las menciones en docstrings están permitidas)."""
    for rel in ("production_pipeline.py", "runtime_assets.py"):
        strings = _non_docstring_strings(_SRC / "src/remasep/services" / rel)
        for s in strings:
            assert "artifacts/" not in s, (rel, s)
            assert "GENERACION DATOS REMASEP" not in s, (rel, s)


def test_default_mode_is_production() -> None:
    args = cli.build_parser().parse_args([])
    assert args.mode == "production"
    assert cli._MODE_ALIASES["production"] == cli.MODE_PRODUCTION
    assert cli._MODE_ALIASES["diagnostic"] == cli.MODE_DIAGNOSTIC_REFERENCE


def test_cli_is_graceful_when_excel_unavailable(monkeypatch, capsys):
    import remasep.services.generation_service as gs

    monkeypatch.setattr(
        gs, "detect_excel_capability",
        lambda **_: gs.ExcelCapability(
            platform="Linux", pywin32_available=False, excel_com_available=False,
            excel_version=None, can_generate=False, reason="WINDOWS_EXCEL_REQUIRED",
        ),
    )
    code = cli.main(["--mode", "production"])
    out = capsys.readouterr().out
    assert code == 3
    assert "Microsoft Excel Desktop on Windows is required" in out
    assert "Traceback" not in out
