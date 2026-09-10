"""Sprint 3.7B — CLI generate_remasep (comportamiento sin Excel)."""

from __future__ import annotations

import generate_remasep as cli


def test_module_imports_without_win32com() -> None:
    # importar el CLI y el adaptador COM no debe fallar en Linux/WSL
    import remasep.adapters.excel_com  # noqa: F401

    assert hasattr(cli, "main")


def test_cli_is_graceful_when_excel_unavailable(monkeypatch, capsys):
    import remasep.services.generation_service as gs

    monkeypatch.setattr(
        gs, "detect_excel_capability",
        lambda **_: gs.ExcelCapability(
            platform="Linux", pywin32_available=False, excel_com_available=False,
            excel_version=None, can_generate=False, reason="WINDOWS_EXCEL_REQUIRED",
        ),
    )
    code = cli.main(["--mode", "diagnostic"])
    out = capsys.readouterr().out
    assert code == 3
    assert "Microsoft Excel Desktop on Windows is required" in out
    assert "Traceback" not in out


def test_mode_aliases_map_to_writer_modes() -> None:
    assert cli._MODE_ALIASES["diagnostic"] == cli.MODE_DIAGNOSTIC_REFERENCE
    assert cli._MODE_ALIASES["production"] == cli.MODE_PRODUCTION
    assert cli._PRODUCER_MODE[cli.MODE_DIAGNOSTIC_REFERENCE] == "LEGACY_EQUIVALENCE_DIAGNOSTIC"
    assert cli._PRODUCER_MODE[cli.MODE_PRODUCTION] == "PRODUCTION_PERIOD_SCOPE"
