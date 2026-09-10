"""Sprint 3.7B — diagnóstico de capacidad de plataforma (sin Excel)."""

from __future__ import annotations

import remasep.services.excel_capability as cap


def test_linux_reports_unavailable_gracefully(monkeypatch):
    monkeypatch.setattr(cap.platform, "system", lambda: "Linux")
    result = cap.detect_excel_capability(probe_com=True)
    assert result.can_generate is False
    assert result.reason == cap.WINDOWS_EXCEL_REQUIRED
    assert result.pywin32_available is False
    assert result.excel_com_available is False
    assert "Windows" in result.user_message
    assert "Traceback" not in result.user_message


def test_windows_without_pywin32_is_unavailable(monkeypatch):
    monkeypatch.setattr(cap.platform, "system", lambda: "Windows")
    monkeypatch.setattr(cap, "_pywin32_installed", lambda: False)
    result = cap.detect_excel_capability(probe_com=True)
    assert result.can_generate is False
    assert result.reason == cap.PYWIN32_NOT_INSTALLED


def test_windows_with_pywin32_no_com_probe(monkeypatch):
    monkeypatch.setattr(cap.platform, "system", lambda: "Windows")
    monkeypatch.setattr(cap, "_pywin32_installed", lambda: True)
    result = cap.detect_excel_capability(probe_com=False)
    assert result.pywin32_available is True
    assert result.excel_com_available is False
    assert result.can_generate is False
    assert result.reason == cap.EXCEL_COM_NOT_AVAILABLE


def test_windows_com_probe_success(monkeypatch):
    monkeypatch.setattr(cap.platform, "system", lambda: "Windows")
    monkeypatch.setattr(cap, "_pywin32_installed", lambda: True)
    monkeypatch.setattr(cap, "_probe_excel_com", lambda: ("16.0", True, "com_probe_ok"))
    result = cap.detect_excel_capability(probe_com=True)
    assert result.can_generate is True
    assert result.excel_version == "16.0"
    assert result.reason == cap.READY
    assert result.user_message == cap.READY


def test_capability_as_dict_is_json_safe(monkeypatch):
    monkeypatch.setattr(cap.platform, "system", lambda: "Linux")
    data = cap.detect_excel_capability(probe_com=False).as_dict()
    assert data["platform"] == "Linux"
    assert data["can_generate"] is False
    assert isinstance(data["details"], list)
