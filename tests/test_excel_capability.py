"""Sprint 3.7B — diagnóstico de capacidad de plataforma (sin Excel)."""

from __future__ import annotations

import sys
import types
import weakref

import pytest

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


# ---------------------------------------------------------------------------
# Ciclo de vida COM del probe — con pythoncom/win32com falsos (sin Windows)
# ---------------------------------------------------------------------------


class _ComError(Exception):
    """Sustituto de ``pythoncom.com_error``."""


class _FakePythoncom:
    def __init__(self, order: list[str]) -> None:
        self._order = order
        self.co_init = 0
        self.co_uninit = 0
        self.com_error = _ComError

    def CoInitialize(self) -> None:
        self.co_init += 1
        self._order.append("CoInitialize")

    def CoUninitialize(self) -> None:
        self.co_uninit += 1
        self._order.append("CoUninitialize")


class _FakeExcelApp:
    def __init__(self, order: list[str], *, quit_raises: bool = False) -> None:
        self._order = order
        self._quit_raises = quit_raises
        self.Visible = None
        self.DisplayAlerts = None
        self.Version = "16.0"
        self.quit_calls = 0

    def Quit(self) -> None:
        self.quit_calls += 1
        self._order.append("Quit")
        if self._quit_raises:
            raise RuntimeError("Quit falló (Excel ya no responde)")

    def __del__(self):  # marca el Release() del proxy
        self._order.append("released")


class _FakeWin32ComClient:
    def __init__(self, order: list[str], *, dispatch_raises=None, quit_raises=False) -> None:
        self._order = order
        self._dispatch_raises = dispatch_raises
        self._quit_raises = quit_raises
        self.last_app_ref = None

    def DispatchEx(self, prog_id: str):
        assert prog_id == "Excel.Application"
        self._order.append("DispatchEx")
        if self._dispatch_raises is not None:
            raise self._dispatch_raises
        app = _FakeExcelApp(self._order, quit_raises=self._quit_raises)
        # weakref: no retenemos el proxy, para poder observar su Release()
        self.last_app_ref = weakref.ref(app)
        return app


def _install_fake_com(monkeypatch, order, **client_kw):
    fake_pythoncom = _FakePythoncom(order)
    fake_client = _FakeWin32ComClient(order, **client_kw)
    fake_win32com = types.ModuleType("win32com")
    fake_win32com.client = fake_client
    monkeypatch.setitem(sys.modules, "pythoncom", fake_pythoncom)
    monkeypatch.setitem(sys.modules, "win32com", fake_win32com)
    monkeypatch.setitem(sys.modules, "win32com.client", fake_client)
    monkeypatch.setattr(cap.platform, "system", lambda: "Windows")
    monkeypatch.setattr(cap, "_pywin32_installed", lambda: True)
    return fake_pythoncom, fake_client


def test_probe_releases_com_before_uninitialize_and_balances_init(monkeypatch):
    order: list[str] = []
    fake_pythoncom, fake_client = _install_fake_com(monkeypatch, order)

    version, ok, note = cap._probe_excel_com()

    assert (version, ok, note) == ("16.0", True, "com_probe_ok")
    assert fake_pythoncom.co_init == fake_pythoncom.co_uninit == 1
    assert order.count("Quit") == 1
    assert fake_client.last_app_ref() is None  # el proxy fue liberado
    # el proxy se libera ANTES de desmontar el apartment
    assert order.index("released") < order.index("CoUninitialize")
    assert order.index("Quit") < order.index("released")
    assert order[0] == "CoInitialize" and order[-1] == "CoUninitialize"


def test_repeated_probe_is_stable(monkeypatch):
    order: list[str] = []
    fake_pythoncom, _ = _install_fake_com(monkeypatch, order)

    results = [cap.detect_excel_capability(probe_com=True) for _ in range(5)]

    assert all(r.can_generate for r in results)
    assert {r.excel_version for r in results} == {"16.0"}
    # CoInitialize / CoUninitialize siempre balanceados, una vez por llamada
    assert fake_pythoncom.co_init == fake_pythoncom.co_uninit == 5
    assert order.count("CoUninitialize") == 5


def test_probe_dispatch_failure_still_balances_and_reports(monkeypatch):
    order: list[str] = []
    fake_pythoncom, _ = _install_fake_com(
        monkeypatch, order, dispatch_raises=RuntimeError("Excel no instalado")
    )
    version, ok, note = cap._probe_excel_com()
    assert version is None and ok is False
    assert note.startswith("com_error=")
    assert fake_pythoncom.co_init == fake_pythoncom.co_uninit == 1
    assert "CoUninitialize" in order


def test_probe_quit_failure_is_defensive(monkeypatch):
    order: list[str] = []
    fake_pythoncom, _ = _install_fake_com(monkeypatch, order, quit_raises=True)
    version, ok, note = cap._probe_excel_com()
    # la versión ya se leyó: el probe reporta éxito con nota de quit_error
    assert version == "16.0" and ok is True
    assert "quit_error=" in note
    assert fake_pythoncom.co_init == fake_pythoncom.co_uninit == 1
    assert order[-1] == "CoUninitialize"


def test_probe_import_error_is_reported(monkeypatch):
    monkeypatch.setattr(cap.platform, "system", lambda: "Windows")
    monkeypatch.setattr(cap, "_pywin32_installed", lambda: True)
    monkeypatch.setitem(sys.modules, "pythoncom", None)  # fuerza ImportError
    version, ok, note = cap._probe_excel_com()
    assert version is None and ok is False
    assert note.startswith("import_error=")


def test_describe_com_error_formats_hresult():
    from remasep.services.excel_writer import describe_com_error

    class _E(Exception):
        pass

    err = _E()
    err.hresult = -2147417848  # 0x80010108 RPC_E_DISCONNECTED
    assert describe_com_error(err) == "_E(0x80010108)"
    assert describe_com_error(ValueError("x")) == "ValueError"


@pytest.mark.parametrize("hresult", [0x80010108, 0x800706BE, 0x800706BA])
def test_describe_com_error_from_args(hresult):
    from remasep.services.excel_writer import describe_com_error

    class _Com(Exception):
        pass

    assert f"0x{hresult:08X}" in describe_com_error(_Com(hresult, "desc"))
