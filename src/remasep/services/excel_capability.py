"""Diagnóstico de plataforma para el Excel COM writer (Sprint 3.7B).

Importable en cualquier sistema operativo: **no** importa ``win32com`` a nivel de
módulo. En Linux/WSL devuelve ``can_generate = False`` con una razón legible, sin
traceback técnico.
"""

from __future__ import annotations

import contextlib
import importlib.util
import platform
import sys
from dataclasses import dataclass, field

WINDOWS_EXCEL_REQUIRED = "WINDOWS_EXCEL_REQUIRED"
PYWIN32_NOT_INSTALLED = "PYWIN32_NOT_INSTALLED"
EXCEL_COM_NOT_AVAILABLE = "EXCEL_COM_NOT_AVAILABLE"
READY = "READY"

_USER_MESSAGE = (
    "Microsoft Excel Desktop on Windows is required to generate the REMASEP "
    "workbook. This environment cannot run it."
)


@dataclass(frozen=True)
class ExcelCapability:
    platform: str
    pywin32_available: bool
    excel_com_available: bool
    excel_version: str | None
    can_generate: bool
    reason: str
    details: tuple[str, ...] = field(default_factory=tuple)

    @property
    def user_message(self) -> str:
        return READY if self.can_generate else _USER_MESSAGE

    def as_dict(self) -> dict:
        return {
            "platform": self.platform,
            "pywin32_available": self.pywin32_available,
            "excel_com_available": self.excel_com_available,
            "excel_version": self.excel_version,
            "can_generate": self.can_generate,
            "reason": self.reason,
            "details": list(self.details),
        }


def _pywin32_installed() -> bool:
    try:
        return importlib.util.find_spec("win32com.client") is not None
    except (ImportError, ValueError):  # pragma: no cover - defensivo
        return False


def detect_excel_capability(*, probe_com: bool = True) -> ExcelCapability:
    """Detecta si este entorno puede generar el REMASEP con Excel COM.

    ``probe_com=False`` evita instanciar Excel (útil en tests): sólo mira
    plataforma y disponibilidad de ``pywin32``.
    """
    system = platform.system()
    details: list[str] = [f"python={sys.version.split()[0]}"]

    if system != "Windows":
        return ExcelCapability(
            platform=system or "unknown",
            pywin32_available=False,
            excel_com_available=False,
            excel_version=None,
            can_generate=False,
            reason=WINDOWS_EXCEL_REQUIRED,
            details=tuple(details),
        )

    has_pywin32 = _pywin32_installed()
    if not has_pywin32:
        return ExcelCapability(
            platform=system,
            pywin32_available=False,
            excel_com_available=False,
            excel_version=None,
            can_generate=False,
            reason=PYWIN32_NOT_INSTALLED,
            details=tuple(details),
        )

    if not probe_com:
        return ExcelCapability(
            platform=system,
            pywin32_available=True,
            excel_com_available=False,
            excel_version=None,
            can_generate=False,
            reason=EXCEL_COM_NOT_AVAILABLE,
            details=(*details, "com_probe_skipped"),
        )

    excel_version, com_ok, note = _probe_excel_com()
    details.append(note)
    if not com_ok:
        return ExcelCapability(
            platform=system,
            pywin32_available=True,
            excel_com_available=False,
            excel_version=excel_version,
            can_generate=False,
            reason=EXCEL_COM_NOT_AVAILABLE,
            details=tuple(details),
        )

    return ExcelCapability(
        platform=system,
        pywin32_available=True,
        excel_com_available=True,
        excel_version=excel_version,
        can_generate=True,
        reason=READY,
        details=tuple(details),
    )


def _probe_excel_com() -> tuple[str | None, bool, str]:  # pragma: no cover - Windows only
    """Instancia una copia aislada de Excel, lee su versión y la cierra."""
    try:
        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]
    except ImportError as exc:
        return None, False, f"import_error={exc.__class__.__name__}"

    pythoncom.CoInitialize()
    excel = None
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        version = str(excel.Version)
        return version, True, "com_probe_ok"
    except Exception as exc:  # noqa: BLE001 - se resume a un motivo legible
        return None, False, f"com_error={exc.__class__.__name__}"
    finally:
        if excel is not None:
            with contextlib.suppress(Exception):
                excel.Quit()
        pythoncom.CoUninitialize()
