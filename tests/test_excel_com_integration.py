"""Sprint 3.7B — tests de integración COM (Windows + Excel).

Se ejecutan **sólo** en Windows con Microsoft Excel Desktop y ``pywin32``. En
cualquier otro entorno se saltan limpiamente. Trabajan siempre sobre una COPIA
de una plantilla de test; nunca sobre un archivo fuente.

Ejecutar (en Windows)::

    pytest -m excel tests/test_excel_com_integration.py
"""

from __future__ import annotations

import platform
import shutil
from pathlib import Path

import pytest

from remasep.services import excel_writer as w
from remasep.services.excel_capability import detect_excel_capability

pytestmark = pytest.mark.excel

_TEMPLATE_ENV = "REMASEP_TEST_TEMPLATE"  # ruta a una plantilla .xlsm de prueba


def _requires_excel() -> None:
    if platform.system() != "Windows":
        pytest.skip("integración COM: sólo Windows")
    if not detect_excel_capability(probe_com=True).can_generate:
        pytest.skip("integración COM: Excel no disponible")


def _test_template(tmp_path: Path) -> Path:
    import os

    src = os.environ.get(_TEMPLATE_ENV)
    if not src or not Path(src).is_file():
        pytest.skip(f"define {_TEMPLATE_ENV} con una plantilla .xlsm de prueba")
    dst = tmp_path / "test_template.xlsm"
    shutil.copy2(src, dst)
    return dst


def test_com_open_write_recalc_save_reopen(tmp_path):
    _requires_excel()
    from remasep.adapters.excel_com import open_excel_com_writer

    template = _test_template(tmp_path)
    sha_before = w.compute_sha256(template)
    working = tmp_path / "out.__working__.xlsm"
    shutil.copy2(template, working)

    with open_excel_com_writer(working) as writer:
        assert writer.sheet_names()
        # se espera que el llamador conozca una celda desbloqueada real:
        pytest.skip("define aquí una celda de input real de la plantilla de prueba")

    assert w.compute_sha256(template) == sha_before  # fuente intacta


def test_com_capability_probe_is_stable():
    _requires_excel()
    cap = detect_excel_capability(probe_com=True)
    assert cap.can_generate is True
    assert cap.excel_version
