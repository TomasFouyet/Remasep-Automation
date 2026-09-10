"""Sprint 3.7B — tests de integración COM (Windows + Microsoft Excel Desktop).

Se ejecutan **sólo** en Windows con Excel + ``pywin32``. En cualquier otro
entorno se saltan limpiamente. Trabajan **siempre** sobre una copia en
``tmp_path``; nunca modifican el archivo fuente.

La hoja/celda de prueba se pasa por variables de entorno (no se elige ninguna
celda automáticamente):

    REMASEP_TEST_TEMPLATE   ruta a una plantilla real REMASEP .xlsm
    REMASEP_TEST_SHEET      hoja con una celda de input desbloqueada
    REMASEP_TEST_CELL       celda de input (sin fórmula, desbloqueada, no merge)

Ejecutar (PowerShell)::

    $env:REMASEP_TEST_TEMPLATE = "C:\\...\\REMASEP_V1.4 Julio 2026.xlsm"
    $env:REMASEP_TEST_SHEET    = "REMASEP_OD"
    $env:REMASEP_TEST_CELL     = "K17"
    pytest -m excel tests/test_excel_com_integration.py -vv -s
"""

from __future__ import annotations

import os
import platform
import shutil
from pathlib import Path

import pytest

from remasep.services import excel_writer as w
from remasep.services.excel_capability import detect_excel_capability

pytestmark = pytest.mark.excel

_ENV_TEMPLATE = "REMASEP_TEST_TEMPLATE"
_ENV_SHEET = "REMASEP_TEST_SHEET"
_ENV_CELL = "REMASEP_TEST_CELL"

# Entero seguro, improbable como valor real de una métrica: fácil de reconocer.
_SENTINEL = 424242


def _skip_if_no_excel() -> None:
    if platform.system() != "Windows":
        pytest.skip("integración COM: sólo Windows")
    capability = detect_excel_capability(probe_com=True)
    if not capability.can_generate:
        pytest.skip(f"integración COM: Excel no disponible ({capability.reason})")


def _env(name: str) -> str:
    value = os.environ.get(name)
    if not value:
        pytest.skip(f"integración COM: define la variable de entorno {name}")
    return value


def _copied_template(tmp_path: Path) -> tuple[Path, Path]:
    src = Path(_env(_ENV_TEMPLATE))
    if not src.is_file():
        pytest.skip(f"{_ENV_TEMPLATE} no apunta a un archivo: {src}")
    if src.suffix.lower() != ".xlsm":
        pytest.skip(f"{_ENV_TEMPLATE} debe ser un .xlsm")
    dst = tmp_path / "REMASEP_TEST_COPY.xlsm"
    shutil.copy2(src, dst)
    return src, dst


def test_com_open_write_recalc_save_reopen(tmp_path):
    """Abre una copia, valida la celda, escribe un sentinel, recalcula, guarda,
    reabre en una sesión nueva y confirma persistencia. La fuente queda intacta."""
    _skip_if_no_excel()
    from remasep.adapters.excel_com import open_excel_com_writer

    source_template, template_copy = _copied_template(tmp_path)
    sheet = _env(_ENV_SHEET)
    cell = _env(_ENV_CELL)

    source_sha_before = w.compute_sha256(source_template)
    copy_sha_before = w.compute_sha256(template_copy)
    working = tmp_path / "REMASEP_TEST_OUT.xlsm"
    shutil.copy2(template_copy, working)

    # 1) abrir la copia de trabajo y validar la celda (defense-in-depth)
    with open_excel_com_writer(working) as writer:
        assert writer.has_sheet(sheet), f"la hoja {sheet!r} no existe"
        assert not writer.cell_has_formula(sheet, cell), f"{sheet}!{cell} contiene fórmula"
        assert not writer.cell_in_incompatible_merge(sheet, cell), (
            f"{sheet}!{cell} pertenece a un merge incompatible"
        )
        assert writer.cell_is_writable(sheet, cell), (
            f"{sheet}!{cell} no es escribible (hoja protegida y celda locked)"
        )
        original_value = writer.read_cell(sheet, cell)

        writer.write_value2(sheet, cell, _SENTINEL)
        assert int(writer.read_cell(sheet, cell)) == _SENTINEL

        writer.recalculate()  # espera real hasta xlDone
        writer.save()

    # 2) reabrir en una sesión COM NUEVA y confirmar que el valor persistió
    with open_excel_com_writer(working) as reopened:
        assert int(reopened.read_cell(sheet, cell)) == _SENTINEL

    # 3) el archivo FUENTE conserva exactamente su SHA256 inicial
    assert w.compute_sha256(source_template) == source_sha_before
    assert w.compute_sha256(template_copy) == copy_sha_before

    # 4) integridad SEMÁNTICA de VBA: el Save de Excel puede reescribir
    #    vbaProject.bin, pero el código de las macros NO debe cambiar.
    from remasep.services.vba_integrity import (
        VBA_FAIL,
        compare_vba_projects,
        read_vba_project,
    )

    before_vba = read_vba_project(source_template)
    after_vba = read_vba_project(working)
    vba_cmp = compare_vba_projects(before_vba, after_vba)
    print(
        f"[com-int] VBA: status={vba_cmp.status} "
        f"binario_estable={vba_cmp.payload_stable} "
        f"modulos_before={len(vba_cmp.module_names_before)} "
        f"code_changed={vba_cmp.code_changed_modules} "
        f"attrs_changed={vba_cmp.attributes_changed_modules}"
    )
    assert vba_cmp.status != VBA_FAIL, vba_cmp.fail_reasons
    assert vba_cmp.code_changed_modules == ()
    assert vba_cmp.removed_modules == () and vba_cmp.added_modules == ()

    print(
        f"[com-int] {sheet}!{cell}: valor original = {original_value!r} -> "
        f"sentinel {_SENTINEL} escrito y verificado en {working.name}"
    )


def test_com_capability_probe_is_stable_across_repeats():
    """Llamadas repetidas a ``detect_excel_capability(probe_com=True)`` no dejan
    procesos huérfanos ni degradan (misma versión, sin excepciones)."""
    _skip_if_no_excel()

    versions = []
    for _ in range(3):
        capability = detect_excel_capability(probe_com=True)
        assert capability.can_generate is True
        assert capability.excel_com_available is True
        assert capability.excel_version
        versions.append(capability.excel_version)

    assert len(set(versions)) == 1
