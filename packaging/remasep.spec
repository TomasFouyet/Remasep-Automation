# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — REMASEP Automation (Sprint 3.11, piloto Windows ONEDIR).

Variante **windowed** (sin consola) — la que se entrega a la persona usuaria.
Para diagnóstico hay una variante con consola: ``remasep_debug.spec``.

Construye una distribución de **carpeta** (ONEDIR, no onefile — más fácil de
diagnosticar, carga de Qt más predecible, assets auditables, más simple para
Excel COM/pywin32; ver docs/WINDOWS_PACKAGING.md):

    dist/REMASEP/REMASEP.exe
    dist/REMASEP/_internal/...   (Python, Qt, config/runtime_2026,
                                   config/excel_writer_2026)

Uso:  pyinstaller packaging/remasep.spec --noconfirm
(o, más simple, ``scripts\\build_windows.ps1`` desde PowerShell en Windows).

Sólo bundlea lo necesario para el flujo validado Medinet -> Excel COM. NO
incluye ``data/local/``, ``artifacts/`` de desarrollo, ``outputs/``, el
workbook legacy (``GENERACION DATOS REMASEP.xlsx``), exports reales de
Medinet, ni plantillas reales del cliente. NO incluye QtWebEngine: ninguna
pantalla la usa (el PDF se genera con ``QPdfWriter``, nativo de ``QtGui``).
"""

import sys
from pathlib import Path

sys.path.insert(0, str(Path(SPECPATH).resolve()))
from _spec_common import common_paths  # noqa: E402

block_cipher = None
cfg = common_paths(SPECPATH)

a = Analysis(
    [cfg["entry_script"]],
    pathex=[str(cfg["src"])],
    binaries=[],
    datas=cfg["datas"],
    hiddenimports=cfg["hiddenimports"],
    hookspath=[],
    hooksconfig={},
    runtime_hooks=[],
    excludes=cfg["excludes"],
    noarchive=False,
    cipher=block_cipher,
)

pyz = PYZ(a.pure, a.zipped_data, cipher=block_cipher)

exe = EXE(
    pyz,
    a.scripts,
    [],
    exclude_binaries=True,
    name="REMASEP",
    debug=False,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,  # evita falsos positivos de antivirus en un piloto local
    console=False,  # windowed: sin consola para la persona usuaria (Fase 4/5)
    icon=str(cfg["icon"]) if cfg["icon"].is_file() else None,
    version=str(cfg["version_file"]) if cfg["version_file"].is_file() else None,
)

coll = COLLECT(
    exe,
    a.binaries,
    a.zipfiles,
    a.datas,
    strip=False,
    upx=False,
    name="REMASEP",  # -> dist/REMASEP/
)
