# -*- mode: python ; coding: utf-8 -*-
"""PyInstaller spec — REMASEP Automation, variante **debug/consola**
(Sprint 3.11, Fase 4). Idéntica a ``remasep.spec`` salvo ``console=True`` y
nombre distinto (``dist/REMASEP-debug/``) para no pisar el build de piloto.
Sólo para diagnosticar errores de arranque/import en Windows; no se entrega
a la persona usuaria final.

Uso:  pyinstaller packaging/remasep_debug.spec --noconfirm
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
    name="REMASEP-debug",
    debug=True,
    bootloader_ignore_signals=False,
    strip=False,
    upx=False,
    console=True,  # con consola: ver tracebacks/imports fallidos en vivo
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
    name="REMASEP-debug",  # -> dist/REMASEP-debug/
)
