"""Config compartida entre ``remasep.spec`` (piloto, sin consola) y
``remasep_debug.spec`` (variante con consola, sólo para diagnóstico —
Sprint 3.11, Fase 4). Nada aquí usa los nombres que PyInstaller inyecta al
ejecutar un ``.spec`` (``Analysis``, ``EXE``, ...): es un módulo plano
importable, sin efectos de import.
"""

from __future__ import annotations

from pathlib import Path


def common_paths(specpath: str) -> dict:
    packaging_dir = Path(specpath).resolve()
    root = packaging_dir.parent
    config = root / "config"

    datas = [
        (str(config / "runtime_2026"), "config/runtime_2026"),
        (str(config / "excel_writer_2026"), "config/excel_writer_2026"),
    ]

    # win32com/pythoncom/pywintypes se importan de forma perezosa (funciones,
    # no nivel de módulo) para seguir siendo importables en Linux/WSL — el
    # análisis estático de PyInstaller no los ve solo. win32timezone es un
    # requisito de runtime de pywin32+PyInstaller ampliamente documentado.
    hiddenimports = [
        "win32com",
        "win32com.client",
        "win32timezone",
        "pythoncom",
        "pywintypes",
    ]

    # La app sólo importa QtCore/QtGui/QtWidgets (confirmado por grep sobre
    # src/); nunca QtWebEngine ni el resto de módulos pesados no usados.
    excludes = [
        "PySide6.QtWebEngineCore",
        "PySide6.QtWebEngineWidgets",
        "PySide6.QtWebEngineQuick",
        "PySide6.QtQml",
        "PySide6.QtQuick",
        "PySide6.QtQuickWidgets",
        "PySide6.QtQuick3D",
        "PySide6.Qt3DCore",
        "PySide6.Qt3DRender",
        "PySide6.Qt3DAnimation",
        "PySide6.Qt3DExtras",
        "PySide6.QtMultimedia",
        "PySide6.QtMultimediaWidgets",
        "PySide6.QtBluetooth",
        "PySide6.QtSensors",
        "PySide6.QtWebChannel",
        "PySide6.QtPositioning",
        "PySide6.QtPositioningQuick",
        "PySide6.QtNfc",
        "PySide6.QtDesigner",
        "PySide6.QtHelp",
        "PySide6.QtRemoteObjects",
        "PySide6.QtSql",
        "PySide6.QtTest",
    ]

    return {
        "root": root,
        "src": root / "src",
        "entry_script": str(packaging_dir / "run_remasep.py"),
        "icon": packaging_dir / "assets" / "remasep.ico",
        "version_file": packaging_dir / "version_info.txt",
        "datas": datas,
        "hiddenimports": hiddenimports,
        "excludes": excludes,
    }
