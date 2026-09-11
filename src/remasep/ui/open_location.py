"""Abrir un archivo o su carpeta con el gestor del sistema operativo.

Usa ``QDesktopServices`` (API segura de Qt): nunca lanza procesos a mano, nunca
cierra ni mata procesos de Excel del usuario.
"""

from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import QUrl
from PySide6.QtGui import QDesktopServices


def open_path(path: str | Path) -> bool:
    """Abre ``path`` (archivo o carpeta) con la aplicación asociada del SO."""
    target = Path(path)
    if not target.exists():
        return False
    return bool(QDesktopServices.openUrl(QUrl.fromLocalFile(str(target.resolve()))))


def open_containing_folder(path: str | Path) -> bool:
    """Abre la carpeta que contiene ``path``."""
    target = Path(path)
    folder = target if target.is_dir() else target.parent
    if not folder.exists():
        return False
    return bool(QDesktopServices.openUrl(QUrl.fromLocalFile(str(folder.resolve()))))
