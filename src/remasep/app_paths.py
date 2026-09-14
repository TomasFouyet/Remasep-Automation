"""Resolución de rutas de la aplicación empaquetada (Sprint 3.11).

Dos categorías, nunca mezcladas:

- **INTERNAL_READ_ONLY** — assets que viajan DENTRO de la app
  (``config/runtime_2026``, ``config/excel_writer_2026``). En desarrollo se
  resuelven subiendo desde este módulo hasta encontrar ``config/<nombre>``; en
  un build PyInstaller (onedir) se resuelven desde ``sys._MEIPASS``. **Nunca**
  dependen del *cwd*.
- **APP_WRITABLE_DATA** — logs y diagnósticos técnicos propios de la app
  (nunca el Medinet, la plantilla, ni el REMASEP/PDF generado — esos los
  elige la persona usuaria con un diálogo y viven donde ella decida). Viven
  bajo ``%LOCALAPPDATA%\\REMASEP Automation\\`` en Windows, o
  ``~/.local/share/REMASEP Automation`` fuera de Windows (desarrollo);
  **nunca** junto al ejecutable ni dentro del bundle de PyInstaller.

``REMASEP_APP_DATA_DIR`` permite redirigir ``user_data_dir()`` (tests, o un
piloto que quiera aislar sus datos) sin tocar el resto de la resolución.
"""

from __future__ import annotations

import os
import sys
from pathlib import Path

APP_DIR_NAME = "REMASEP Automation"
_ENV_APP_DATA_DIR = "REMASEP_APP_DATA_DIR"

# Marcadores que distinguen el `config/` de datos versionadas (raíz del repo /
# bundle) de cualquier otro directorio llamado "config" que exista por el
# camino — p. ej. el subpaquete Python (no relacionado, no borrado por este
# sprint) ``src/remasep/config/``. Sin esto, subir por los `parents` de este
# módulo puede detenerse en el subpaquete equivocado.
_ROOT_MARKERS = ("runtime_2026", "excel_writer_2026")


def is_frozen() -> bool:
    """``True`` dentro de un ejecutable PyInstaller (onedir u onefile)."""
    return bool(getattr(sys, "frozen", False)) and hasattr(sys, "_MEIPASS")


def resource_root() -> Path:
    """Raíz desde la que cuelga ``config/`` — nunca el *cwd*.

    Empaquetado: ``sys._MEIPASS`` (en onedir, la carpeta ``_internal``). En
    desarrollo: el directorio del repo, encontrado subiendo desde este mismo
    archivo hasta dar con una carpeta ``config/`` real.
    """
    meipass = getattr(sys, "_MEIPASS", None)
    if meipass:
        return Path(meipass)
    here = Path(__file__).resolve()
    for parent in here.parents:
        cfg = parent / "config"
        if cfg.is_dir() and all((cfg / marker).is_dir() for marker in _ROOT_MARKERS):
            return parent
    # No debería ocurrir en un checkout válido; último recurso determinista.
    return here.parents[3] if len(here.parents) > 3 else here.parent


def find_config_dir(name: str, *, required_file: str | None = None) -> Path:
    """Localiza ``config/<name>/`` sin depender del *cwd*.

    Busca primero bajo :func:`resource_root`; si no está ahí (estructura de
    empaquetado distinta a la esperada), recorre los ``parents`` de este
    módulo como respaldo — el mismo patrón que
    :func:`remasep.services.runtime_assets.resolve_runtime_root`. Devuelve la
    ruta candidata **aunque no exista**: quien llama decide cómo reportar
    "falta" con su propio tipo de error de dominio (p. ej.
    ``RuntimeAssetError`` / ``GenerationServiceError``); esta función sólo
    resuelve la ubicación, no valida el contenido.
    """
    candidates = [resource_root() / "config" / name]
    here = Path(__file__).resolve()
    candidates += [parent / "config" / name for parent in here.parents]

    for candidate in candidates:
        if required_file is None:
            if candidate.is_dir():
                return candidate
        elif (candidate / required_file).is_file():
            return candidate
    return candidates[0]


def runtime_config_dir() -> Path:
    """``config/runtime_2026/`` — delega en el resolver ya validado (Sprint 3.8)."""
    from remasep.services.runtime_assets import resolve_runtime_root

    return resolve_runtime_root()


def excel_writer_config_dir() -> Path:
    """``config/excel_writer_2026/`` (política + mapa de CONTROL del writer)."""
    return find_config_dir("excel_writer_2026", required_file="policy.yaml")


def user_data_dir() -> Path:
    """Carpeta de datos propios de la app, escribible. Nunca junto al .exe
    ni dentro del bundle de PyInstaller. Se crea si no existe."""
    override = os.environ.get(_ENV_APP_DATA_DIR)
    if override:
        root = Path(override)
    else:
        local_appdata = os.environ.get("LOCALAPPDATA")
        root = (
            Path(local_appdata) / APP_DIR_NAME
            if local_appdata
            else Path.home() / ".local" / "share" / APP_DIR_NAME
        )
    root.mkdir(parents=True, exist_ok=True)
    return root


def logs_dir() -> Path:
    d = user_data_dir() / "logs"
    d.mkdir(parents=True, exist_ok=True)
    return d


def diagnostics_dir() -> Path:
    d = user_data_dir() / "diagnostics"
    d.mkdir(parents=True, exist_ok=True)
    return d
