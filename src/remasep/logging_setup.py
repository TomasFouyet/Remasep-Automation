"""Logging técnico para la app empaquetada (Sprint 3.11).

La persona usuaria **nunca** ve un traceback ni una consola: los mensajes en
pantalla ya son humanos (``remasep.ui.errors``). El detalle técnico va a un
archivo local bajo ``app_paths.logs_dir()`` — nunca junto al ejecutable, nunca
dentro del bundle de PyInstaller. Los mensajes que el resto de la app ya loguea
(``remasep.ui`` y afines) están redactados para no incluir PII (nombre, RUT,
fecha de nacimiento, teléfono, email, filas de Medinet); este módulo sólo les
da un destino real y evita que el log crezca indefinidamente
(``RotatingFileHandler`` — sin rotación sofisticada, alcanza para un piloto).
"""

from __future__ import annotations

import logging
import os
import sys
from logging.handlers import RotatingFileHandler

from remasep import app_paths

_LOG_FILE = "remasep.log"
_MAX_BYTES = 1_000_000  # ~1 MB por archivo
_BACKUP_COUNT = 2  # remasep.log + remasep.log.1 + remasep.log.2

_configured = False


def _guard_frozen_streams() -> None:
    """En un build ``--windowed`` de PyInstaller, ``sys.stdout``/``stderr``
    pueden ser ``None``: cualquier ``print()`` o el ``StreamHandler`` por
    defecto de ``logging`` explotaría con ``AttributeError``. Los redirige a
    un sumidero para que nada dependa de una consola inexistente."""
    if sys.stdout is None:
        sys.stdout = open(os.devnull, "w")  # noqa: SIM115 - vive todo el proceso
    if sys.stderr is None:
        sys.stderr = open(os.devnull, "w")  # noqa: SIM115


def configure_logging() -> logging.Logger:
    """Engancha un ``RotatingFileHandler`` al logger ``remasep`` (idempotente).

    Nunca lanza: si no se puede escribir el log (permisos, disco), la app
    sigue arrancando sin logging de archivo en vez de morir en el intento.
    """
    global _configured
    logger = logging.getLogger("remasep")
    if _configured:
        return logger

    _guard_frozen_streams()
    logger.setLevel(logging.INFO)
    try:
        handler = RotatingFileHandler(
            app_paths.logs_dir() / _LOG_FILE,
            maxBytes=_MAX_BYTES,
            backupCount=_BACKUP_COUNT,
            encoding="utf-8",
        )
        handler.setFormatter(
            logging.Formatter("%(asctime)s %(levelname)s %(name)s: %(message)s")
        )
        logger.addHandler(handler)
    except OSError:
        pass
    _configured = True
    return logger
