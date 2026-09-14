"""Sprint 3.11 — logging técnico de la app empaquetada (Fase 8).

Nunca un traceback en pantalla; el detalle va a un archivo bajo
``user_data_dir()/logs`` — nunca crece indefinidamente (rotación simple) y
nunca junto al ejecutable ni dentro del bundle.
"""

from __future__ import annotations

import logging
import sys
from logging.handlers import RotatingFileHandler

import pytest

from remasep import app_paths, logging_setup


@pytest.fixture(autouse=True)
def _isolated_app_data(tmp_path, monkeypatch):
    monkeypatch.setenv("REMASEP_APP_DATA_DIR", str(tmp_path / "appdata"))
    yield
    # cada test parte de un logger "remasep" limpio, sin handlers acumulados
    # de tests anteriores (logging_setup es idempotente por diseño, pero el
    # logger vive en el módulo `logging` global entre tests).
    logger = logging.getLogger("remasep")
    for handler in list(logger.handlers):
        logger.removeHandler(handler)
        handler.close()
    logging_setup._configured = False


def test_configure_logging_writes_to_a_file_under_logs_dir():
    logger = logging_setup.configure_logging()
    logger.info("mensaje técnico de prueba")
    for handler in logger.handlers:
        handler.flush()

    log_file = app_paths.logs_dir() / "remasep.log"
    assert log_file.is_file()
    assert "mensaje técnico de prueba" in log_file.read_text(encoding="utf-8")


def test_configure_logging_never_writes_next_to_the_bundle(tmp_path, monkeypatch):
    monkeypatch.setattr("sys._MEIPASS", str(tmp_path / "_internal"), raising=False)
    logging_setup.configure_logging()
    log_file = app_paths.logs_dir() / "remasep.log"
    assert not str(log_file).startswith(str(tmp_path / "_internal"))


def test_configure_logging_is_idempotent():
    logger1 = logging_setup.configure_logging()
    handlers_after_first = list(logger1.handlers)
    logger2 = logging_setup.configure_logging()
    assert logger2 is logger1
    assert logger2.handlers == handlers_after_first  # no duplica el handler


def test_log_handler_is_size_bounded_rotating_handler():
    logger = logging_setup.configure_logging()
    handlers = [h for h in logger.handlers if isinstance(h, RotatingFileHandler)]
    assert len(handlers) == 1
    handler = handlers[0]
    assert handler.maxBytes == logging_setup._MAX_BYTES
    assert handler.backupCount == logging_setup._BACKUP_COUNT


def test_configure_logging_never_raises_if_logs_dir_is_not_writable(monkeypatch):
    def _boom():
        raise OSError("disco lleno (simulado)")

    monkeypatch.setattr(app_paths, "logs_dir", _boom)
    logger = logging_setup.configure_logging()  # no debe lanzar
    assert logger.name == "remasep"


def test_guard_frozen_streams_replaces_none_streams(monkeypatch):
    monkeypatch.setattr("sys.stdout", None)
    monkeypatch.setattr("sys.stderr", None)
    logging_setup._guard_frozen_streams()
    assert sys.stdout is not None
    assert sys.stderr is not None
    print("no debe lanzar AttributeError aunque no haya consola")
