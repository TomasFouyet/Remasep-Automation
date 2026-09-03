"""Configuración compartida de tests.

Fuerza Qt en modo offscreen para que la maqueta de UI pueda testearse sin
display físico.
"""

import os

import pytest

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")


@pytest.fixture(scope="session")
def qapp():
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    yield app
