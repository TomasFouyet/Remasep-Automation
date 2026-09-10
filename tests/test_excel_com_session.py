"""Sprint 3.7B (patch de cierre) — espera de recálculo de ``ExcelSession``.

Sin Excel real: se inyecta una ``Application`` falsa y se controlan ``_sleep`` /
``_monotonic``.
"""

from __future__ import annotations

import pytest

import remasep.adapters.excel_com as ec
from remasep.services.excel_writer import ExcelWriterError

_XL_DONE = 0


class _FakeApp:
    """Application falsa: ``CalculationState`` recorre ``states`` en cada consulta."""

    def __init__(self, states, *, state_raises=None) -> None:
        self._states = list(states)
        self._state_raises = state_raises
        self.Calculation = None
        self.full_rebuild_calls = 0
        self.state_reads = 0

    def CalculateFullRebuild(self) -> None:
        self.full_rebuild_calls += 1

    @property
    def CalculationState(self):
        self.state_reads += 1
        if self._state_raises is not None:
            raise self._state_raises
        if self._states:
            return self._states.pop(0)
        return _XL_DONE


@pytest.fixture
def fake_clock(monkeypatch):
    """Reloj monótono controlable + sleep que sólo avanza el reloj."""
    now = {"t": 0.0}
    monkeypatch.setattr(ec, "_monotonic", lambda: now["t"])
    monkeypatch.setattr(ec, "_sleep", lambda seconds: now.__setitem__("t", now["t"] + seconds))
    return now


def _session(app) -> ec.ExcelSession:
    session = ec.ExcelSession(visible=False)
    session.excel = app  # se inyecta directamente; no se abre Excel
    return session


def test_calculate_full_returns_when_state_is_done_immediately(fake_clock):
    app = _FakeApp([_XL_DONE])
    _session(app).calculate_full(timeout=120.0)
    assert app.full_rebuild_calls == 1
    assert app.Calculation == ec._XL_CALC_AUTOMATIC
    assert app.state_reads == 1


def test_calculate_full_waits_several_polls_then_returns(fake_clock):
    app = _FakeApp([2, 1, 1, _XL_DONE])  # xlPending, xlCalculating, xlCalculating, xlDone
    _session(app).calculate_full(timeout=120.0)
    assert app.state_reads == 4
    assert fake_clock["t"] == pytest.approx(3 * ec.RECALC_POLL_SECONDS)


def test_calculate_full_raises_on_timeout_and_does_not_swallow(fake_clock):
    app = _FakeApp([1] * 10_000)  # nunca termina
    with pytest.raises(ExcelWriterError) as excinfo:
        _session(app).calculate_full(timeout=1.0)
    assert "no terminó en 1s" in str(excinfo.value)
    assert "no se guarda" in str(excinfo.value)


def test_calculate_full_raises_on_com_error_during_polling(fake_clock):
    app = _FakeApp([], state_raises=RuntimeError("RPC server unavailable"))
    with pytest.raises(ExcelWriterError) as excinfo:
        _session(app).calculate_full(timeout=120.0)
    assert "CalculationState" in str(excinfo.value)


def test_calculate_full_requires_a_session():
    session = ec.ExcelSession(visible=False)
    with pytest.raises(RuntimeError):
        session.calculate_full()


def test_teardown_is_safe_without_excel():
    session = ec.ExcelSession(visible=False)
    session._teardown()  # no debe lanzar
    session.__exit__(None, None, None)
    assert session.excel is None
