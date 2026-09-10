"""Adaptador de Microsoft Excel Desktop vía COM (Sprint 3.7B).

**Sólo Windows.** ``win32com`` se importa de forma perezosa dentro de los
métodos: importar este módulo en Linux/WSL no falla. La escritura del REMASEP
oficial se hace exclusivamente con esta ruta (Excel COM), nunca con
``openpyxl.save()``.

Ciclo de vida COM determinista (Sprint 3.7B — patch de cierre): se sueltan
**todas** las referencias COM y se fuerza su ``Release()`` (``= None`` +
``gc.collect()``) **antes** de ``CoUninitialize`` / de que muera el proceso de
Excel, para no dejar proxies que el GC de Python liberaría en un apartment ya
desmontado (causa de ``RPC_E_DISCONNECTED`` / ``RPC server unavailable``).
"""

from __future__ import annotations

import contextlib
import gc
import platform
import time
from collections.abc import Iterator
from contextlib import AbstractContextManager, contextmanager
from pathlib import Path

from remasep.core.errors import TemplateValidationError
from remasep.services.excel_writer import (
    ExcelWriterError,
    WorkbookSnapshot,
    describe_com_error,
)

# Constantes de Excel (evitan depender de la typelib).
_XL_CALC_MANUAL = -4135
_XL_CALC_AUTOMATIC = -4105
_XL_CALC_STATE_DONE = 0  # xlDone
_XL_OPENXML_MACRO_ENABLED = 52

# Espera de recálculo.
RECALC_TIMEOUT_SECONDS = 120.0
RECALC_POLL_SECONDS = 0.1

# Puntos de inyección para tests (se monkeypatchean; en producción son los reales).
_sleep = time.sleep
_monotonic = time.monotonic


class ExcelSession(AbstractContextManager):
    """Sesión aislada de Microsoft Excel (``DispatchEx``, no instancia compartida)."""

    def __init__(self, *, visible: bool = False):
        self.visible = visible
        self.excel = None
        self._co_initialized = False

    def __enter__(self):
        if platform.system() != "Windows":
            raise RuntimeError("Excel COM solamente está disponible en Windows.")

        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]

        pythoncom.CoInitialize()
        self._co_initialized = True
        try:
            self.excel = win32com.client.DispatchEx("Excel.Application")
            self.excel.Visible = self.visible
            self.excel.DisplayAlerts = False
        except BaseException:
            self._teardown()
            raise
        return self

    def open_workbook(self, path: str | Path, *, read_only: bool = False):
        if self.excel is None:
            raise RuntimeError("La sesión Excel no está inicializada.")
        return self.excel.Workbooks.Open(
            str(Path(path).resolve()),
            ReadOnly=read_only,
        )

    def calculate_full(self, *, timeout: float = RECALC_TIMEOUT_SECONDS) -> None:
        """Recálculo completo con **espera real** hasta ``xlDone``.

        Llama ``CalculateFullRebuild()`` y luego consulta
        ``Application.CalculationState`` en bucle usando ``time.monotonic()`` y
        un ``sleep`` breve entre consultas. Si el cálculo no termina dentro de
        ``timeout`` segundos lanza :class:`ExcelWriterError` y **no** se guarda
        nada. No se ocultan las excepciones que impidan saber si terminó.
        """
        if self.excel is None:
            raise RuntimeError("La sesión Excel no está inicializada.")
        self.excel.Calculation = _XL_CALC_AUTOMATIC
        self.excel.CalculateFullRebuild()

        deadline = _monotonic() + timeout
        while True:
            try:
                state = int(self.excel.CalculationState)
            except Exception as exc:
                raise ExcelWriterError(
                    "no se pudo consultar Application.CalculationState durante el "
                    f"recálculo: {describe_com_error(exc)}"
                ) from exc
            if state == _XL_CALC_STATE_DONE:
                return
            if _monotonic() >= deadline:
                raise ExcelWriterError(
                    f"el recálculo de Excel no terminó en {timeout:.0f}s "
                    f"(Application.CalculationState={state}); no se guarda el workbook"
                )
            _sleep(RECALC_POLL_SECONDS)

    # -- teardown determinista -------------------------------------
    def _teardown(self) -> None:
        excel = self.excel
        self.excel = None
        if excel is not None:
            # Quit defensivo: Excel puede haber muerto ya.
            with contextlib.suppress(Exception):
                excel.DisplayAlerts = False
            with contextlib.suppress(Exception):
                excel.Quit()
        # Soltar el proxy y forzar su Release() ANTES de desmontar el apartment.
        excel = None
        gc.collect()
        if self._co_initialized:
            self._co_initialized = False
            with contextlib.suppress(Exception):
                import pythoncom  # type: ignore[import-not-found]

                pythoncom.CoUninitialize()

    def __exit__(self, exc_type, exc, tb):
        self._teardown()
        return False


def validate_required_sheets(workbook, required_sheets: list[str]) -> None:
    existing = {sheet.Name for sheet in workbook.Worksheets}
    missing = [sheet for sheet in required_sheets if sheet not in existing]
    if missing:
        raise TemplateValidationError(
            "La plantilla no contiene hojas requeridas: " + ", ".join(missing)
        )


class ExcelComWorkbookWriter:
    """Implementación COM del contrato ``remasep.services.excel_writer.WorkbookWriter``.

    Trabaja siempre sobre la ruta que se le pasa (la **copia de trabajo**), nunca
    sobre la plantilla. No ejecuta macros, no llama ``Application.Run``, no toca
    Trust Center / Trusted Locations / protección global.
    """

    def __init__(self, session: ExcelSession, path: str | Path) -> None:
        self._session = session
        self._path = str(Path(path).resolve())
        self._wb = session.open_workbook(self._path)
        self._sheet_cache: dict[str, object] = {}

    # -- hojas / celdas ----------------------------------------------
    def sheet_names(self) -> list[str]:
        return [ws.Name for ws in self._wb.Worksheets]

    def has_sheet(self, sheet: str) -> bool:
        return sheet in set(self.sheet_names())

    def _ws(self, sheet: str):
        if sheet not in self._sheet_cache:
            self._sheet_cache[sheet] = self._wb.Worksheets(sheet)
        return self._sheet_cache[sheet]

    def _range(self, sheet: str, cell: str):
        return self._ws(sheet).Range(cell)

    def cell_has_formula(self, sheet: str, cell: str) -> bool:
        return bool(self._range(sheet, cell).HasFormula)

    def cell_in_incompatible_merge(self, sheet: str, cell: str) -> bool:
        rng = self._range(sheet, cell)
        if not bool(rng.MergeCells):
            return False
        # Un merge que no sea exactamente 1x1 es incompatible con escribir un
        # único valor de forma predecible.
        area = rng.MergeArea
        return not (area.Rows.Count == 1 and area.Columns.Count == 1)

    def cell_is_writable(self, sheet: str, cell: str) -> bool:
        ws = self._ws(sheet)
        rng = self._range(sheet, cell)
        if not bool(ws.ProtectContents):
            return True
        return not bool(rng.Locked)

    def read_cell(self, sheet: str, cell: str) -> object:
        return self._range(sheet, cell).Value2

    # -- escritura -------------------------------------------------
    def write_value2(self, sheet: str, cell: str, value: object) -> None:
        rng = self._range(sheet, cell)
        if bool(rng.HasFormula):
            raise ExcelWriterError(
                f"TARGET_FORMULA_CONFLICT: {sheet}!{cell} contiene fórmula"
            )
        try:
            rng.Value2 = value
        except Exception as exc:
            raise ExcelWriterError(
                f"no se pudo escribir {sheet}!{cell}: {describe_com_error(exc)}"
            ) from exc

    def recalculate(self) -> None:
        self._session.calculate_full()

    def save(self) -> None:
        # El archivo ya tiene extensión .xlsm y se copió del template: Save()
        # preserva formato y VBA.
        self._wb.Save()

    def snapshot(self) -> WorkbookSnapshot:  # pragma: no cover - la ruta COM verifica con openpyxl
        raise ExcelWriterError(
            "ExcelComWorkbookWriter no produce snapshots; la verificación reabre "
            "el archivo guardado con openpyxl (inspección estática)."
        )

    def close(self, *, save_changes: bool = False) -> None:
        wb = self._wb
        self._wb = None
        self._sheet_cache = {}
        if wb is not None:
            with contextlib.suppress(Exception):  # Close defensivo
                wb.Close(SaveChanges=save_changes)
        # Soltar los proxies del workbook/hojas antes de que la sesión desmonte
        # el apartment.
        wb = None
        gc.collect()


@contextmanager
def open_excel_com_writer(path: str | Path) -> Iterator[ExcelComWorkbookWriter]:
    """``open_writer`` para :func:`remasep.services.excel_writer.generate` en Windows.

    Crea una instancia aislada de Excel, abre la copia de trabajo, y al salir
    cierra **sólo** ese workbook y esa instancia, soltando todas las referencias
    COM antes de ``CoUninitialize``.
    """
    if platform.system() != "Windows":  # pragma: no cover - guard
        raise ExcelWriterError("Excel COM sólo está disponible en Windows.")
    session = ExcelSession(visible=False)
    session.__enter__()
    writer: ExcelComWorkbookWriter | None = None
    try:
        writer = ExcelComWorkbookWriter(session, path)
        yield writer
    finally:
        if writer is not None:
            writer.close(save_changes=False)
        writer = None
        gc.collect()
        session.__exit__(None, None, None)
