"""Adaptador de Microsoft Excel Desktop vía COM (Sprint 3.7B).

**Sólo Windows.** ``win32com`` se importa de forma perezosa dentro de los
métodos: importar este módulo en Linux/WSL no falla. La escritura del REMASEP
oficial se hace exclusivamente con esta ruta (Excel COM), nunca con
``openpyxl.save()``.
"""

from __future__ import annotations

import contextlib
import platform
from collections.abc import Iterator
from contextlib import AbstractContextManager, contextmanager
from pathlib import Path

from remasep.core.errors import TemplateValidationError
from remasep.services.excel_writer import ExcelWriterError, WorkbookSnapshot

# Constantes de Excel (evitan depender de la typelib).
_XL_CALC_MANUAL = -4135
_XL_CALC_AUTOMATIC = -4105
_XL_CALC_STATE_DONE = 0
_XL_OPENXML_MACRO_ENABLED = 52


class ExcelSession(AbstractContextManager):
    """Sesión aislada de Microsoft Excel (``DispatchEx``, no instancia compartida)."""

    def __init__(self, *, visible: bool = False):
        self.visible = visible
        self.excel = None

    def __enter__(self):
        if platform.system() != "Windows":
            raise RuntimeError("Excel COM solamente está disponible en Windows.")

        import pythoncom  # type: ignore[import-not-found]
        import win32com.client  # type: ignore[import-not-found]

        pythoncom.CoInitialize()
        self.excel = win32com.client.DispatchEx("Excel.Application")
        self.excel.Visible = self.visible
        self.excel.DisplayAlerts = False
        return self

    def open_workbook(self, path: str | Path, *, read_only: bool = False):
        if self.excel is None:
            raise RuntimeError("La sesión Excel no está inicializada.")
        return self.excel.Workbooks.Open(
            str(Path(path).resolve()),
            ReadOnly=read_only,
        )

    def calculate_full(self) -> None:
        if self.excel is None:
            raise RuntimeError("La sesión Excel no está inicializada.")
        self.excel.Calculation = _XL_CALC_AUTOMATIC
        self.excel.CalculateFullRebuild()
        # Esperar a que Excel termine el cálculo antes de guardar.
        with contextlib.suppress(Exception):
            for _ in range(600):
                if int(self.excel.CalculationState) == _XL_CALC_STATE_DONE:
                    break
                self.excel.Wait(0)

    def __exit__(self, exc_type, exc, tb):
        if self.excel is not None:
            try:
                self.excel.DisplayAlerts = False
                self.excel.Quit()
            finally:
                self.excel = None
                with contextlib.suppress(Exception):
                    import pythoncom  # type: ignore[import-not-found]

                    pythoncom.CoUninitialize()
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
                f"no se pudo escribir {sheet}!{cell}: {exc.__class__.__name__}"
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
        with contextlib.suppress(Exception):
            self._wb.Close(SaveChanges=save_changes)


@contextmanager
def open_excel_com_writer(path: str | Path) -> Iterator[ExcelComWorkbookWriter]:
    """``open_writer`` para :func:`remasep.services.excel_writer.generate` en Windows.

    Crea una instancia aislada de Excel, abre la copia de trabajo, y al salir
    cierra **sólo** ese workbook y esa instancia.
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
        session.__exit__(None, None, None)
