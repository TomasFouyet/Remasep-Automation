from contextlib import AbstractContextManager
from pathlib import Path
import platform

from remasep.core.errors import TemplateValidationError


class ExcelSession(AbstractContextManager):
    """
    Sesión aislada de Microsoft Excel.

    Usa DispatchEx para crear una instancia propia de Excel.
    """

    def __init__(self, *, visible: bool = False):
        self.visible = visible
        self.excel = None

    def __enter__(self):
        if platform.system() != "Windows":
            raise RuntimeError("Excel COM solamente está disponible en Windows.")

        import win32com.client

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

        self.excel.Calculation = -4105  # xlCalculationAutomatic
        self.excel.CalculateFullRebuild()

    def __exit__(self, exc_type, exc, tb):
        if self.excel is not None:
            try:
                self.excel.DisplayAlerts = False
                self.excel.Quit()
            finally:
                self.excel = None
        return False


def validate_required_sheets(workbook, required_sheets: list[str]) -> None:
    existing = {sheet.Name for sheet in workbook.Worksheets}
    missing = [sheet for sheet in required_sheets if sheet not in existing]

    if missing:
        raise TemplateValidationError(
            "La plantilla no contiene hojas requeridas: " + ", ".join(missing)
        )
