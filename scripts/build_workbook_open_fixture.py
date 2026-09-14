"""Sprint C (F06) — construye el fixture sintético de ``Workbook_Open`` (dev-only,
Windows + Excel, una sola vez).

Genera ``tests/fixtures/synthetic_workbook_open.xlsm``: un workbook benigno,
sin datos clínicos ni PII, cuyo único código es::

    Private Sub Workbook_Open()
        Sheets(1).Range("A1").Value = "MARKER_FIRED"
    End Sub

Este archivo se COMMITEA al repo (es sintético y benigno, a diferencia de las
plantillas reales en ``data/local/`` que están gitignored) para que el test
de integración F06 (``tests/test_excel_com_integration.py::
test_workbook_open_macro_does_not_fire_through_our_adapter``) no dependa, en
cada corrida, de que el Excel de esa máquina tenga habilitado "Trust access
to the VBA project object model" — sólo ESTE script, ejecutado una vez por un
desarrollador, necesita esa opción (Trust Center > Macro Settings), y sólo
para poder escribir el módulo VBA por código. REMASEP en producción NUNCA
necesita ni cambia esa opción.

Si el fixture ya existe y la plantilla no cambió, no hace falta regenerarlo.

Uso (Windows, con Excel instalado, y con esa opción de Trust Center
habilitada SÓLO durante esta ejecución)::

    python scripts/build_workbook_open_fixture.py
"""

from __future__ import annotations

import sys
from pathlib import Path

DEFAULT_OUT = Path("tests/fixtures/synthetic_workbook_open.xlsm")
MARKER_CELL = "A1"
MARKER_VALUE = "MARKER_FIRED"


def build(out_path: Path) -> int:
    import platform

    if platform.system() != "Windows":
        print("ABORTADO: este builder requiere Windows + Excel real.", file=sys.stderr)
        return 1

    import pythoncom
    import win32com.client

    out_path.parent.mkdir(parents=True, exist_ok=True)

    pythoncom.CoInitialize()
    app = win32com.client.DispatchEx("Excel.Application")
    app.Visible = False
    app.DisplayAlerts = False
    try:
        wb = app.Workbooks.Add()
        try:
            vb_project = wb.VBProject
        except Exception as exc:  # noqa: BLE001 - se traduce a un mensaje accionable
            wb.Close(SaveChanges=False)
            print(
                "ABORTADO: no se pudo acceder a VBProject. Habilita "
                "'Trust access to the VBA project object model' en "
                "Archivo > Opciones > Centro de confianza > Configuración "
                "del Centro de confianza > Configuración de macros, "
                "SÓLO mientras corres este script una vez, y vuelve a "
                f"intentar. Detalle: {exc}",
                file=sys.stderr,
            )
            return 1

        this_workbook = next(c for c in vb_project.VBComponents if c.Name == "ThisWorkbook")
        this_workbook.CodeModule.AddFromString(
            "Private Sub Workbook_Open()\n"
            f'    Sheets(1).Range("{MARKER_CELL}").Value = "{MARKER_VALUE}"\n'
            "End Sub\n"
        )
        if out_path.exists():
            out_path.unlink()
        wb.SaveAs(str(out_path.resolve()), FileFormat=52)  # xlOpenXMLWorkbookMacroEnabled
        wb.Close(SaveChanges=False)
    finally:
        app.Quit()
        pythoncom.CoUninitialize()

    print(f"OK: {out_path} (marcador esperado {MARKER_VALUE!r} en {MARKER_CELL})")
    return 0


def main(argv: list[str] | None = None) -> int:
    import argparse

    parser = argparse.ArgumentParser(prog="build_workbook_open_fixture.py")
    parser.add_argument("--out", default=str(DEFAULT_OUT))
    args = parser.parse_args(argv)
    return build(Path(args.out))


if __name__ == "__main__":
    raise SystemExit(main())
