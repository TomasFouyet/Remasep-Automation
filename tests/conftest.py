"""Configuración compartida de tests.

Fuerza Qt en modo offscreen para que la maqueta de UI pueda testearse sin
display físico, y expone un constructor de workbooks Medinet sintéticos.
"""

import os
import re
import shutil
import zipfile

import pytest
from openpyxl import Workbook

os.environ.setdefault("QT_QPA_PLATFORM", "offscreen")

# Encabezados equivalentes a los del workbook real (con tildes y espacios).
MEDINET_HEADERS = [
    "DIA CITA",
    "FECHA NACIMIENTO",
    "SEXO",
    "SUCURSAL",
    "ESPECIALIDAD",
    "TIPO DE CITA",
    "PRESTACIÓN",
    "ESTADO",
    "MODALIDAD",
    "PRESTACIÓN REALIZADA",
]


@pytest.fixture(scope="session")
def qapp():
    from PySide6.QtWidgets import QApplication

    app = QApplication.instance() or QApplication([])
    yield app


def _xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


@pytest.fixture
def inject_formula_cache():
    """Inyecta valores *cacheados* (`<v>`) en celdas con fórmula de un `.xlsx`.

    openpyxl no escribe valores cacheados; los comparadores leen `data_only=True`,
    así que los tests parchean el XML de la hoja tras guardar.
    ``cache`` = ``{(fila_int, columna_letra): valor}``.
    """

    def _inject(path, cache, *, sheet_xml="xl/worksheets/sheet1.xml"):
        with zipfile.ZipFile(path) as zf:
            names = zf.namelist()
            xml = zf.read(sheet_xml).decode("utf-8")

        for (row, col), value in cache.items():
            ref = f"{col}{row}"
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                replacement = rf'<c r="{ref}"\g<attrs>><f>\g<f></f><v>{value}</v></c>'
            else:
                replacement = (
                    rf'<c r="{ref}"\g<attrs> t="str"><f>\g<f></f>'
                    rf'<v>{_xml_escape(str(value))}</v></c>'
                )
            pattern = re.compile(
                rf'<c r="{ref}"(?P<attrs>[^>]*)><f>(?P<f>[^<]*)</f>(?:<v\s*/>|<v></v>)</c>'
            )
            xml, n = pattern.subn(replacement, xml)
            assert n == 1, f"no se pudo parchear {ref}"

        tmp = str(path) + ".tmp"
        with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
            for name in names:
                data = xml.encode("utf-8") if name == sheet_xml else src.read(name)
                dst.writestr(name, data)
        shutil.move(tmp, path)

    return _inject


@pytest.fixture
def make_medinet(tmp_path):
    """Devuelve un builder de archivos Medinet sintéticos (.xlsx).

    ``rows`` es una lista de listas alineadas a ``headers`` (por defecto
    :data:`MEDINET_HEADERS`). No usa data/local.
    """

    def _build(
        rows,
        *,
        headers=None,
        sheet="Atenciones",
        filename="medinet.xlsx",
        decoy_sheets=None,
        extra_columns=None,
    ):
        wb = Workbook()
        first = wb.active
        if decoy_sheets:
            (name0, rows0), *rest = decoy_sheets
            first.title = name0
            for decoy_row in rows0:
                first.append(list(decoy_row))
            for name, drows in rest:
                extra = wb.create_sheet(name)
                for decoy_row in drows:
                    extra.append(list(decoy_row))
            data_ws = wb.create_sheet(sheet)
        else:
            first.title = sheet
            data_ws = first

        # extra_columns: [(header, value_para_todas_las_filas), ...] — p.ej. una
        # columna derivada "arrastrada" que evita que openpyxl recorte las filas
        # semánticamente vacías al final.
        extra_columns = list(extra_columns or [])
        data_ws.append(list(headers or MEDINET_HEADERS) + [name for name, _ in extra_columns])
        tail = [value for _, value in extra_columns]
        for row in rows:
            data_ws.append(list(row) + tail)
        path = tmp_path / filename
        wb.save(path)
        return path

    return _build
