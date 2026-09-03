"""Adaptador de lectura de un export Medinet.

Responsabilidades:

- verificar que el archivo existe y es ``.xlsx``;
- localizar la hoja de datos por sus **encabezados** (no por letra de columna);
- **normalizar los encabezados** a nombres semánticos
  (``DIA_CITA``, ``FECHA_NACIMIENTO``, ...);
- devolver un ``DataFrame`` con solo las columnas reconocidas, **preservando el
  valor textual de cada celda tal cual llega de Excel** (sin strip/collapse), para
  que la clasificación legacy use exactamente el mismo texto que el workbook.
  ``None``/``NaN`` se representan como ``""``; las fechas se parsean a ``datetime``.

Solo se normaliza el ENCABEZADO, nunca el valor de la celda de texto.

NO clasifica, NO valida registros y NO calcula edades: eso vive en
``remasep.services.medinet_analysis``.

Privacidad: solo se conservan las columnas semánticas reconocidas; cualquier otra
columna del export (RUN, nombre, teléfono, ...) se descarta al leer.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from remasep.core.errors import SourceValidationError
from remasep.core.text import normalize_field_name

# Nombres semánticos con los que trabaja el core. Nunca D/G/H/K/M/O/AA.
REQUIRED_FIELDS = (
    "DIA_CITA",
    "FECHA_NACIMIENTO",
    "SEXO",
    "SUCURSAL",
    "ESPECIALIDAD",
    "TIPO_DE_CITA",
    "PRESTACION",
)
OPTIONAL_FIELDS = (
    "ESTADO",
    "MODALIDAD",
    "PRESTACION_REALIZADA",
)
DATE_FIELDS = ("DIA_CITA", "FECHA_NACIMIENTO")

# Alias de encabezado -> nombre semántico. Solo alias documentados y testeados;
# la clave ya está pasada por normalize_field_name (MAYÚSCULAS, sin tildes, "_").
_HEADER_ALIASES: dict[str, str] = {
    "DIA_CITA": "DIA_CITA",
    "FECHA_CITA": "DIA_CITA",
    "FECHA_DE_CITA": "DIA_CITA",
    "FECHA_ATENCION": "DIA_CITA",
    "FECHA_NACIMIENTO": "FECHA_NACIMIENTO",
    "FECHA_DE_NACIMIENTO": "FECHA_NACIMIENTO",
    "SEXO": "SEXO",
    "SUCURSAL": "SUCURSAL",
    "ESPECIALIDAD": "ESPECIALIDAD",
    "TIPO_DE_CITA": "TIPO_DE_CITA",
    "TIPO_CITA": "TIPO_DE_CITA",
    "ESTADO": "ESTADO",
    "ESTADO_CITA": "ESTADO",
    "MODALIDAD": "MODALIDAD",
    "PRESTACION": "PRESTACION",
    "PRESTACION_REALIZADA": "PRESTACION_REALIZADA",
}


@dataclass
class MedinetFrame:
    frame: pd.DataFrame
    sheet_name: str
    detected_fields: list[str]
    missing_optional_fields: list[str]
    blank_masks: dict[str, pd.Series]
    notes: list[str] = field(default_factory=list)

    @property
    def record_count(self) -> int:
        return len(self.frame)


def _score_headers(columns: list[object]) -> tuple[int, dict[str, str]]:
    """Cuántos campos semánticos reconoce la fila de encabezados, y el mapa."""
    mapping: dict[str, str] = {}
    for raw in columns:
        semantic = _HEADER_ALIASES.get(normalize_field_name(raw))
        if semantic and semantic not in mapping.values():
            mapping[str(raw)] = semantic
    score = len({v for v in mapping.values() if v in REQUIRED_FIELDS})
    return score, mapping


def semantic_column_map(headers) -> dict[str, str]:
    """`{str(encabezado): nombre_semántico}` para los encabezados reconocidos."""
    _score, mapping = _score_headers(list(headers))
    return mapping


def is_blank_cell(value: object) -> bool:
    """Definición escalar de "celda vacía" (equivalente a los `blank_masks` del adapter)."""
    if value is None:
        return True
    text = str(value).strip()
    return text == "" or text in {"NaT", "nan", "None", "<NA>"}


def _locate_data_sheet(excel: pd.ExcelFile) -> tuple[str, dict[str, str]]:
    best: tuple[int, str, dict[str, str]] | None = None
    for name in excel.sheet_names:
        header = pd.read_excel(excel, sheet_name=name, nrows=0)
        score, mapping = _score_headers(list(header.columns))
        if best is None or score > best[0]:
            best = (score, name, mapping)
    assert best is not None  # siempre hay al menos una hoja
    return best[1], best[2]


def read_medinet(path: str | Path, *, sheet_name: str | int | None = None) -> MedinetFrame:
    source = Path(path)
    if not source.exists():
        raise SourceValidationError(f"El archivo Medinet no existe: {source}")
    if not source.is_file():
        raise SourceValidationError(f"La ruta Medinet no es un archivo: {source}")
    if source.suffix.lower() != ".xlsx":
        raise SourceValidationError(
            f"Formato no soportado ('{source.suffix}'). Por ahora solo se acepta .xlsx."
        )

    try:
        excel = pd.ExcelFile(source)
    except (ValueError, OSError, KeyError, ImportError) as exc:
        raise SourceValidationError(f"No se pudo abrir el archivo Medinet: {exc}") from exc

    notes: list[str] = []
    with excel:
        if sheet_name is not None:
            resolved_sheet = (
                excel.sheet_names[sheet_name]
                if isinstance(sheet_name, int)
                else str(sheet_name)
            )
            header = pd.read_excel(excel, sheet_name=resolved_sheet, nrows=0)
            _score, rename_map = _score_headers(list(header.columns))
        else:
            resolved_sheet, rename_map = _locate_data_sheet(excel)
            if len(excel.sheet_names) > 1:
                notes.append(f"Hoja de datos detectada por encabezados: {resolved_sheet!r}")

        raw = pd.read_excel(excel, sheet_name=resolved_sheet)

    frame = raw.rename(columns=rename_map)
    detected = [f for f in (*REQUIRED_FIELDS, *OPTIONAL_FIELDS) if f in frame.columns]
    frame = frame[detected].copy()

    missing_required = [f for f in REQUIRED_FIELDS if f not in detected]
    if missing_required:
        raise SourceValidationError(
            "El archivo Medinet no contiene columnas para los campos requeridos: "
            + ", ".join(missing_required)
            + f". Hoja analizada: {resolved_sheet!r}."
        )

    missing_optional = [f for f in OPTIONAL_FIELDS if f not in detected]

    blank_masks: dict[str, pd.Series] = {}
    for column in detected:
        original = frame[column]
        # Detección de "vacío" (para validación y structural_empty_rows): incluye
        # celdas solo-whitespace. SOLO detección: no muta el valor almacenado.
        blank = original.isna() | (
            original.astype("string").str.strip().isin(["", "NaT", "nan", "None", "<NA>"])
        )
        blank_masks[column] = blank.fillna(True).astype(bool)

        if column in DATE_FIELDS:
            frame[column] = pd.to_datetime(original, errors="coerce", dayfirst=True)
        else:
            # Se PRESERVA el whitespace del valor de texto (semántica legacy
            # exacta): el core clasifica/concatena con el mismo texto que usaría
            # el workbook Excel. None/NaN -> "".
            frame[column] = original.astype("string").fillna("")

    frame = frame.reset_index(drop=True)
    for mask in blank_masks.values():
        mask.index = frame.index

    return MedinetFrame(
        frame=frame,
        sheet_name=resolved_sheet,
        detected_fields=detected,
        missing_optional_fields=missing_optional,
        blank_masks=blank_masks,
        notes=notes,
    )
