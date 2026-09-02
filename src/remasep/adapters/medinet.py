from pathlib import Path

import pandas as pd

from remasep.core.errors import SourceValidationError
from remasep.core.text import normalize_field_name


def read_medinet(
    path: str | Path,
    *,
    required_columns: list[str],
    sheet_name: str | int = 0,
) -> pd.DataFrame:
    """
    Adaptador preliminar. Se ajustará con un export mensual real de Medinet.
    """
    frame = pd.read_excel(Path(path), sheet_name=sheet_name)
    frame = frame.rename(
        columns={column: normalize_field_name(column) for column in frame.columns}
    )

    required = [normalize_field_name(column) for column in required_columns]
    missing = [column for column in required if column not in frame.columns]

    if missing:
        raise SourceValidationError(
            "El archivo Medinet no contiene las columnas requeridas: "
            + ", ".join(missing)
        )

    return frame
