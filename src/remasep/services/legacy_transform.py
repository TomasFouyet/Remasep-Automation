"""Transformaciones derivadas legacy AC:AL — reproducen `GENERACION DATOS REMASEP.xlsx`.

Módulo de **producción reutilizable**: lo consumen tanto el análisis Medinet
(`legacy_age_years`) como el comparador de equivalencia
(`scripts/compare_legacy_derived.py`). El objetivo es reproducir el comportamiento
**exacto** del workbook, no "mejorarlo":

- `AC = TIPO_DE_CITA & SUCURSAL` (concatenación sin separador; celda vacía -> "").
- `AD = TIPO_DE_CITA & SEXO`.
- `AE = PRESTACION & ESPECIALIDAD & SEXO`.
- `AF = IF(G>D, 0, DATEDIF(G, D, "Y"))` — años completos, con la excepción legacy
  `nacimiento > atención -> 0`.
- `AG/AH/AI = IFS(TIPO_DE_CITA = literal, salida, ...) & SEXO` — igualdad de texto
  case-insensitive (como Excel); sin match, el IFS devuelve `#N/A` y el `&` lo
  propaga tal cual (valor de celda `"#N/A"`).
- `AJ/AK/AL = (Σ COUNTIF(PRESTACION, "*patrón*")) & SEXO` — nº de patrones
  presentes (substring case-insensitive, con tildes), concatenado con el sexo.

La equivalencia de este módulo es contra el **workbook legacy**, no contra reglas
oficiales MINSAL.
"""

from __future__ import annotations

from collections.abc import Mapping

import pandas as pd

from remasep.core.text import normalize_legacy_text
from remasep.services.legacy_rules import LegacyCategory, LegacyRuleSet

LEGACY_NA = "#N/A"
DERIVED_COLUMNS: tuple[str, ...] = ("AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL")

# Campos fuente de cada concatenación (en el orden en que Excel los une).
_CONCAT_FIELDS: dict[str, tuple[str, ...]] = {
    "AC": ("TIPO_DE_CITA", "SUCURSAL"),
    "AD": ("TIPO_DE_CITA", "SEXO"),
    "AE": ("PRESTACION", "ESPECIALIDAD", "SEXO"),
}


def _is_missing(value: object) -> bool:
    if value is None:
        return True
    try:
        return bool(pd.isna(value))
    except (TypeError, ValueError):
        return False


def excel_str(value: object) -> str:
    """Coerción a texto al estilo del operador `&` de Excel."""
    if _is_missing(value):
        return ""
    if isinstance(value, bool):
        return "TRUE" if value else "FALSE"
    if isinstance(value, float):
        return str(int(value)) if value.is_integer() else repr(value)
    if isinstance(value, int):
        return str(value)
    if isinstance(value, str):
        return value
    return str(value)


def _to_timestamp(value: object) -> pd.Timestamp | None:
    if _is_missing(value):
        return None
    if isinstance(value, str):
        # Excel (locale es-CL) interpreta el texto como dd/mm/aaaa.
        parsed = pd.to_datetime(value, dayfirst=True, errors="coerce")
        return None if pd.isna(parsed) else pd.Timestamp(parsed)
    return pd.Timestamp(value)


def legacy_age_years(birth: object, service: object) -> int | None:
    """Años completos cumplidos a la fecha de atención (equivale a ``DATEDIF(b, s, "Y")``).

    Excepción de compatibilidad legacy: si ``birth > service`` la fórmula actual del
    workbook devuelve ``0``. Devuelve ``None`` si falta alguna de las dos fechas.
    """
    birth_ts = _to_timestamp(birth)
    service_ts = _to_timestamp(service)
    if birth_ts is None or service_ts is None:
        return None
    if birth_ts > service_ts:
        return 0
    years = service_ts.year - birth_ts.year
    if (service_ts.month, service_ts.day) < (birth_ts.month, birth_ts.day):
        years -= 1
    return int(years)


def legacy_concat(rec: Mapping[str, object], fields: tuple[str, ...]) -> str:
    return "".join(excel_str(rec.get(field)) for field in fields)


def legacy_exact_map(rec: Mapping[str, object], category: LegacyCategory) -> str:
    """`IFS(campo = literal, salida, ...) & SEXO`; sin match -> ``#N/A`` (como Excel).

    Igualdad de texto legacy: case-insensitive, **con tildes** (`normalize_legacy_text`).
    """
    target = normalize_legacy_text(rec.get(category.field))
    for value in category.normalized_values:
        if target == value:
            return category.output_literal + excel_str(rec.get("SEXO"))
    return LEGACY_NA


def legacy_pattern_flag(rec: Mapping[str, object], category: LegacyCategory) -> str:
    """`(Σ COUNTIF(campo, "*patrón*")) & SEXO` — nº de patrones presentes + sexo.

    `contains` legacy: substring case-insensitive, **con tildes** (`normalize_legacy_text`).
    """
    haystack = normalize_legacy_text(rec.get(category.field))
    count = sum(1 for value in category.normalized_values if value and value in haystack)
    return f"{count}" + excel_str(rec.get("SEXO"))


def legacy_derived_value(column: str, rec: Mapping[str, object], ruleset: LegacyRuleSet) -> object:
    """Valor legacy exacto de una columna derivada (`AC`..`AL`) para un registro raw."""
    if column in _CONCAT_FIELDS:
        return legacy_concat(rec, _CONCAT_FIELDS[column])
    if column == "AF":
        return legacy_age_years(rec.get("FECHA_NACIMIENTO"), rec.get("DIA_CITA"))
    category = ruleset.category(column)
    if category.operator == "equals":
        return legacy_exact_map(rec, category)
    return legacy_pattern_flag(rec, category)


def legacy_derived_record(rec: Mapping[str, object], ruleset: LegacyRuleSet) -> dict[str, object]:
    return {column: legacy_derived_value(column, rec, ruleset) for column in DERIVED_COLUMNS}
