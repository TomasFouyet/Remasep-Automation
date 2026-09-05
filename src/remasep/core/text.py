import re
import unicodedata


def normalize_text(value: object) -> str:
    if value is None:
        return ""

    text = str(value).strip().upper()
    text = " ".join(text.split())
    text = unicodedata.normalize("NFD", text)
    return "".join(char for char in text if unicodedata.category(char) != "Mn")


def normalize_field_name(value: object) -> str:
    text = normalize_text(value)
    return re.sub(r"[^A-Z0-9]+", "_", text).strip("_")


def normalize_legacy_text(value: object) -> str:
    """Normalización que reproduce las comparaciones de texto del workbook Excel legacy.

    Objetivo: **equivalencia legacy exacta**, no matching tolerante.

    - conversión segura a string (``None`` -> ``""``);
    - **case-insensitive** (``casefold``);
    - **preserva tildes y diacríticos** (a diferencia de :func:`normalize_text`);
    - **preserva el whitespace tal cual** — Excel (`=` y `COUNTIF "*x*"`) no
      recorta ni colapsa espacios, así que aquí tampoco.

    Es la ÚNICA semántica para ``equals`` / ``contains`` cuando se reproduce la
    lógica legacy (``LegacyRuleSet`` y ``legacy_transform``). No modifica
    :func:`normalize_text`.
    """
    if value is None:
        return ""
    return str(value).casefold()
