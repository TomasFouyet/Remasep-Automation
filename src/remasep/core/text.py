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
