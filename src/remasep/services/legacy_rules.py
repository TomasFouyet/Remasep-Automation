"""Reglas legacy de clasificación (compatibilidad con el workbook generador).

Estas reglas reproducen la lógica observada en ``GENERACION DATOS REMASEP.xlsx``
(columnas AG:AL) y están **pendientes de validación funcional**. No son reglas
oficiales MINSAL. Se cargan desde ``config/legacy_current_logic_2026/rules.yaml``.

El matching usa :func:`remasep.core.text.normalize_text` (mayúsculas, sin tildes,
espacios colapsados). No hay fuzzy matching.
"""

from __future__ import annotations

from collections.abc import Mapping
from dataclasses import dataclass
from pathlib import Path

import yaml

from remasep.core.errors import RemasepError
from remasep.core.text import normalize_text

DEFAULT_LEGACY_RULES_PATH = (
    Path(__file__).resolve().parents[3]
    / "config"
    / "legacy_current_logic_2026"
    / "rules.yaml"
)

_VALID_OPERATORS = {"equals", "contains"}
_VALID_FIELDS = {"TIPO_DE_CITA", "PRESTACION"}


class LegacyRulesError(RemasepError):
    """Configuración de reglas legacy inválida."""


@dataclass(frozen=True)
class LegacyCategory:
    code: str
    label: str
    field: str
    operator: str
    values: tuple[str, ...]

    def __post_init__(self) -> None:
        object.__setattr__(
            self, "_normalized", tuple(normalize_text(value) for value in self.values)
        )

    def matches(self, row: Mapping[str, object]) -> bool:
        target = normalize_text(row.get(self.field))
        if not target:
            return False
        if self.operator == "equals":
            return any(target == value for value in self._normalized)
        # contains
        return any(value and value in target for value in self._normalized)


@dataclass(frozen=True)
class LegacyRuleSet:
    version: str
    source_workbook: str
    status: str
    disclaimer: str
    categories: tuple[LegacyCategory, ...]

    @property
    def codes(self) -> tuple[str, ...]:
        return tuple(category.code for category in self.categories)

    def classify(self, row: Mapping[str, object]) -> list[str]:
        """Códigos de categoría legacy que activa el registro (0, 1 o más)."""
        return [category.code for category in self.categories if category.matches(row)]


def load_legacy_rules(path: str | Path | None = None) -> LegacyRuleSet:
    rules_path = Path(path) if path is not None else DEFAULT_LEGACY_RULES_PATH
    if not rules_path.is_file():
        raise LegacyRulesError(f"No se encontró el archivo de reglas legacy: {rules_path}")

    try:
        payload = yaml.safe_load(rules_path.read_text(encoding="utf-8")) or {}
    except yaml.YAMLError as exc:  # pragma: no cover - error de formato
        raise LegacyRulesError(f"YAML inválido en {rules_path}: {exc}") from exc

    raw_categories = payload.get("categories") or []
    if not raw_categories:
        raise LegacyRulesError(f"El archivo de reglas legacy no define categorías: {rules_path}")

    categories: list[LegacyCategory] = []
    for raw in raw_categories:
        code = str(raw["code"]).strip()
        field = str(raw["field"]).strip()
        operator = str(raw["operator"]).strip()
        values = tuple(str(value) for value in raw.get("values", []))
        if operator not in _VALID_OPERATORS:
            raise LegacyRulesError(f"Operador no soportado en categoría {code}: {operator!r}")
        if field not in _VALID_FIELDS:
            raise LegacyRulesError(f"Campo no soportado en categoría {code}: {field!r}")
        if not values:
            raise LegacyRulesError(f"La categoría {code} no tiene valores")
        categories.append(
            LegacyCategory(
                code=code,
                label=str(raw.get("label", code)),
                field=field,
                operator=operator,
                values=values,
            )
        )

    return LegacyRuleSet(
        version=str(payload.get("version", "unknown")),
        source_workbook=str(payload.get("source_workbook", "")),
        status=str(payload.get("status", "")),
        disclaimer=str(payload.get("disclaimer", "")).strip(),
        categories=tuple(categories),
    )
