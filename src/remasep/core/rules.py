import re
from typing import Any

from remasep.core.errors import AmbiguousClassificationError, UnknownClassificationError
from remasep.core.text import normalize_text
from remasep.domain.models import Classification, Rule


def _matches_condition(actual: Any, specification: Any) -> bool:
    actual_norm = normalize_text(actual)

    if not isinstance(specification, dict):
        return actual_norm == normalize_text(specification)

    if "eq" in specification:
        return actual_norm == normalize_text(specification["eq"])

    if "in" in specification:
        allowed = {normalize_text(value) for value in specification["in"]}
        return actual_norm in allowed

    if "contains" in specification:
        return normalize_text(specification["contains"]) in actual_norm

    if "starts_with" in specification:
        return actual_norm.startswith(normalize_text(specification["starts_with"]))

    if "regex" in specification:
        return re.search(specification["regex"], actual_norm) is not None

    raise ValueError(f"Operador de condición no soportado: {specification}")


def rule_matches(row: dict[str, Any], rule: Rule) -> bool:
    return all(
        _matches_condition(row.get(field), specification)
        for field, specification in rule.conditions.items()
    )


class RuleEngine:
    """
    Motor deliberadamente estricto:
    0 coincidencias -> error
    1 coincidencia  -> OK
    >1 coincidencia -> error
    """

    def __init__(self, rules: list[Rule]):
        self.rules = rules

    def classify(
        self,
        row: dict[str, Any],
        *,
        row_reference: str | None = None,
    ) -> Classification:
        matches = [rule for rule in self.rules if rule_matches(row, rule)]

        if not matches:
            raise UnknownClassificationError(row_reference)

        if len(matches) > 1:
            raise AmbiguousClassificationError(
                [rule.id for rule in matches],
                row_reference=row_reference,
            )

        rule = matches[0]
        return Classification(rule_id=rule.id, dimensions=dict(rule.result))
