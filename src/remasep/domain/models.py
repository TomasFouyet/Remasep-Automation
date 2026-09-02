from dataclasses import dataclass, field
from typing import Any


@dataclass(frozen=True)
class Rule:
    id: str
    conditions: dict[str, Any]
    result: dict[str, Any]
    description: str = ""
    priority: int = 0


@dataclass(frozen=True)
class Classification:
    rule_id: str
    dimensions: dict[str, Any]


@dataclass(frozen=True)
class Metric:
    formulario: str
    categoria: str
    valor: int | float
    dimensions: dict[str, Any] = field(default_factory=dict)

    def as_record(self) -> dict[str, Any]:
        return {
            "formulario": self.formulario,
            "categoria": self.categoria,
            **self.dimensions,
            "valor": self.valor,
        }
