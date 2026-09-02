from pathlib import Path

import yaml

from remasep.domain.models import Rule


def load_rules(path: str | Path) -> list[Rule]:
    with Path(path).open("r", encoding="utf-8") as file:
        payload = yaml.safe_load(file) or {}

    return [
        Rule(
            id=raw["id"],
            description=raw.get("description", ""),
            priority=int(raw.get("priority", 0)),
            conditions=raw["conditions"],
            result=raw["result"],
        )
        for raw in payload.get("rules", [])
    ]
