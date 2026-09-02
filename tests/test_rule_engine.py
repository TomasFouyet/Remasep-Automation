import pytest

from remasep.core.errors import AmbiguousClassificationError, UnknownClassificationError
from remasep.core.rules import RuleEngine
from remasep.domain.models import Rule


def test_single_match_is_valid():
    engine = RuleEngine(
        [
            Rule(
                id="genetica",
                conditions={"ESPECIALIDAD": {"eq": "Genética"}},
                result={"categoria": "CONSULTA_MEDICA"},
            )
        ]
    )

    result = engine.classify({"ESPECIALIDAD": "GENETICA"})

    assert result.rule_id == "genetica"
    assert result.dimensions["categoria"] == "CONSULTA_MEDICA"


def test_unknown_is_blocking():
    engine = RuleEngine(
        [
            Rule(
                id="genetica",
                conditions={"ESPECIALIDAD": "GENETICA"},
                result={"categoria": "CONSULTA_MEDICA"},
            )
        ]
    )

    with pytest.raises(UnknownClassificationError):
        engine.classify({"ESPECIALIDAD": "OTORRINO"})


def test_ambiguous_is_blocking():
    engine = RuleEngine(
        [
            Rule(
                id="general",
                conditions={"ESPECIALIDAD": "GENETICA"},
                result={"categoria": "A"},
            ),
            Rule(
                id="specific",
                conditions={"ESPECIALIDAD": {"contains": "GENETICA"}},
                result={"categoria": "B"},
            ),
        ]
    )

    with pytest.raises(AmbiguousClassificationError):
        engine.classify({"ESPECIALIDAD": "GENETICA"})
