from remasep.core.aggregation import aggregate_long


def test_aggregate_long():
    rows = [
        {
            "formulario": "REMASEP_01",
            "categoria": "CONSULTA_MEDICA",
            "sexo": "MUJER",
            "edad": "20_24",
            "valor": 1,
        },
        {
            "formulario": "REMASEP_01",
            "categoria": "CONSULTA_MEDICA",
            "sexo": "MUJER",
            "edad": "20_24",
            "valor": 1,
        },
    ]

    result = aggregate_long(
        rows,
        dimensions=["formulario", "categoria", "sexo", "edad"],
    )

    assert len(result) == 1
    assert result.iloc[0]["valor"] == 2
