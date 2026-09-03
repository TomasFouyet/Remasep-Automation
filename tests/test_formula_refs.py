"""Tests del analizador de referencias (Sprint 1.2, Parte A).

Todo sintético: cadenas de fórmula escritas a mano, sin workbooks ni datos reales.
"""

import formula_refs as fr


def _refs(formula):
    """(sheet, ref_type, col_start, col_end, row_start, row_end) por referencia."""
    return [
        (r.sheet, r.ref_type, r.col_start, r.col_end, r.row_start, r.row_end)
        for r in fr.parse_references(formula).references
    ]


def test_local_cell():
    assert _refs("=A1+1") == [(None, fr.CELL, "A", "A", 1, 1)]


def test_absolute_cell():
    assert _refs("=$A$1") == [(None, fr.CELL, "A", "A", 1, 1)]
    # Mixta y con minúsculas normalizadas.
    assert _refs("=a$1") == [(None, fr.CELL, "A", "A", 1, 1)]


def test_range():
    assert _refs("=SUM(A1:B10)") == [(None, fr.RANGE, "A", "B", 1, 10)]
    # Rango escrito al revés se normaliza a start<=end.
    assert _refs("=SUM(B10:A1)") == [(None, fr.RANGE, "A", "B", 1, 10)]


def test_whole_column():
    assert _refs("=SUM(A:A)") == [(None, fr.WHOLE_COLUMN, "A", "A", None, None)]
    assert _refs("=SUM($A:$C)") == [(None, fr.WHOLE_COLUMN, "A", "C", None, None)]


def test_whole_row():
    assert _refs("=SUM(1:1)") == [(None, fr.WHOLE_ROW, None, None, 1, 1)]
    assert _refs("=SUM(2:5)") == [(None, fr.WHOLE_ROW, None, None, 2, 5)]


def test_cross_sheet_simple():
    assert _refs("=Sheet2!A1") == [("Sheet2", fr.CELL, "A", "A", 1, 1)]
    assert _refs("=Sheet2!A:A") == [("Sheet2", fr.WHOLE_COLUMN, "A", "A", None, None)]


def test_cross_sheet_quoted_name_with_spaces():
    formula = "=COUNTIF('Atenciones - Detalles de citas'!$AF:$AF,\">10\")"
    assert _refs(formula) == [
        ("Atenciones - Detalles de citas", fr.WHOLE_COLUMN, "AF", "AF", None, None)
    ]


def test_quoted_sheet_name_with_escaped_quote():
    assert _refs("='Bob''s data'!A1") == [("Bob's data", fr.CELL, "A", "A", 1, 1)]


def test_strings_that_look_like_references_are_ignored():
    # "A1:B2" y "Sheet2!C3" viven dentro de literales -> no son referencias.
    formula = '=IF(A1="A1:B2","Sheet2!C3",B2)'
    assert _refs(formula) == [
        (None, fr.CELL, "A", "A", 1, 1),
        (None, fr.CELL, "B", "B", 2, 2),
    ]


def test_multiple_references_in_one_formula():
    assert _refs("=A1+B2+C3") == [
        (None, fr.CELL, "A", "A", 1, 1),
        (None, fr.CELL, "B", "B", 2, 2),
        (None, fr.CELL, "C", "C", 3, 3),
    ]


def test_repeated_references_are_deduplicated():
    assert _refs("=A1+A1+A1") == [(None, fr.CELL, "A", "A", 1, 1)]
    assert _refs("=SUM(A1:A9)+AVERAGE(A1:A9)") == [(None, fr.RANGE, "A", "A", 1, 9)]


def test_nested_functions():
    formula = '=IF(DATEDIF(D2,G2,"Y")>0,SUM(H2:H10),0)'
    assert _refs(formula) == [
        (None, fr.CELL, "D", "D", 2, 2),
        (None, fr.CELL, "G", "G", 2, 2),
        (None, fr.RANGE, "H", "H", 2, 10),
    ]


def test_function_names_are_not_mistaken_for_cells():
    # LOG10( parece la celda LOG10 pero va seguido de '(' -> no es referencia.
    assert _refs("=LOG10(A1)") == [(None, fr.CELL, "A", "A", 1, 1)]
    assert _refs("=SUM(A1)") == [(None, fr.CELL, "A", "A", 1, 1)]


def test_a1a1_collapses_to_cell():
    assert _refs("=A1:A1") == [(None, fr.CELL, "A", "A", 1, 1)]


def test_raw_is_preserved():
    refs = fr.parse_references("=COUNTIF('X Y'!$A:$A,1)").references
    assert refs[0].raw == "'X Y'!$A:$A"


# --- constructs no soportados ---------------------------------------------


def test_unsupported_indirect():
    result = fr.parse_references('=INDIRECT("A"&B1)')
    assert "indirect" in result.unsupported
    # El argumento literal B1 sí se detecta como referencia; el rango es dinámico.
    assert (None, fr.CELL, "B", "B", 1, 1) in _to_tuples(result.references)


def test_unsupported_offset():
    assert "offset" in fr.parse_references("=OFFSET(A1,1,0)").unsupported


def test_unsupported_external_reference():
    assert "external_reference" in fr.parse_references("=[Book1.xlsx]Sheet1!A1").unsupported
    assert "external_reference" in fr.parse_references("='[Book1.xlsx]Sheet1'!A1").unsupported


def test_unsupported_structured_reference():
    assert "structured_reference" in fr.parse_references("=Tabla1[Importe]").unsupported
    assert "structured_reference" in fr.parse_references("=SUM(Tabla1[@[Neto]])").unsupported


def test_unsupported_three_d_reference():
    assert "three_d_reference" in fr.parse_references("=SUM(Hoja1:Hoja3!A1)").unsupported


def test_unsupported_ref_error():
    assert "ref_error" in fr.parse_references("=#REF!+1").unsupported


def test_clean_formula_has_no_unsupported():
    assert fr.parse_references('=COUNTIFS(A:A,"x",B:B,"y")').unsupported == []


# --- utilidades de columna ----------------------------------------------------


def test_column_index_roundtrip():
    for col in ("A", "Z", "AA", "AL", "CM", "XFD"):
        assert fr.index_to_col(fr.col_to_index(col)) == col


def _to_tuples(references):
    return [
        (r.sheet, r.ref_type, r.col_start, r.col_end, r.row_start, r.row_end)
        for r in references
    ]
