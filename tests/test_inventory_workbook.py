"""Tests de scripts/inventory_workbook.py.

Todos los workbooks usados aquí se construyen sintéticamente durante el test:
nunca se leen datos reales ni datos de pacientes.
"""

import csv
import json
import zipfile

import inventory_workbook as iw
import pytest
from openpyxl import Workbook
from openpyxl.workbook.defined_name import DefinedName


@pytest.fixture
def synthetic_workbook(tmp_path):
    """Workbook con varias hojas, fórmulas repetidas, hoja oculta, merges y named range."""
    wb = Workbook()

    data = wb.active
    data.title = "Data"
    data["A1"] = "sexo"
    data["B1"] = "edad"
    data["C1"] = "monto"
    data["D1"] = "doble"
    for row in range(2, 7):
        data[f"A{row}"] = "F" if row % 2 else "M"
        data[f"B{row}"] = row * 3
        data[f"C{row}"] = row * 10
        # Fórmula repetida (misma familia tras normalizar referencias relativas).
        data[f"D{row}"] = f"=C{row}*2"
    data["C7"] = "=SUM(C2:C6)"
    data.freeze_panes = "A2"

    summary = wb.create_sheet("Summary")
    summary["A1"] = "resumen"
    summary["A2"] = '=COUNTIFS(Data!A2:A6,"F",Data!B2:B6,">=10")'
    summary["A3"] = '=COUNTIFS(Data!A2:A6,"M",Data!B2:B6,">=10")'
    summary["A4"] = '=VLOOKUP("x",Data!A2:D6,4,FALSE)'
    summary.merge_cells("B1:D1")
    summary["B1"] = "cabecera combinada"

    hidden = wb.create_sheet("HiddenSheet")
    hidden.sheet_state = "hidden"
    hidden["A1"] = "=Data!C7*1"

    wb.defined_names.add(DefinedName("montos", attr_text="Data!$C$2:$C$6"))

    path = tmp_path / "synthetic.xlsx"
    wb.save(path)
    return path


def _read_csv(path):
    with open(path, newline="", encoding="utf-8") as handle:
        return list(csv.DictReader(handle))


def test_generates_all_expected_files(synthetic_workbook, tmp_path):
    out = tmp_path / "inventory"
    assert iw.main([str(synthetic_workbook), "--output", str(out)]) == 0

    for name in (
        "workbook_summary.json",
        "sheets.csv",
        "formulas.csv",
        "formula_patterns.csv",
        "named_ranges.csv",
        "README.md",
    ):
        assert (out / name).is_file(), name


def test_summary_contents(synthetic_workbook, tmp_path):
    out = tmp_path / "inventory"
    iw.main([str(synthetic_workbook), "--output", str(out)])

    summary = json.loads((out / "workbook_summary.json").read_text(encoding="utf-8"))

    assert summary["file"]["name"] == "synthetic.xlsx"
    assert len(summary["file"]["sha256"]) == 64
    assert summary["workbook"]["sheet_count"] == 3
    assert summary["workbook"]["sheet_names"] == ["Data", "Summary", "HiddenSheet"]

    states = {s["name"]: s["state"] for s in summary["sheets"]}
    assert states["HiddenSheet"] == "hidden"
    assert states["Data"] == "visible"

    data_sheet = next(s for s in summary["sheets"] if s["name"] == "Data")
    assert data_sheet["frozen_panes"] == "A2"
    assert data_sheet["formula_cells"] == 6  # 5 x "=C{n}*2" + SUM

    summary_sheet = next(s for s in summary["sheets"] if s["name"] == "Summary")
    assert summary_sheet["merged_range_count"] == 1
    assert "B1:D1" in summary_sheet["merged_ranges"]

    assert summary["formula_cells_by_sheet"]["HiddenSheet"] == 1


def test_input_file_is_not_modified(synthetic_workbook, tmp_path):
    before = synthetic_workbook.read_bytes()
    iw.main([str(synthetic_workbook), "--output", str(tmp_path / "inventory")])
    assert synthetic_workbook.read_bytes() == before


def test_sheets_csv(synthetic_workbook, tmp_path):
    out = tmp_path / "inventory"
    iw.main([str(synthetic_workbook), "--output", str(out)])

    rows = _read_csv(out / "sheets.csv")
    assert [r["sheet"] for r in rows] == ["Data", "Summary", "HiddenSheet"]
    hidden = next(r for r in rows if r["sheet"] == "HiddenSheet")
    assert hidden["state"] == "hidden"
    assert hidden["formula_cells"] == "1"


def test_formulas_csv_has_one_row_per_formula_cell(synthetic_workbook, tmp_path):
    out = tmp_path / "inventory"
    iw.main([str(synthetic_workbook), "--output", str(out)])

    rows = _read_csv(out / "formulas.csv")
    # 6 en Data + 3 en Summary + 1 en HiddenSheet
    assert len(rows) == 10

    d2 = next(r for r in rows if r["sheet"] == "Data" and r["cell"] == "D2")
    assert d2["formula"] == "=C2*2"
    assert d2["row"] == "2"
    assert d2["column"] == "4"
    assert d2["formula_type"] == "arithmetic"

    countifs = next(r for r in rows if r["cell"] == "A2" and r["sheet"] == "Summary")
    assert "COUNTIFS" in countifs["function_names"]
    assert countifs["formula_type"] == "aggregation"

    vlookup = next(r for r in rows if r["cell"] == "A4" and r["sheet"] == "Summary")
    assert vlookup["formula_type"] == "lookup"


def test_formula_patterns_group_repeated_and_similar(synthetic_workbook, tmp_path):
    out = tmp_path / "inventory"
    iw.main([str(synthetic_workbook), "--output", str(out)])

    rows = _read_csv(out / "formula_patterns.csv")
    patterns = {r["pattern"]: int(r["count"]) for r in rows}

    # Las 5 fórmulas "=C{n}*2" (Data) + "=Data!C7*1" (HiddenSheet) colapsan en
    # un único patrón: la referencia con prefijo de hoja normaliza igual que la local.
    assert patterns.get("=<CELL>*<NUM>") == 6

    # Las dos COUNTIFS (solo cambia el literal de sexo) comparten patrón.
    countifs_patterns = [r for r in rows if "COUNTIFS" in r["function_names"]]
    assert len(countifs_patterns) == 1
    assert countifs_patterns[0]["count"] == "2"


def test_named_ranges_csv(synthetic_workbook, tmp_path):
    out = tmp_path / "inventory"
    iw.main([str(synthetic_workbook), "--output", str(out)])

    rows = _read_csv(out / "named_ranges.csv")
    names = {r["name"] for r in rows}
    assert "montos" in names
    montos = next(r for r in rows if r["name"] == "montos")
    assert montos["scope"] == "workbook"
    assert "C$2" in montos["refers_to"] or "C2" in montos["refers_to"]


def test_no_patient_cell_values_in_outputs(synthetic_workbook, tmp_path):
    """Los outputs no deben contener valores de celdas fuente, solo fórmulas/estructura."""
    out = tmp_path / "inventory"
    iw.main([str(synthetic_workbook), "--output", str(out)])

    formulas_text = (out / "formulas.csv").read_text(encoding="utf-8")
    # "monto" es una cabecera; los valores numéricos de C (10,20,...) no aparecen
    # como celdas fuente. Solo deben verse dentro de fórmulas (=SUM(C2:C6)).
    for line in formulas_text.splitlines()[1:]:
        assert line.count(",") >= 6


def test_missing_file_fails_clearly(tmp_path, capsys):
    code = iw.main([str(tmp_path / "nope.xlsx"), "--output", str(tmp_path / "o")])
    assert code == 2
    assert "no existe" in capsys.readouterr().err.lower()


def test_invalid_workbook_fails_clearly(tmp_path, capsys):
    bogus = tmp_path / "bogus.xlsx"
    bogus.write_text("esto no es un workbook", encoding="utf-8")

    code = iw.main([str(bogus), "--output", str(tmp_path / "o")])
    assert code == 2
    assert "no se pudo leer" in capsys.readouterr().err.lower()


def test_valid_zip_but_not_xlsx_fails_clearly(tmp_path, capsys):
    fake = tmp_path / "fake.xlsx"
    with zipfile.ZipFile(fake, "w") as zf:
        zf.writestr("hello.txt", "contenido")

    code = iw.main([str(fake), "--output", str(tmp_path / "o")])
    assert code == 2
    assert "error" in capsys.readouterr().err.lower()


def test_directory_path_fails_clearly(tmp_path, capsys):
    code = iw.main([str(tmp_path), "--output", str(tmp_path / "o")])
    assert code == 2
    assert "no es un archivo" in capsys.readouterr().err.lower()


def test_normalize_formula_examples():
    a = iw.normalize_formula('=COUNTIFS(A2:A100,"X",B2:B100,"F")')
    b = iw.normalize_formula('=COUNTIFS(A2:A100,"Y",B2:B100,"M")')
    assert a == b == "=COUNTIFS(<RANGE>,<STR>,<RANGE>,<STR>)"


def test_extract_function_names_strips_xlfn():
    assert iw.extract_function_names("=_xlfn.XLOOKUP(A1,B:B,C:C)") == ["XLOOKUP"]
    assert iw.extract_function_names("=A1+B1") == []


def test_extract_function_names_ignores_string_literals():
    # Palabras seguidas de "(" dentro de un literal no son funciones.
    assert iw.extract_function_names(
        '=IF(A1<>0,"* EMBARAZADAS (Digite CERO) *","")'
    ) == ["IF"]
    assert iw.extract_function_names(
        '=COUNTIF(A:A,"CONTROL CLINICO (ADULTOS)")'
    ) == ["COUNTIF"]

    # Otros falsos positivos reales encontrados en el workbook.
    for word in ("MIGRANTES", "ADULTOS", "INSTRUMENTAL", "OS"):
        assert iw.extract_function_names(f'=SUM(A1,"texto {word} (x)")') == ["SUM"]


def test_extract_function_names_still_detects_real_functions():
    cases = {
        "=COUNTIF(A:A,1)": ["COUNTIF"],
        '=COUNTIFS(A:A,"x",B:B,"y")': ["COUNTIFS"],
        "=SUM(A1:A9)": ["SUM"],
        '=IF(A1>0,"si","no")': ["IF"],
        '=_xlfn.IFS(A1=1,"a",A1=2,"b")': ["IFS"],
        '=DATEDIF(A1,B1,"Y")': ["DATEDIF"],
    }
    for formula, expected in cases.items():
        assert iw.extract_function_names(formula) == expected

    combined = '=IF(DATEDIF(A1,B1,"Y")>0,SUM(C1:C9),0)'
    assert iw.extract_function_names(combined) == ["DATEDIF", "IF", "SUM"]
