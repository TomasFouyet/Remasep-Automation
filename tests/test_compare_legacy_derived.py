"""Tests del comparador de equivalencia legacy AC:AL (Sprint 2.2).

Los workbooks sintéticos se construyen con openpyxl (fórmulas ficticias en AC:AL)
y luego se parchea el XML para inyectar valores *cacheados* — el comparador solo
lee la cache, nunca recalcula. No se usa data/local.
"""

from __future__ import annotations

import re
import shutil
import zipfile

import compare_legacy_derived as cmp
import pytest
from openpyxl import Workbook

from remasep.services.legacy_rules import load_legacy_rules
from remasep.services.legacy_transform import legacy_derived_record

_HEADERS = [
    "DIA CITA", "FECHA NACIMIENTO", "SEXO", "SUCURSAL", "ESPECIALIDAD",
    "TIPO DE CITA", "PRESTACIÓN", "ESTADO", "MODALIDAD", "PRESTACIÓN REALIZADA",
]
_DERIVED = ("AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL")


def _xml_escape(text: str) -> str:
    return text.replace("&", "&amp;").replace("<", "&lt;").replace(">", "&gt;")


def _patch_cache(path, cache: dict[tuple[int, str], object]) -> None:
    """Inyecta `<v>` en las celdas AC:AL indicadas (ref -> valor cacheado)."""
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        sheet_xml = zf.read("xl/worksheets/sheet1.xml").decode("utf-8")

    for (row, col), value in cache.items():
        ref = f"{col}{row}"
        if isinstance(value, (int, float)) and not isinstance(value, bool):
            replacement = rf'<c r="{ref}"\g<attrs>><f>\g<f></f><v>{value}</v></c>'
        else:
            replacement = (
                rf'<c r="{ref}"\g<attrs> t="str"><f>\g<f></f>'
                rf'<v>{_xml_escape(str(value))}</v></c>'
            )
        pattern = re.compile(
            rf'<c r="{ref}"(?P<attrs>[^>]*)><f>(?P<f>[^<]*)</f>(?:<v\s*/>|<v></v>)</c>'
        )
        sheet_xml, n = pattern.subn(replacement, sheet_xml)
        assert n == 1, f"no se pudo parchear {ref}"

    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(path) as src, zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for name in names:
            data = sheet_xml.encode("utf-8") if name == "xl/worksheets/sheet1.xml" else src.read(name)
            dst.writestr(name, data)
    shutil.move(tmp, path)


@pytest.fixture
def build_workbook(tmp_path):
    """Devuelve un builder: (rows, cache_overrides, drop_cache_for) -> ruta xlsx.

    `rows`: lista de dicts con claves semánticas (DIA_CITA, ...). Por defecto la
    cache AC:AL se rellena con el valor legacy EXACTO (equivalencia perfecta).
    `cache_overrides`: {(excel_row, col): valor} para forzar mismatches.
    `drop_cache_for`: set de (excel_row, col) que quedan SIN cache (unavailable).
    """
    rules = load_legacy_rules()

    def _build(rows, *, cache_overrides=None, drop_cache_for=None, filename="legacy.xlsx"):
        cache_overrides = cache_overrides or {}
        drop_cache_for = drop_cache_for or set()

        wb = Workbook()
        ws = wb.active
        ws.title = "Atenciones - Detalles de citas"
        ws.append(_HEADERS)
        cache: dict[tuple[int, str], object] = {}
        for offset, row in enumerate(rows):
            excel_row = offset + 2
            ws.append(
                [
                    row.get("DIA_CITA"), row.get("FECHA_NACIMIENTO"), row.get("SEXO"),
                    row.get("SUCURSAL"), row.get("ESPECIALIDAD"), row.get("TIPO_DE_CITA"),
                    row.get("PRESTACION"), row.get("ESTADO"), row.get("MODALIDAD"),
                    row.get("PRESTACION_REALIZADA"),
                ]
            )
            expected = legacy_derived_record({k: row.get(k) for k in row}, rules)
            for col in _DERIVED:
                ws[f"{col}{excel_row}"] = "=1"  # fórmula ficticia (solo importa data_type)
                if (excel_row, col) in drop_cache_for:
                    continue
                value = cache_overrides.get((excel_row, col), expected[col])
                if value is None:
                    continue
                cache[(excel_row, col)] = value

        path = tmp_path / filename
        wb.save(path)
        _patch_cache(path, cache)
        return path

    return _build


def _row(**kw):
    base = {
        "DIA_CITA": "01-07-2026",
        "FECHA_NACIMIENTO": "01-01-2000",
        "SEXO": "Mujer",
        "SUCURSAL": "Santiago",
        "ESPECIALIDAD": "GENETICA",
        "TIPO_DE_CITA": "CONSULTA GENERAL",
        "PRESTACION": "prestacion x",
        "ESTADO": "Atendido",
        "MODALIDAD": "Fonasa B",
        "PRESTACION_REALIZADA": "",
    }
    base.update(kw)
    return base


# --- caso feliz: equivalencia exacta ---------------------------------


def test_all_columns_exact_pass(build_workbook):
    rows = [
        _row(),
        _row(TIPO_DE_CITA="CONTROL PERIODONCIA", SEXO="Hombre"),  # AH match
        _row(PRESTACION="x CONTROL SPLINT y"),  # AJ match
        _row(FECHA_NACIMIENTO="28/06/2024"),  # AF = 2
    ]
    result = cmp.compare_legacy_derived(build_workbook(rows))

    assert result.active_records == 4
    assert result.structural_empty_rows == 0
    assert result.physical_rows == 4
    assert result.mismatches == 0
    assert result.unavailable == 0
    assert result.overall_status == "PASS"
    for c in result.column_summaries:
        assert c.status == "PASS", c.column
        assert c.exact_match_rate == 1.0


@pytest.mark.parametrize(
    ("column", "row_kwargs"),
    [
        ("AC", {"TIPO_DE_CITA": "CTRL", "SUCURSAL": "Concepcion"}),
        ("AD", {"TIPO_DE_CITA": "CTRL", "SEXO": "Hombre"}),
        ("AE", {"PRESTACION": "5001", "ESPECIALIDAD": "Odo", "SEXO": "Mujer"}),
        ("AF", {"FECHA_NACIMIENTO": "07/03/1980"}),
        ("AG", {"TIPO_DE_CITA": "CONSULTA MEDICA DE ESPECIALIDAD EN GENETICA CLINICA - COD.0101325"}),
        ("AH", {"TIPO_DE_CITA": "CONTROL ODONTOPEDIATRÍA"}),
        ("AI", {"TIPO_DE_CITA": "EVALUACIÓN PERIODONCIA"}),
        ("AJ", {"PRESTACION": "z CONTROL APARATO FIJO z"}),
        ("AK", {"PRESTACION": "z CONTROL PLACA ORTOP. PREQ. z"}),
        ("AL", {"PRESTACION": "INSTALACIÓN TAPE LABIAL PREQ."}),
    ],
)
def test_each_column_matches_its_own_cache(build_workbook, column, row_kwargs):
    result = cmp.compare_legacy_derived(build_workbook([_row(**row_kwargs)]))
    summary = {c.column: c for c in result.column_summaries}[column]
    assert summary.exact_matches == 1
    assert summary.mismatches == 0
    assert summary.status == "PASS"


# --- casing / acentos --------------------------------------------------


def test_casing_and_accents_still_exact(build_workbook):
    rows = [
        _row(TIPO_DE_CITA="control periodoncia", SEXO="Hombre"),  # AH: casing
        _row(PRESTACION="previo CONTROL MÁSCARA Y/O DISYUNTOR posterior"),  # AJ: acento
    ]
    result = cmp.compare_legacy_derived(build_workbook(rows))
    assert result.overall_status == "PASS"
    assert result.mismatches == 0


# --- structural empty ------------------------------------------------


def test_structural_empty_rows_excluded(build_workbook):
    rows = [_row(), _row()] + [_row(**{k: None for k in _row()}) for _ in range(3)]
    result = cmp.compare_legacy_derived(build_workbook(rows))
    assert result.physical_rows == 5
    assert result.structural_empty_rows == 3
    assert result.active_records == 2
    assert result.overall_status == "PASS"
    assert all(c.compared_rows == 2 for c in result.column_summaries)


# --- mismatch detectado ------------------------------------------------


def test_value_mismatch_is_detected_and_fails(build_workbook):
    path = build_workbook([_row()], cache_overrides={(2, "AC"): "VALOR DISTINTO"})
    result = cmp.compare_legacy_derived(path, capture_details=True)
    ac = {c.column: c for c in result.column_summaries}["AC"]
    assert ac.mismatches == 1
    assert ac.status == "FAIL"
    assert result.overall_status == "FAIL"
    types = {m.mismatch_type for m in result.mismatch_rows if m.column == "AC"}
    assert types == {"VALUE_MISMATCH"}


def test_semantic_mismatch_type(build_workbook):
    # mismo valor salvo tildes -> SEMANTIC_MISMATCH (no exacto)
    row = _row(TIPO_DE_CITA="CONTROL", SUCURSAL="ÑUÑOA")
    path = build_workbook([row], cache_overrides={(2, "AC"): "CONTROLNUNOA"})
    result = cmp.compare_legacy_derived(path)
    ac = {c.column: c for c in result.column_summaries}["AC"]
    assert ac.mismatches == 1
    assert {m.mismatch_type for m in result.mismatch_rows if m.column == "AC"} == {
        "SEMANTIC_MISMATCH"
    }


def test_stale_cache_af_zero_is_unavailable_not_fail(build_workbook):
    # Excel dejó 0 en AF pero la edad real es > 0: cache no confiable, no es mismatch.
    path = build_workbook([_row(FECHA_NACIMIENTO="28/06/2024")], cache_overrides={(2, "AF"): 0})
    result = cmp.compare_legacy_derived(path)
    af = {c.column: c for c in result.column_summaries}["AF"]
    assert af.mismatches == 0
    assert af.unavailable == 1
    assert af.status == "CACHE_UNAVAILABLE"
    assert result.overall_status == "CACHE_UNAVAILABLE"
    assert {m.mismatch_type for m in result.mismatch_rows} == {"STALE_CACHE_SUSPECTED"}


# --- cache unavailable ----------------------------------------------


def test_cache_unavailable_for_a_column(build_workbook):
    drop = {(2, "AC"), (3, "AC")}
    result = cmp.compare_legacy_derived(build_workbook([_row(), _row()], drop_cache_for=drop))
    ac = {c.column: c for c in result.column_summaries}["AC"]
    assert ac.unavailable == 2
    assert ac.exact_matches == 0
    assert ac.status == "CACHE_UNAVAILABLE"
    assert result.cached_values_available is False
    assert result.overall_status == "CACHE_UNAVAILABLE"


# --- múltiples categorías por registro (punto 9) ------------------


def test_record_can_hit_multiple_legacy_categories(build_workbook):
    # PRESTACION contiene un patrón de AJ y otro de AK a la vez.
    row = _row(PRESTACION="CONTROL APARATO REMOVIBLE + CONTROL PLACA ORTOP. PREQ.")
    result = cmp.compare_legacy_derived(build_workbook([row]))
    assert result.legacy_records_with_match == 1
    assert result.legacy_category_hits.get("AJ") == 1
    assert result.legacy_category_hits.get("AK") == 1
    # la suma de hits (2) supera el nº de registros con match (1): sin exclusividad
    assert sum(result.legacy_category_hits.values()) > result.legacy_records_with_match
    assert result.overall_status == "PASS"  # los valores AC:AL siguen siendo exactos


# --- privacidad del output -----------------------------------------


def test_mismatches_csv_has_no_patient_values(build_workbook, tmp_path):
    row = _row(SEXO="Hombre", PRESTACION="SECRET_PRESTACION", SUCURSAL="SECRET_SUCURSAL")
    path = build_workbook(
        [row],
        cache_overrides={(2, "AC"): "otro", (2, "AD"): "otro", (2, "AE"): "otro"},
    )
    result = cmp.compare_legacy_derived(path)
    out = tmp_path / "eq"
    cmp.write_equivalence_outputs(result, out)

    mismatches_text = (out / "mismatches.csv").read_text(encoding="utf-8")
    assert mismatches_text.splitlines()[0] == "row_number,derived_column,comparison_status,mismatch_type"
    for token in ("SECRET_PRESTACION", "SECRET_SUCURSAL", "Hombre", "otro"):
        assert token not in mismatches_text
    for name in ("summary.json", "column_summary.csv"):
        assert "SECRET" not in (out / name).read_text(encoding="utf-8")


# --- errores de entrada -------------------------------------------


def test_missing_file(tmp_path):
    with pytest.raises(cmp.EquivalenceError):
        cmp.compare_legacy_derived(tmp_path / "no-existe.xlsx")


def test_non_excel(tmp_path):
    bogus = tmp_path / "x.csv"
    bogus.write_text("a,b\n1,2\n", encoding="utf-8")
    with pytest.raises(cmp.EquivalenceError):
        cmp.compare_legacy_derived(bogus)


def test_cli_main_runs(build_workbook, tmp_path, capsys):
    path = build_workbook([_row()])
    code = cmp.main([str(path), "--output", str(tmp_path / "eq")])
    assert code == 0
    assert "OVERALL: PASS" in capsys.readouterr().out
    assert (tmp_path / "eq" / "summary.json").is_file()
