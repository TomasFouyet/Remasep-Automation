"""Tests del cierre por dependencias (Sprint 2.4), con workbooks sintéticos.

Hoja de detalle con columnas en posiciones reales (D, G, H, K, M, O, AA, …) +
`REMASEP_OD` con fórmulas ``COUNTIF(S)`` base y derivadas ``base ± celda``,
con valores cacheados parcheados en el XML. No se usa data/local.
"""

from __future__ import annotations

import csv
import json
import xml.etree.ElementTree as ET
import zipfile

import close_legacy_aggregations as clo
import pytest
from openpyxl import Workbook

_DETAIL = "Atenciones - Detalles de citas"
_COLS = {
    "DIA_CITA": "D", "FECHA_NACIMIENTO": "G", "SEXO": "H", "SUCURSAL": "K",
    "ESPECIALIDAD": "M", "TIPO_DE_CITA": "O", "PRESTACION": "AA",
    "ESTADO": "P", "MODALIDAD": "J", "PRESTACION_REALIZADA": "AB",
}
_HDR = {
    "D": "DIA CITA", "G": "FECHA NACIMIENTO", "H": "SEXO", "K": "SUCURSAL",
    "M": "ESPECIALIDAD", "O": "TIPO DE CITA", "AA": "PRESTACIÓN",
    "P": "ESTADO", "J": "MODALIDAD", "AB": "PRESTACIÓN REALIZADA", "BA": "MARKER",
}


def _sheet_xml_map(path) -> dict[str, str]:
    ns_m = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
    ns_r = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
    with zipfile.ZipFile(path) as zf:
        wbxml = ET.fromstring(zf.read("xl/workbook.xml"))
        rels = ET.fromstring(zf.read("xl/_rels/workbook.xml.rels"))
    rid_target = {rel.get("Id"): rel.get("Target") for rel in rels}
    out = {}
    for sheet in wbxml.find(f"{{{ns_m}}}sheets"):
        target = rid_target[sheet.get(f"{{{ns_r}}}id")].lstrip("/")
        if not target.startswith("xl/"):
            target = "xl/" + target
        out[sheet.get("name")] = target
    return out


@pytest.fixture
def build_closure_workbook(tmp_path, inject_formula_cache):
    def _build(detail_rows, sheet_formulas, cache=None, *, filename="closure.xlsx"):
        cache = cache or {}
        wb = Workbook()
        detail = wb.active
        detail.title = _DETAIL
        for letter, header in _HDR.items():
            detail[f"{letter}1"] = header
        for i, row in enumerate(detail_rows, start=2):
            for sem, letter in _COLS.items():
                if row.get(sem) is not None:
                    detail[f"{letter}{i}"] = row[sem]
            detail[f"BA{i}"] = "x"
        for sheet_name, formulas in sheet_formulas.items():
            ws = wb.create_sheet(sheet_name)
            for coord, formula in formulas.items():
                ws[coord] = formula
        path = tmp_path / filename
        wb.save(path)

        xml_map = _sheet_xml_map(path)
        for sheet_name, cells in cache.items():
            patch = {(int(c[1:]), c[0]): v for c, v in cells.items()}
            if patch:
                inject_formula_cache(path, patch, sheet_xml=xml_map[sheet_name])
        return path

    return _build


def _row(**kw):
    base = {
        "DIA_CITA": "01-07-2026", "FECHA_NACIMIENTO": "01-01-2000", "SEXO": "Hombre",
        "SUCURSAL": "Santiago", "ESPECIALIDAD": "Odontología General",
        "TIPO_DE_CITA": "OTRA", "PRESTACION": "COMPOSITE", "ESTADO": "Atendido",
        "MODALIDAD": "Fonasa B", "PRESTACION_REALIZADA": "",
    }
    base.update(kw)
    return base


_AE_H = '"*COMPOSITE*"&"*Odontología General*"&"*Hombre*"'


def _countifs(af_crit: str, ae: str = _AE_H) -> str:
    return (
        f"=COUNTIFS('{_DETAIL}'!$AE:$AE,{ae},"
        f"'{_DETAIL}'!$AF:$AF,\"{af_crit}\")"
    )


@pytest.fixture
def scenario(build_closure_workbook):
    rows = [
        _row(FECHA_NACIMIENTO="01-01-2025"),   # AF 1, AE COMPOSITE·OdontGeneral·Hombre
        _row(FECHA_NACIMIENTO="01-06-2025"),   # AF 1
        _row(FECHA_NACIMIENTO="01-01-2020"),   # AF 6
        _row(TIPO_DE_CITA="CONTROL PERIODONCIA", PRESTACION="x", FECHA_NACIMIENTO="01-01-2000"),
        {},  # fila estructuralmente vacía
    ]
    od = {
        # --- base -----------------------------------------------------
        "F30": _countifs("<2"),                       # -> 2 (R1,R2), depende de AF
        "P10": _countifs("<2"),                       # -> 2, raíz de la cadena
        # --- derivadas ------------------------------------------------
        "F46": _countifs("<2", '"*COMPOSITE*"&"*Hombre*"') + "-F30",  # 2 - 2 = 0
        "U10": f"=COUNTIF('{_DETAIL}'!$AC:$AC,\"zzz\")-F30",          # 0 - 2 = -2; AF por F30
        "P11": f"=COUNTIF('{_DETAIL}'!$AC:$AC,\"zzz\")+P10",          # cadena depth 1 -> 2
        "P12": f"=COUNTIF('{_DETAIL}'!$AC:$AC,\"zzz\")+P11",          # cadena depth 2 -> 2
        # --- ciclo --------------------------------------------------
        "Q10": f"=COUNTIF('{_DETAIL}'!$AC:$AC,\"zzz\")+Q11",
        "Q11": f"=COUNTIF('{_DETAIL}'!$AC:$AC,\"zzz\")+Q10",
        # --- dependencia faltante --------------------------------
        "R10": f"=COUNTIF('{_DETAIL}'!$AC:$AC,\"zzz\")-Z99",  # Z99 no es fórmula
        "Z99": "=1+1",  # nodo que NO depende de Medinet (no evaluable como dep de valor)
        # --- downstream / validation / other -----------------------
        "S10": "=SUM(F30:F46)",             # DOWNSTREAM_TOTAL
        "T10": "=IF(F30>0,1,0)",            # VALIDATION
        "S11": "=F30*2",                    # UNSUPPORTED_OTHER
    }
    cache = {
        "REMASEP_OD": {
            "F30": 2, "P10": 2, "P11": 2,
            "F46": 9,      # <- fuerza CACHE_DIFFERENCE (depende de AF -> nota)
            "U10": -2,
            # P12 sin cache -> CACHE_UNAVAILABLE
        }
    }
    path = build_closure_workbook(rows, {"REMASEP_OD": od}, cache)
    return clo.close_legacy_aggregations(path), path


def _nodes(result):
    return {(n.sheet, n.cell): n for n in result.medinet_nodes()}


# --- universo / clasificación --------------------------------------


def test_dataset_counts(scenario):
    result, _ = scenario
    assert result.physical_rows == 5
    assert result.structural_empty_rows == 1
    assert result.active_records == 4


def test_classification_base_derived_total_validation_other(scenario):
    result, _ = scenario
    nodes = _nodes(result)
    assert nodes[("REMASEP_OD", "F30")].kind == clo.KIND_BASE
    assert nodes[("REMASEP_OD", "F46")].kind == clo.KIND_DERIVED
    assert nodes[("REMASEP_OD", "S10")].kind == clo.KIND_TOTAL
    assert nodes[("REMASEP_OD", "T10")].kind == clo.KIND_VALIDATION
    assert nodes[("REMASEP_OD", "S11")].kind == clo.KIND_OTHER


def test_transitive_vs_direct(scenario):
    result, _ = scenario
    nodes = _nodes(result)
    assert nodes[("REMASEP_OD", "F30")].direct is True     # referencia detalle
    assert nodes[("REMASEP_OD", "P11")].direct is True     # tiene COUNTIF sobre detalle
    assert nodes[("REMASEP_OD", "S10")].direct is False    # solo SUM(F30:F46)
    assert nodes[("REMASEP_OD", "S10")].medinet is True    # transitivo vía F30..F46
    # Z99 = 1+1 no toca Medinet
    assert ("REMASEP_OD", "Z99") not in nodes


# --- evaluación base + derivadas -----------------------------


def test_base_aggregations_evaluate(scenario):
    result, _ = scenario
    nodes = _nodes(result)
    assert nodes[("REMASEP_OD", "F30")].python_value == 2
    assert nodes[("REMASEP_OD", "P10")].python_value == 2


def test_derived_countifs_minus_cell(scenario):
    result, _ = scenario
    node = _nodes(result)[("REMASEP_OD", "F46")]
    assert node.python_value == 0        # COUNTIFS(2) - F30(2)
    assert node.evaluation_supported is True


def test_derived_countif_plus_reference_chain(scenario):
    result, _ = scenario
    nodes = _nodes(result)
    assert nodes[("REMASEP_OD", "P11")].python_value == 2
    assert nodes[("REMASEP_OD", "P12")].python_value == 2
    assert nodes[("REMASEP_OD", "P11")].depth == 1
    assert nodes[("REMASEP_OD", "P12")].depth == 2


def test_topological_order_places_dependency_first(scenario):
    result, _ = scenario
    nodes = _nodes(result)
    # P10 (base) es raíz; P11 y P12 la necesitan
    assert nodes[("REMASEP_OD", "P10")].depth == 0
    assert nodes[("REMASEP_OD", "P12")].depth > nodes[("REMASEP_OD", "P11")].depth


# --- ciclo / dependencia faltante -------------------------


def test_cycle_detected_and_not_evaluated(scenario):
    result, _ = scenario
    nodes = _nodes(result)
    q10 = nodes[("REMASEP_OD", "Q10")]
    q11 = nodes[("REMASEP_OD", "Q11")]
    assert q10.status == clo.STATUS_CYCLE
    assert q11.status == clo.STATUS_CYCLE
    assert q10.evaluation_supported is False
    assert len(result.cycle_nodes) >= 2


def test_missing_dependency_reported(scenario):
    result, _ = scenario
    node = _nodes(result)[("REMASEP_OD", "R10")]
    assert node.status == clo.STATUS_MISSING
    assert node.evaluation_supported is False
    assert any(cell == "R10" for _s, cell, _dep in result.missing_dependencies)


def test_downstream_unsupported_does_not_break_base_evaluation(scenario):
    result, _ = scenario
    summary = clo._summary_dict(result)
    assert summary["base_supported"] == summary["base_aggregation_cells"]
    assert summary["unsupported_other_cells"] == 1  # S11 = F30*2
    assert summary["downstream_total_cells"] == 1   # S10


# --- depends_on_AF -----------------------------------------


def test_depends_on_af_propagates_through_value_reference(scenario):
    result, _ = scenario
    nodes = _nodes(result)
    # U10 = COUNTIF(AC,...) - F30 : su COUNTIF no toca AF, pero F30 sí
    assert nodes[("REMASEP_OD", "U10")].depends_on_age_local is False
    assert nodes[("REMASEP_OD", "F30")].depends_on_age is True
    assert nodes[("REMASEP_OD", "U10")].depends_on_age is True


# --- cache ------------------------------------------------


def test_cache_match(scenario):
    result, _ = scenario
    assert _nodes(result)[("REMASEP_OD", "F30")].cache_status == clo.CACHE_MATCH


def test_cache_difference_with_af_note(scenario):
    result, _ = scenario
    node = _nodes(result)[("REMASEP_OD", "F46")]
    assert node.cache_status == clo.CACHE_DIFFERENCE
    assert "depends on AF" in node.notes


def test_cache_unavailable(scenario):
    result, _ = scenario
    node = _nodes(result)[("REMASEP_OD", "P12")]
    assert node.evaluation_supported is True
    assert node.cache_status == clo.CACHE_UNAVAILABLE


# --- artefactos ------------------------------------------


def test_metric_id_is_stable(scenario, tmp_path):
    result, _ = scenario
    clo.write_outputs(result, tmp_path / "out")
    with open(tmp_path / "out" / "metric_nodes.csv", encoding="utf-8", newline="") as handle:
        rows = {r["metric_id"]: r for r in csv.DictReader(handle)}
    assert "LEGACY::REMASEP_OD::F46" in rows
    assert rows["LEGACY::REMASEP_OD::F46"]["kind"] == "DERIVED_AGGREGATION"
    assert rows["LEGACY::REMASEP_OD::F46"]["dependencies"] == "LEGACY::REMASEP_OD::F30"


def test_dependency_edges_use_same_sheet_value(scenario, tmp_path):
    result, _ = scenario
    clo.write_outputs(result, tmp_path / "out")
    with open(
        tmp_path / "out" / "aggregation_dependency_edges.csv", encoding="utf-8", newline=""
    ) as handle:
        edges = list(csv.DictReader(handle))
    value_edges = {
        (e["source_cell"], e["target_cell"])
        for e in edges
        if e["dependency_type"] == "SAME_SHEET_VALUE"
    }
    assert ("F30", "F46") in value_edges
    assert ("P10", "P11") in value_edges


def test_summary_json_separates_support_and_cache(scenario, tmp_path):
    result, _ = scenario
    clo.write_outputs(result, tmp_path / "out")
    summary = json.loads((tmp_path / "out" / "summary.json").read_text(encoding="utf-8"))
    # kind es independiente de si se pudo evaluar:
    assert summary["derived_aggregation_cells"] == 7     # F46,U10,P11,P12 + Q10,Q11 + R10
    assert summary["derived_supported"] == 4             # ciclo y faltante NO se evalúan
    assert summary["base_aggregation_cells"] == 2
    assert summary["cycles"] >= 2
    assert summary["missing_dependencies"] == 1
    assert summary["cache_matches"] >= 1
    assert summary["cache_differences"] >= 1
    assert summary["cache_unavailable"] >= 1
    assert summary["formula_support_status"] in {"PASS", "PARTIAL"}
    assert summary["cache_consistency_status"] == "DIFFERENCES"


def test_artifacts_contain_no_individual_patient_data(build_closure_workbook, tmp_path):
    secret = "SECRET-PATIENT-XYZ"
    rows = [
        _row(FECHA_NACIMIENTO="01-01-2025", ESTADO=secret, PRESTACION_REALIZADA=secret),
        _row(FECHA_NACIMIENTO="01-01-2020", ESTADO=secret),
    ]
    od = {"F30": _countifs("<2"), "F46": _countifs("<2", '"*COMPOSITE*"&"*Hombre*"') + "-F30"}
    path = build_closure_workbook(rows, {"REMASEP_OD": od}, {"REMASEP_OD": {"F30": 1, "F46": 0}})
    result = clo.close_legacy_aggregations(path)
    files = clo.write_outputs(result, tmp_path / "out")
    for f in files:
        assert secret not in f.read_text(encoding="utf-8"), f.name


# --- errores de entrada ---------------------------------


def test_missing_file(tmp_path):
    with pytest.raises(clo.ClosureError):
        clo.close_legacy_aggregations(tmp_path / "nope.xlsx")


def test_cli_main_returns_zero(scenario, tmp_path, capsys):
    _result, path = scenario
    code = clo.main([str(path), "--output", str(tmp_path / "out")])
    assert code == 0
    out = capsys.readouterr().out
    assert "formula_support_status:" in out
    assert "cache_consistency_status: DIFFERENCES" in out
