"""Tests del servicio `workbook_delta` (Sprint 3.4). Workbooks sintéticos."""

from __future__ import annotations

import re
import shutil
import zipfile

import pytest
from openpyxl import Workbook

from remasep.services import workbook_delta as wd


def _inject_cache(path, cache_by_sheet: dict[str, dict[str, object]]) -> None:
    """Inyecta ``<v>`` en celdas con ``<f>`` (openpyxl no cachea resultados)."""
    with zipfile.ZipFile(path) as zf:
        names = zf.namelist()
        blobs = {n: zf.read(n) for n in names}

    order = [n for n in names if re.fullmatch(r"xl/worksheets/sheet\d+\.xml", n)]
    order.sort(key=lambda n: int(re.search(r"(\d+)", n).group(1)))
    sheet_files = dict(zip(cache_by_sheet.keys(), order, strict=False))

    for sheet, cache in cache_by_sheet.items():
        target = sheet_files[sheet]
        xml = blobs[target].decode("utf-8")
        for ref, value in cache.items():
            if isinstance(value, (int, float)) and not isinstance(value, bool):
                repl = rf'<c r="{ref}"\g<a>><f>\g<f></f><v>{value}</v></c>'
            else:
                esc = str(value).replace("&", "&amp;").replace("<", "&lt;")
                repl = rf'<c r="{ref}"\g<a> t="str"><f>\g<f></f><v>{esc}</v></c>'
            pat = re.compile(
                rf'<c r="{ref}"(?P<a>[^>]*)><f>(?P<f>[^<]*)</f>(?:<v\s*/>|<v></v>)</c>'
            )
            xml, n = pat.subn(repl, xml)
            assert n == 1, f"no se pudo cachear {sheet}!{ref}"
        blobs[target] = xml.encode("utf-8")

    tmp = str(path) + ".tmp"
    with zipfile.ZipFile(tmp, "w", zipfile.ZIP_DEFLATED) as dst:
        for name in names:
            dst.writestr(name, blobs[name])
    shutil.move(tmp, path)


@pytest.fixture
def make_pair(tmp_path):
    def _build(build_fn, *, cache_before=None, cache_after=None):
        wb_b = Workbook()
        wb_a = Workbook()
        build_fn(wb_b, "before")
        build_fn(wb_a, "after")
        pb = tmp_path / "before.xlsx"
        pa = tmp_path / "after.xlsx"
        wb_b.save(pb)
        wb_a.save(pa)
        if cache_before:
            _inject_cache(pb, cache_before)
        if cache_after:
            _inject_cache(pa, cache_after)
        return pb, pa
    return _build


# ---------------------------------------------------------------------------


def test_identical_workbooks_have_no_delta(make_pair):
    def build(wb, _phase):
        ws = wb.active
        ws.title = "REMASEP 01"
        ws["A1"] = "x"
        ws["B2"] = 5

    pb, pa = make_pair(build)
    pair = wd.open_pair(pb, pa)
    try:
        assert wd.diff_formulas(pair) == []
        assert wd.diff_values(pair) == []
        sc = wd.compare_structure(pair)
        assert sc.significant_changes == []
    finally:
        pair.close()


def test_direct_input_change_detected(make_pair):
    def build(wb, phase):
        ws = wb.active
        ws.title = "REMASEP B1"
        ws["C13"] = 0 if phase == "before" else 7

    pb, pa = make_pair(build)
    pair = wd.open_pair(pb, pa)
    try:
        deltas = wd.diff_values(pair)
        assert len(deltas) == 1
        d = deltas[0]
        assert (d.sheet, d.cell, d.change_kind) == ("REMASEP B1", "C13", wd.DIRECT_INPUT_CHANGE)
        assert (d.before_value, d.after_value) == (0, 7)
        assert d.formula == ""
    finally:
        pair.close()


def test_formula_result_change_when_cache_differs_but_formula_same(make_pair):
    def build(wb, _phase):
        ws = wb.active
        ws.title = "REMASEP 01"
        ws["A1"] = 1
        ws["B1"] = "=A1*2"

    pb, pa = make_pair(
        build,
        cache_before={"REMASEP 01": {"B1": 2}},
        cache_after={"REMASEP 01": {"B1": 8}},
    )
    pair = wd.open_pair(pb, pa)
    try:
        assert wd.diff_formulas(pair) == []  # la expresión no cambió
        deltas = wd.diff_values(pair)
        by = {(d.sheet, d.cell): d for d in deltas}
        assert ("REMASEP 01", "B1") in by
        d = by[("REMASEP 01", "B1")]
        assert d.change_kind == wd.FORMULA_RESULT_CHANGE
        assert d.formula == "=A1*2"
        assert (d.before_value, d.after_value) == (2, 8)
    finally:
        pair.close()


def test_formula_expression_change_is_reported(make_pair):
    def build(wb, phase):
        ws = wb.active
        ws.title = "REMASEP 01"
        ws["A1"] = 3
        ws["B1"] = "=A1*2" if phase == "before" else "=A1*3"

    pb, pa = make_pair(build)
    pair = wd.open_pair(pb, pa)
    try:
        changes = wd.diff_formulas(pair)
        assert len(changes) == 1
        c = changes[0]
        assert (c.sheet, c.cell, c.change_type) == ("REMASEP 01", "B1", wd.FORMULA_MODIFIED)
        assert (c.before_formula, c.after_formula) == ("=A1*2", "=A1*3")
    finally:
        pair.close()


# --- change_kind con fórmula ANTES + DESPUÉS -----------------------


def test_formula_added_is_expression_change_not_direct_nor_result(make_pair):
    def build(wb, phase):
        ws = wb.active
        ws.title = "REMASEP 01"
        ws["A1"] = 3
        ws["B1"] = 5 if phase == "before" else "=A1*2"  # antes: valor; después: fórmula

    pb, pa = make_pair(build, cache_after={"REMASEP 01": {"B1": 6}})
    pair = wd.open_pair(pb, pa)
    try:
        d = {x.key: x for x in wd.diff_values(pair)}[("REMASEP 01", "B1")]
        assert d.change_kind == wd.FORMULA_EXPRESSION_CHANGE
        assert d.change_kind not in (wd.DIRECT_INPUT_CHANGE, wd.FORMULA_RESULT_CHANGE)
        assert d.before_formula == "" and d.formula == "=A1*2"
        assert d.formula_change_type == wd.FORMULA_ADDED
        assert (d.before_value, d.after_value) == (5, 6)
        # y diff_formulas también lo lista
        fc = wd.diff_formulas(pair)
        assert [(c.cell, c.change_type) for c in fc] == [("B1", wd.FORMULA_ADDED)]
    finally:
        pair.close()


def test_formula_removed_is_expression_change(make_pair):
    def build(wb, phase):
        ws = wb.active
        ws.title = "REMASEP 01"
        ws["A1"] = 3
        ws["B1"] = "=A1*2" if phase == "before" else 9  # antes: fórmula; después: valor

    pb, pa = make_pair(build, cache_before={"REMASEP 01": {"B1": 6}})
    pair = wd.open_pair(pb, pa)
    try:
        d = {x.key: x for x in wd.diff_values(pair)}[("REMASEP 01", "B1")]
        assert d.change_kind == wd.FORMULA_EXPRESSION_CHANGE
        assert d.before_formula == "=A1*2" and d.formula == ""
        assert d.formula_change_type == wd.FORMULA_REMOVED
        assert (d.before_value, d.after_value) == (6, 9)
    finally:
        pair.close()


def test_formula_modified_with_value_change_is_expression_change(make_pair):
    def build(wb, phase):
        ws = wb.active
        ws.title = "REMASEP 01"
        ws["A1"] = 3
        ws["B1"] = "=A1*2" if phase == "before" else "=A1*3"

    pb, pa = make_pair(
        build,
        cache_before={"REMASEP 01": {"B1": 6}},
        cache_after={"REMASEP 01": {"B1": 9}},
    )
    pair = wd.open_pair(pb, pa)
    try:
        deltas = wd.diff_values(pair)
        d = {x.key: x for x in deltas}[("REMASEP 01", "B1")]
        assert d.change_kind == wd.FORMULA_EXPRESSION_CHANGE
        assert d.formula_change_type == wd.FORMULA_MODIFIED
        assert (d.before_formula, d.formula) == ("=A1*2", "=A1*3")
        # ninguna clasificación como input directo ni resultado de fórmula
        assert all(
            x.change_kind not in (wd.DIRECT_INPUT_CHANGE, wd.FORMULA_RESULT_CHANGE)
            for x in deltas if x.key == ("REMASEP 01", "B1")
        )
    finally:
        pair.close()


def test_formula_change_type_empty_for_non_expression_kinds(make_pair):
    def build(wb, phase):
        ws = wb.active
        ws.title = "REMASEP 01"
        ws["C1"] = 0 if phase == "before" else 1

    pb, pa = make_pair(build)
    pair = wd.open_pair(pb, pa)
    try:
        d = wd.diff_values(pair)[0]
        assert d.change_kind == wd.DIRECT_INPUT_CHANGE
        assert d.formula_change_type == ""
    finally:
        pair.close()


def test_structure_comparison_flags_added_sheet_and_dims(make_pair):
    def build(wb, phase):
        ws = wb.active
        ws.title = "REMASEP 01"
        ws["A1"] = "h"
        if phase == "after":
            ws["A50"] = "grew"
            wb.create_sheet("NUEVA")

    pb, pa = make_pair(build)
    pair = wd.open_pair(pb, pa)
    try:
        sc = wd.compare_structure(pair)
        assert "NUEVA" in sc.added_sheets
        assert any("dimensiones" in s for s in sc.significant_changes)
    finally:
        pair.close()


def test_delta_edges_cover_cell_range_and_cross_sheet(make_pair):
    def build(wb, phase):
        r01 = wb.active
        r01.title = "REMASEP 01"
        b2 = wb.create_sheet("B2 ANEXO")
        v = 0 if phase == "before" else 1
        r01["A1"] = v                       # direct
        r01["A2"] = "=A1"                    # cell dep on A1
        r01["A3"] = "=SUM(A1:A2)"            # range dep spanning A1
        b2["C100"] = v                       # direct in other sheet
        r01["A4"] = "='B2 ANEXO'!C100"       # cross-sheet dep

    pb, pa = make_pair(
        build,
        cache_before={"REMASEP 01": {"A2": 0, "A3": 0, "A4": 0}},
        cache_after={"REMASEP 01": {"A2": 1, "A3": 1, "A4": 1}},
    )
    pair = wd.open_pair(pb, pa)
    try:
        deltas = wd.diff_values(pair)
        edges = wd.build_delta_edges(deltas)
        got = {(e.source_cell, e.target_cell, e.dependency_type) for e in edges}
        assert ("A1", "A2", "cell") in got
        assert ("A1", "A3", "range") in got
        assert ("C100", "A4", "cross_sheet") in got
    finally:
        pair.close()


def test_open_pair_does_not_modify_source_files(make_pair):
    def build(wb, _phase):
        wb.active["A1"] = 1

    pb, pa = make_pair(build)
    before_hash = (wd.compute_sha256(pb), wd.compute_sha256(pa))
    pair = wd.open_pair(pb, pa)
    wd.diff_values(pair)
    wd.diff_formulas(pair)
    wd.compare_structure(pair)
    pair.close()
    assert (wd.compute_sha256(pb), wd.compute_sha256(pa)) == before_hash


def test_missing_file_raises(tmp_path):
    with pytest.raises(wd.WorkbookDeltaError):
        wd.open_pair(tmp_path / "nope.xlsx", tmp_path / "nope2.xlsx")
