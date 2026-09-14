"""F02 — contrato semántico de plantilla (complementa el fingerprint estructural).

``structural_template_fingerprint`` (writable_target_mapping.py) protege
coordenadas/protección/merges pero no el TEXTO de las fórmulas: una plantilla
con ``REMASEP 01!C15`` o ``CONTROL!E16`` reescritas a ``=0`` conserva el mismo
fingerprint (auditoría 2026-09-14, F02).

Estrategia: certificar TODA fórmula estática de la plantilla oficial
(``config/excel_writer_2026/template_static_formulas.csv``, generado por
``scripts/build_template_semantic_contract.py``), no un subconjunto curado a
mano — un subconjunto pequeño dejaba pasar sin detectar exactamente el repro
original de C15. Estas pruebas son sintéticas —no abren ni modifican el
workbook real— salvo las marcadas explícitamente, gateadas a que exista
``data/local`` (igual que el resto de la suite).
"""

from __future__ import annotations

from pathlib import Path

import pytest
import yaml

from remasep.services import excel_writer as w
from remasep.services.metric_value_producer import PendingWrite
from remasep.testing import FakeWorkbookModel, FakeWriterHarness

_CONFIG_DIR = Path("config/excel_writer_2026")
_CONTRACT_PATH = _CONFIG_DIR / "template_semantic_contract.yaml"
_OFFICIAL_TEMPLATE = Path("data/local/REMASEP 2026_V1.4.xlsm")


def _inline_contract(formulas: list[tuple[str, str, str]]) -> w.TemplateSemanticContract:
    """Contrato sintético pequeño (formato inline, sin CSV externo)."""
    return w.load_template_semantic_contract(
        {
            "version": "test_semantic_contract",
            "critical_formulas": [
                {"sheet": s, "cell": c, "formula": f} for s, c, f in formulas
            ],
        }
    )


def _real_frozen_contract() -> w.TemplateSemanticContract:
    """El contrato de PRODUCCIÓN congelado (CSV real, miles de filas)."""
    raw = yaml.safe_load(_CONTRACT_PATH.read_text(encoding="utf-8"))
    return w.load_template_semantic_contract(raw, base_dir=_CONFIG_DIR)


# ---------------------------------------------------------------------------
# 1. plantilla oficial (contrato congelado real) -> PASS
# ---------------------------------------------------------------------------


def test_official_contract_is_self_compatible():
    contract = _real_frozen_contract()
    assert len(contract.critical_formulas) > 10_000  # toda fórmula estática, no un subconjunto
    formula_map = {(cf.sheet, cf.cell): cf.formula for cf in contract.critical_formulas}
    result = w.verify_template_semantic_contract(formula_map, contract)
    assert result.ok is True
    assert result.violations == ()


# ---------------------------------------------------------------------------
# 2. CONTROL!E16 -> "=0" -> FAIL (repro original de la auditoría)
# ---------------------------------------------------------------------------


def test_real_contract_catches_control_e16_rewritten_to_zero():
    contract = _real_frozen_contract()
    formula_map = {(cf.sheet, cf.cell): cf.formula for cf in contract.critical_formulas}
    formula_map[("CONTROL", "E16")] = "=0"
    result = w.verify_template_semantic_contract(formula_map, contract)
    assert result.ok is False
    violated = {(v.sheet, v.cell) for v in result.violations}
    assert ("CONTROL", "E16") in violated


# ---------------------------------------------------------------------------
# 3. REMASEP 01!C15 -> "=0" -> FAIL (el repro que el contrato de 28 celdas
#    curadas a mano dejaba pasar; ESTE es el test que cierra F02 de verdad)
# ---------------------------------------------------------------------------


def test_real_contract_catches_remasep01_c15_rewritten_to_zero():
    contract = _real_frozen_contract()
    formula_map = {(cf.sheet, cf.cell): cf.formula for cf in contract.critical_formulas}
    assert ("REMASEP 01", "C15") in formula_map  # el contrato SÍ la conoce
    formula_map[("REMASEP 01", "C15")] = "=0"
    result = w.verify_template_semantic_contract(formula_map, contract)
    assert result.ok is False
    violated = {(v.sheet, v.cell) for v in result.violations}
    assert ("REMASEP 01", "C15") in violated


# ---------------------------------------------------------------------------
# 4. cualquier fórmula declarada eliminada (se volvió valor/blanco) -> FAIL
# ---------------------------------------------------------------------------


def test_critical_formula_removed_is_incompatible():
    contract = _inline_contract([("REMASEP 01", "C15", "=SUM(D15:K15)")])
    formula_map: dict[tuple[str, str], str] = {}  # ya no es fórmula: no está en formula_map
    result = w.verify_template_semantic_contract(formula_map, contract)
    assert result.ok is False
    assert result.violations[0].reason == "MISSING"


def test_real_contract_removed_formula_is_incompatible():
    contract = _real_frozen_contract()
    formula_map = {(cf.sheet, cf.cell): cf.formula for cf in contract.critical_formulas}
    del formula_map[("REMASEP 01", "C15")]
    result = w.verify_template_semantic_contract(formula_map, contract)
    assert result.ok is False
    violated = {(v.sheet, v.cell): v.reason for v in result.violations}
    assert violated[("REMASEP 01", "C15")] == "MISSING"


# ---------------------------------------------------------------------------
# re-guardado / representación legítima equivalente -> sin falso rechazo
# ---------------------------------------------------------------------------


def test_legitimate_resave_representation_is_not_rejected():
    contract = _inline_contract([("CONTROL", "E7", "=+'REMASEP 01'!B307")])
    # mismo significado, distinta representación binaria tras un re-guardado:
    # espacios y case fuera de literales, nunca dentro de un literal.
    formula_map = {("CONTROL", "E7"): " =+'remasep 01'!B307 "}
    result = w.verify_template_semantic_contract(formula_map, contract)
    assert result.ok is True


# ---------------------------------------------------------------------------
# 5. cambio en un valor mensual/manual -> sin falso rechazo
# ---------------------------------------------------------------------------


def test_unrelated_monthly_value_change_does_not_affect_contract():
    """El contrato sólo mira celdas que SON fórmula en la plantilla en
    blanco: un dato mensual/manual (nunca una fórmula ahí) no participa."""
    contract = _inline_contract([("CONTROL", "E16", "=E6+E7+E8+E9+E10+E11+E12+E13+E14")])
    formula_map = {
        ("CONTROL", "E16"): "=E6+E7+E8+E9+E10+E11+E12+E13+E14",
        ("NOMBRE", "B5"): "Julio",  # dato manual del mes, no es una fórmula
    }
    result = w.verify_template_semantic_contract(formula_map, contract)
    assert result.ok is True


def test_real_contract_write_targets_never_collide_with_certified_formulas():
    """Ningún target de write_manifest.csv (donde escribimos valores cada mes)
    coincide con una celda certificada — si coincidiera, el contrato
    bloquearía cada generación real. Verificado también en build time por
    scripts/build_template_semantic_contract.py; aquí se re-confirma sobre el
    contrato ya congelado."""
    import csv as _csv

    contract = _real_frozen_contract()
    certified = {(cf.sheet, cf.cell) for cf in contract.critical_formulas}
    manifest_path = Path("config/runtime_2026/write_manifest.csv")
    with manifest_path.open(newline="", encoding="utf-8") as fh:
        targets = {(row["target_sheet"], row["target_cell"]) for row in _csv.DictReader(fh)}
    assert certified & targets == set()


# ---------------------------------------------------------------------------
# 6. fórmula nueva inesperada en una celda que debería ser input.
#
# DECISIÓN (documentada, no asumida en silencio): el contrato NO audita
# "celdas que deberían ser input" — sólo certifica que las fórmulas YA
# declaradas en la plantilla en blanco sigan intactas. Si una celda-target
# (donde nosotros escribimos) aparece con una fórmula inesperada, eso ya lo
# detecta una barrera EXISTENTE y anterior: `write_value2` se niega a
# sobrescribir una fórmula (`TARGET_FORMULA_CONFLICT`, defensa en profundidad
# §9) — el contrato semántico no necesita duplicar esa responsabilidad, y
# extenderlo a "toda celda que no es target tampoco debería tener fórmula"
# no es una regla confirmada (podría ser una fórmula auxiliar legítima fuera
# de nuestro alcance de escritura). No se amplía el contrato para esto.
# ---------------------------------------------------------------------------


def test_unexpected_formula_on_a_target_cell_is_caught_by_write_value2_not_the_contract(tmp_path):
    pending = [
        PendingWrite("wi:0", "LEGACY::X::0", "REMASEP_OD", "B10", 1, "INTEGER_COUNT", "Julio 2026")
    ]
    model = FakeWorkbookModel(
        sheet_names=["REMASEP_OD", "CONTROL"],
        # B10 debería ser un input (write target); aquí aparece con una
        # fórmula inesperada -- simula una plantilla alterada.
        formula_map={("REMASEP_OD", "B10"): "=1+1"},
        values={("CONTROL", "E16"): 0},
        writable_cells={("REMASEP_OD", "B10")},
        structural_fingerprint_id="stf:contract-test",
    )
    harness = FakeWriterHarness(model)
    result = w.generate(
        _request(tmp_path),
        pending,
        open_writer=harness.open_writer,
        control_map=_CONTROL_MAP,
        inspect=harness.inspect,
    )
    assert result.status == w.STATUS_FAILED_INTEGRITY_CHECK
    assert any("TARGET_FORMULA_CONFLICT" in e for e in result.errors)


# ---------------------------------------------------------------------------
# integración con generate(): opt-in, no rompe llamadas existentes
# ---------------------------------------------------------------------------

_SHEETS = ["REMASEP_OD", "CONTROL"]
_CONTROL_MAP = w.load_control_map(
    {
        "sheet": "CONTROL",
        "total_errors_cell": "E16",
        "data_status_ok_value": "OK",
        "data_status_missing_value": "SIN DATOS",
        "modules_not_yet_produced": [],
        "rows": [],
    }
)


def _pending() -> list[PendingWrite]:
    return [PendingWrite("wi:0", "LEGACY::X::0", "REMASEP_OD", "B10", 1, "INTEGER_COUNT", "Julio 2026")]


def _request(tmp_path: Path) -> w.GenerationRequest:
    template = tmp_path / "tpl.xlsm"
    template.write_bytes(b"PK\x03\x04 placeholder xlsm")
    return w.GenerationRequest(
        mode=w.MODE_DIAGNOSTIC_REFERENCE,
        template_path=template,
        output_path=tmp_path / "out.xlsm",
        period_year=2026,
        period_month=7,
        expected_fingerprint_id="stf:contract-test",
        zero_write_policy="WRITE_ZERO",
        zero_write_policy_accepted=("WRITE_ZERO",),
        manifest_instruction_ids=("wi:0",),
    )


def test_generate_rejects_output_when_critical_formula_altered(tmp_path):
    model = FakeWorkbookModel(
        sheet_names=_SHEETS,
        formula_map={("CONTROL", "E16"): "=0"},  # ya no suma nada: la brecha F02
        values={("REMASEP_OD", "B10"): 0, ("CONTROL", "E16"): 0},
        writable_cells={("REMASEP_OD", "B10")},
        structural_fingerprint_id="stf:contract-test",
    )
    harness = FakeWriterHarness(model)
    contract = _inline_contract([("CONTROL", "E16", "=E6+E7+E8+E9+E10+E11+E12+E13+E14")])
    result = w.generate(
        _request(tmp_path),
        _pending(),
        open_writer=harness.open_writer,
        control_map=_CONTROL_MAP,
        inspect=harness.inspect,
        semantic_contract=contract,
    )
    assert result.status == w.STATUS_TEMPLATE_INCOMPATIBLE
    assert any("TEMPLATE_SEMANTIC_CONTRACT_VIOLATION" in e for e in result.errors)
    assert any("E16" in e for e in result.errors)
    assert not (tmp_path / "out.xlsm").exists()
    assert harness.last_writer is None  # nunca llegó a abrir el writer


def test_generate_without_semantic_contract_is_unaffected(tmp_path):
    """Backward-compat: los llamadores que no pasan ``semantic_contract`` (todo
    el resto de la suite, código existente) no cambian de comportamiento."""
    model = FakeWorkbookModel(
        sheet_names=_SHEETS,
        formula_map={("CONTROL", "E16"): "=0"},
        values={("REMASEP_OD", "B10"): 0, ("CONTROL", "E16"): 0},
        writable_cells={("REMASEP_OD", "B10")},
        structural_fingerprint_id="stf:contract-test",
    )
    harness = FakeWriterHarness(model)
    result = w.generate(
        _request(tmp_path),
        _pending(),
        open_writer=harness.open_writer,
        control_map=_CONTROL_MAP,
        inspect=harness.inspect,
    )
    assert result.status != w.STATUS_TEMPLATE_INCOMPATIBLE


# ---------------------------------------------------------------------------
# validación cruzada contra el workbook oficial real (gateada, como el resto
# de la suite: no se ejecuta sin data/local de desarrollo)
# ---------------------------------------------------------------------------


@pytest.mark.skipif(
    not _OFFICIAL_TEMPLATE.is_file(), reason="requiere la plantilla oficial en data/local"
)
def test_official_template_matches_frozen_contract():
    import openpyxl

    contract = _real_frozen_contract()
    wb = openpyxl.load_workbook(_OFFICIAL_TEMPLATE, data_only=False, keep_vba=True, read_only=True)
    formula_map = {}
    for ws in wb.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    formula_map[(ws.title, cell.coordinate)] = str(cell.value)
    wb.close()
    # el contrato debe conocer EXACTAMENTE las mismas celdas de fórmula que
    # el workbook real -- ni de menos (huecos) ni de más (celdas que ya no
    # son fórmula habrían quedado congeladas por error).
    assert set(formula_map) == {(cf.sheet, cf.cell) for cf in contract.critical_formulas}
    result = w.verify_template_semantic_contract(formula_map, contract)
    assert result.ok is True, result.violations
