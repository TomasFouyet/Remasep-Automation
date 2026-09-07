"""Tests CLI de la validación del input de producción Medinet (hotfix)."""

from __future__ import annotations

import csv
import hashlib
import json
from datetime import date

import pytest
import validate_medinet_production_input as vmpi

from remasep.services.common import Period

JULY = Period(7, 2026)


def _row(**kw):
    base = {
        "dia": date(2026, 7, 10), "nac": date(1990, 1, 1), "sexo": "Mujer",
        "sucursal": "Santiago", "especialidad": "GENETICA", "tipo": "CONSULTA GENERAL",
        "prestacion": "prestacion x", "estado": "Atendido", "modalidad": "Fonasa B",
        "prest_real": "",
    }
    base.update(kw)
    return [
        base["dia"], base["nac"], base["sexo"], base["sucursal"], base["especialidad"],
        base["tipo"], base["prestacion"], base["estado"], base["modalidad"],
        base["prest_real"],
    ]


@pytest.fixture
def pair(make_medinet):
    direct = make_medinet(
        [_row(dia=date(2026, 7, d), estado="Atendido") for d in range(1, 9)]      # 8 kept
        + [_row(dia=date(2026, 7, 9), estado="Cancelado")]                        # excluded
        + [_row(dia=date(2026, 7, 10), estado="No Se Presenta")]                  # excluded
        + [_row(dia=date(2026, 5, 1), estado="Atendido")]                         # other month
        + [_row(dia=date(2026, 9, 20), estado="Atendido")],                       # other month
        filename="direct.xlsx",
    )
    legacy = make_medinet(
        [_row(dia=date(2026, 7, d), estado="Atendido") for d in range(1, 9)],     # 8 active
        filename="legacy.xlsx",
    )
    return direct, legacy


@pytest.fixture
def outcome(pair, tmp_path):
    direct, legacy = pair
    return vmpi.run(str(direct), str(legacy), JULY, str(tmp_path / "out")), tmp_path / "out"


def test_files_written(outcome):
    _o, out = outcome
    names = {p.name for p in out.iterdir()}
    assert names == {
        "period_scope_summary.csv", "legacy_subset_comparison.csv",
        "extra_records_by_status.csv", "extra_records_by_appointment_type.csv",
        "extra_records_by_branch.csv", "extra_records_by_modality.csv",
        "extra_records_by_specialty.csv", "summary.json", "README.md",
    }


def test_summary_numbers(outcome):
    _o, out = outcome
    s = json.loads((out / "summary.json").read_text(encoding="utf-8"))
    assert s["period"] == "Julio 2026"
    assert s["raw_file_records"] == 12
    assert s["july_records"] == 10
    assert s["out_of_period_records"] == 2
    assert s["processing_scope_records"] == 10
    assert s["legacy_reference_records"] == 8
    assert s["legacy_is_exact_subset_no_estado_fp"] is True
    assert s["direct_export_excess_records"] == 2
    assert s["extra_by_status"] == {"Cancelado": 1, "No Se Presenta": 1}
    h = s["estado_processing_hypothesis"]
    assert h["status"] == "CURRENT_LEGACY_BEHAVIOR_PENDING_FUNCTIONAL_CONFIRMATION"
    assert h["direct_kept_records"] == 8
    assert h["direct_excluded_records"] == 2
    assert h["kept_states_reproduce_legacy_count_exactly"] is True


def test_period_scope_csv_has_both_files(outcome):
    _o, out = outcome
    with open(out / "period_scope_summary.csv", encoding="utf-8", newline="") as fh:
        rows = list(csv.DictReader(fh))
    assert len(rows) == 2
    direct = next(r for r in rows if r["file"].startswith("direct_export"))
    assert int(direct["physical_records"]) == 12
    assert int(direct["processing_scope_records"]) == 10


def test_source_files_unchanged(pair, tmp_path):
    direct, legacy = pair
    before = (hashlib.sha256(direct.read_bytes()).hexdigest(),
              hashlib.sha256(legacy.read_bytes()).hexdigest())
    vmpi.run(str(direct), str(legacy), JULY, str(tmp_path / "o2"))
    after = (hashlib.sha256(direct.read_bytes()).hexdigest(),
             hashlib.sha256(legacy.read_bytes()).hexdigest())
    assert before == after


def test_cli_main_smoke(pair, tmp_path, capsys):
    direct, legacy = pair
    code = vmpi.main([str(direct), str(legacy), "--month", "7", "--year", "2026",
                      "--output", str(tmp_path / "cli")])
    assert code == 0
    out = capsys.readouterr().out
    assert "processing_scope_records" in out
    assert "CURRENT_LEGACY_BEHAVIOR_PENDING_FUNCTIONAL_CONFIRMATION" in out


def test_cli_main_bad_file(tmp_path):
    code = vmpi.main([str(tmp_path / "nope.xlsx"), str(tmp_path / "nope2.xlsx")])
    assert code == 2


def test_no_pii_row_lists_in_summary(outcome):
    _o, out = outcome
    data = json.loads((out / "summary.json").read_text(encoding="utf-8"))

    def _walk(node):
        if isinstance(node, dict):
            for v in node.values():
                _walk(v)
        elif isinstance(node, list):
            assert not (node and isinstance(node[0], dict)), "summary lleva filas individuales"

    _walk(data)
