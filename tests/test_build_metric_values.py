"""Sprint 3.7A — pruebas de los ayudantes analíticos de ``scripts/build_metric_values.py``.

Sintéticas: no se abren workbooks reales ni ``data/local``. Sólo se ejercitan
funciones puras con estructuras de datos fabricadas.
"""

from __future__ import annotations

import build_metric_values as bmv
import pytest

from remasep.core.errors import RemasepError
from remasep.services.metric_value_producer import (
    ZERO_POLICY_UNRESOLVED,
    ZERO_POLICY_WRITE,
)


def test_normalize_form_maps_the_three_sheets() -> None:
    assert bmv.normalize_form("REMASEP_OD") == "REMASEP_OD"
    assert bmv.normalize_form("REMASEP 01") == "REMASEP_01"
    assert bmv.normalize_form("B2 ANEXO") == "B2_ANEXO"
    assert bmv.normalize_form("Otra Hoja") == "Otra_Hoja"


def test_parse_period_roundtrip_and_error() -> None:
    period = bmv._parse_period("2026-07")
    assert (period.month, period.year) == (7, 2026)
    with pytest.raises(RemasepError):
        bmv._parse_period("julio-2026")


def test_section_of_signature() -> None:
    sig = "REMASEP_OD :: SECCION B: ODONTOLOGIA :: 5004020 - CONTROL :: X :: 25-34"
    assert bmv._section_of(sig) == "SECCION B: ODONTOLOGIA"
    assert bmv._section_of("solo-un-tramo") == ""


def _zero_row(form: str, sem: str) -> dict:
    return {"form": form, "zero_semantics": sem}


def test_derive_zero_policy_write_zero_when_evidence_is_consistent() -> None:
    rows = [_zero_row("REMASEP_OD", "SOURCE_ZERO_TARGET_ZERO") for _ in range(120)]
    rows += [_zero_row("REMASEP_OD", "SOURCE_ZERO_TARGET_OTHER") for _ in range(2)]
    result = bmv.derive_zero_policy(rows, min_cases=30, min_ratio=0.98)
    assert result["resolution"] == ZERO_POLICY_WRITE
    assert result["observed"]["source_zero_target_zero"] == 120
    assert result["observed"]["source_zero_target_blank"] == 0


def test_derive_zero_policy_unresolved_when_insufficient_or_mixed() -> None:
    too_few = [_zero_row("REMASEP_OD", "SOURCE_ZERO_TARGET_ZERO") for _ in range(10)]
    assert bmv.derive_zero_policy(too_few, min_cases=30, min_ratio=0.98)["resolution"] == (
        ZERO_POLICY_UNRESOLVED
    )
    mixed = (
        [_zero_row("A", "SOURCE_ZERO_TARGET_ZERO") for _ in range(40)]
        + [_zero_row("A", "SOURCE_ZERO_TARGET_BLANK") for _ in range(40)]
    )
    assert bmv.derive_zero_policy(mixed, min_cases=30, min_ratio=0.98)["resolution"] == (
        ZERO_POLICY_UNRESOLVED
    )


def test_compare_zero_policy_config_reports_agreement() -> None:
    derived = {"resolution": ZERO_POLICY_WRITE, "observed": {}}
    ok = bmv.compare_zero_policy_config({"resolution": "WRITE_ZERO"}, derived)
    assert ok["config_matches_evidence"] is True
    bad = bmv.compare_zero_policy_config({"resolution": "UNRESOLVED"}, derived)
    assert bad["config_matches_evidence"] is False


def test_value_equivalence_summary_buckets_by_form() -> None:
    rows = [
        {"form": "REMASEP_OD", "comparison_status": "MATCH"},
        {"form": "REMASEP_OD", "comparison_status": "MISMATCH"},
        {"form": "REMASEP_01", "comparison_status": "ZERO_VS_BLANK_EQUIVALENT"},
        {"form": "REMASEP_01", "comparison_status": "TARGET_UNAVAILABLE"},
    ]
    summary = {row["form"]: row for row in bmv.value_equivalence_summary(rows)}
    assert summary["ALL"]["total"] == 4
    assert summary["REMASEP_OD"]["MATCH"] == 1
    assert summary["REMASEP_OD"]["MISMATCH"] == 1
    assert summary["REMASEP_01"]["ZERO_VS_BLANK"] == 1
    assert summary["REMASEP_01"]["UNAVAILABLE"] == 1


class _Ref:
    def __init__(self) -> None:
        self.kind = {"m1": "BASE_AGGREGATION", "m2": "DERIVED_AGGREGATION"}
        self.cache_status = {"m1": "CACHE_DIFFERENCE", "m2": ""}


def test_cluster_mismatches_groups_and_counts() -> None:
    mismatch_rows = [
        {
            "instruction_id": f"wi:{i}", "source_metric_id": "m1", "form": "REMASEP_OD",
            "source_kind": "BASE_AGGREGATION", "target_sheet": "REMASEP_OD",
            "target_cell": f"A{i}", "delta": i - 1, "comparison_status": "MISMATCH",
            "zero_semantics": "NOT_APPLICABLE", "investigation_axis": "LEGACY_CACHE_SUSPECT",
        }
        for i in range(3)
    ]
    mismatch_rows.append(
        {
            "instruction_id": "wi:x", "source_metric_id": "m2", "form": "REMASEP_01",
            "source_kind": "DERIVED_AGGREGATION", "target_sheet": "REMASEP 01",
            "target_cell": "B2", "delta": 5, "comparison_status": "MISMATCH",
            "zero_semantics": "NOT_APPLICABLE", "investigation_axis": "UNKNOWN",
        }
    )
    clusters = bmv.cluster_mismatches(mismatch_rows)
    assert clusters[0]["count"] == 3
    assert clusters[0]["delta_min"] == -1 and clusters[0]["delta_max"] == 1
    assert {c["form"] for c in clusters} == {"REMASEP_OD", "REMASEP_01"}


def test_investigation_axis_prioritises_legacy_cache_match() -> None:
    ref = _Ref()
    row = {
        "source_metric_id": "m1", "zero_semantics": "SOURCE_ZERO_TARGET_OTHER",
        "comparison_status": "MISMATCH", "reason": "LEGACY_CACHE_MATCHES_TARGET",
    }
    assert bmv._investigation_axis(row, ref) == "LEGACY_CACHE_MATCHES_TARGET"
    row["reason"] = ""
    assert bmv._investigation_axis(row, ref) == "ZERO_SEMANTICS"


def test_manifest_row_exposes_write_instruction_interface() -> None:
    raw = {
        "instruction_id": "wi:1", "source_metric_id": "LEGACY::REMASEP_OD::A1",
        "source_semantic_signature": "REMASEP_OD :: SEC :: x",
        "target_sheet": "REMASEP_OD", "target_cell": "B2",
        "expected_value_type": "INTEGER_COUNT",
    }
    row = bmv.ManifestRow(raw)
    assert row.instruction_id == "wi:1"
    assert row.form == "REMASEP_OD"
    assert row.expected_value_type == "INTEGER_COUNT"
