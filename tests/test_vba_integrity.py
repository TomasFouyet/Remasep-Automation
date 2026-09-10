"""Sprint 3.7B (patch final) — integridad **semántica** del proyecto VBA."""

from __future__ import annotations

import dataclasses

import pytest

from remasep.services.vba_integrity import (
    EXTRACTION_EXTRACTED,
    EXTRACTION_UNAVAILABLE_PARSE_ERROR,
    VBA_FAIL,
    VBA_PASS,
    VBA_SEMANTIC_UNAVAILABLE,
    VbaDecompressionError,
    VbaModule,
    VbaProject,
    compare_vba_projects,
    decompress_vba_container,
)

# Vector real: stream ``VBA/ThisWorkbook`` del REMASEP oficial (180 bytes
# comprimidos) y su descompresión exacta. Fijo en el test: sin dependencias.
_REAL_COMPRESSED = bytes.fromhex(
    "01b0b000417474726962757400652056425f4e616d0065203d20225468690073576f726b626f6f"
    "106b220d0a0a8c42617301028c307b3030303230503831392d001030030843070014021201243030"
    "34367d810d7c476c6f62616c01d01053706163019246616c0473650c6443726561740861626c151f"
    "507265649065636c610006496400b1085472750d424578706f047365141c54656d706c0061746544"
    "6572697603021292427573746f6d69067a04440332"
)
_REAL_DECOMPRESSED = (
    b'Attribute VB_Name = "ThisWorkbook"\r\n'
    b'Attribute VB_Base = "0{00020819-0000-0000-C000-000000000046}"\r\n'
    b"Attribute VB_GlobalNameSpace = False\r\n"
    b"Attribute VB_Creatable = False\r\n"
    b"Attribute VB_PredeclaredId = True\r\n"
    b"Attribute VB_Exposed = True\r\n"
    b"Attribute VB_TemplateDerived = False\r\n"
    b"Attribute VB_Customizable = True\r\n"
)


# ---------------------------------------------------------------------------
# Descompresión MS-OVBA §2.4.1
# ---------------------------------------------------------------------------


def test_decompress_matches_real_vector():
    assert decompress_vba_container(_REAL_COMPRESSED) == _REAL_DECOMPRESSED


def test_decompress_empty_container_is_empty():
    assert decompress_vba_container(b"\x01") == b""


def test_decompress_rejects_missing_signature():
    with pytest.raises(VbaDecompressionError):
        decompress_vba_container(b"\x00abc")


def test_decompress_rejects_truncated_header():
    with pytest.raises(VbaDecompressionError):
        decompress_vba_container(b"\x01\xb0")


# ---------------------------------------------------------------------------
# compare_vba_projects
# ---------------------------------------------------------------------------


def _mod(name: str, code: str = "Sub X\nEnd Sub", attrs: str = "a") -> VbaModule:
    import hashlib

    return VbaModule(
        name=name,
        code_sha256=hashlib.sha256(code.encode()).hexdigest(),
        attributes_sha256=hashlib.sha256(attrs.encode()).hexdigest(),
    )


def _project(mods, *, sha="sha-a", status=EXTRACTION_EXTRACTED) -> VbaProject:
    return VbaProject(
        present=True, payload_sha256=sha, modules=tuple(mods), extraction_status=status
    )


def test_identical_projects_pass():
    p = _project([_mod("A"), _mod("B")])
    cmp = compare_vba_projects(p, p)
    assert cmp.status == VBA_PASS
    assert cmp.ok is True
    assert cmp.warnings == ()


def test_binary_payload_change_only_is_pass_with_warning():
    before = _project([_mod("A"), _mod("B")], sha="sha-old")
    after = _project([_mod("A"), _mod("B")], sha="sha-new")  # Save de Excel
    cmp = compare_vba_projects(before, after)
    assert cmp.status == VBA_PASS
    assert cmp.ok is True
    assert cmp.payload_stable is False
    assert any("SEMANTIC_EQUIVALENT" in w for w in cmp.warnings)


def test_attribute_reorder_only_is_pass_with_warning():
    before = _project([_mod("A", attrs="ctrl1,ctrl2")])
    after = _project([_mod("A", attrs="ctrl2,ctrl1")])
    cmp = compare_vba_projects(before, after)
    assert cmp.status == VBA_PASS
    assert cmp.attributes_changed_modules == ("A",)
    assert any("VBA_MODULE_ATTRIBUTES_CHANGED" in w for w in cmp.warnings)


def test_code_change_fails():
    before = _project([_mod("A", code="Sub X\nEnd Sub")])
    after = _project([_mod("A", code="Sub X\n  Shell \"evil\"\nEnd Sub")])
    cmp = compare_vba_projects(before, after)
    assert cmp.status == VBA_FAIL
    assert cmp.ok is False
    assert cmp.code_changed_modules == ("A",)
    assert any("VBA_MODULE_SOURCE_CHANGED" in r for r in cmp.fail_reasons)


def test_module_added_or_removed_fails():
    before = _project([_mod("A")])
    after = _project([_mod("A"), _mod("Nuevo")])
    cmp = compare_vba_projects(before, after)
    assert cmp.status == VBA_FAIL
    assert cmp.added_modules == ("Nuevo",)
    assert any("VBA_MODULE_SET_CHANGED" in r for r in cmp.fail_reasons)


def test_vba_lost_fails():
    before = _project([_mod("A")])
    after = VbaProject(present=False, payload_sha256=None, modules=(), extraction_status="ABSENT")
    cmp = compare_vba_projects(before, after)
    assert cmp.status == VBA_FAIL
    assert any("VBA_LOST" in r for r in cmp.fail_reasons)


def test_vba_added_fails():
    before = VbaProject(present=False, payload_sha256=None, modules=(), extraction_status="ABSENT")
    after = _project([_mod("A")])
    cmp = compare_vba_projects(before, after)
    assert cmp.status == VBA_FAIL
    assert any("VBA_ADDED" in r for r in cmp.fail_reasons)


def test_no_vba_anywhere_passes():
    absent = VbaProject(
        present=False, payload_sha256=None, modules=(), extraction_status="ABSENT"
    )
    cmp = compare_vba_projects(absent, absent)
    assert cmp.status == VBA_PASS
    assert cmp.ok is True


def test_semantic_unavailable_is_fail_closed():
    """VBA presente antes y después pero sin poder comparar el código ->
    FAIL (no se asume equivalencia por presencia)."""
    before = VbaProject(
        present=True, payload_sha256="sha-a", modules=(),
        extraction_status=EXTRACTION_UNAVAILABLE_PARSE_ERROR, note="olefile roto",
    )
    after = dataclasses.replace(before, payload_sha256="sha-b")
    cmp = compare_vba_projects(before, after)
    assert cmp.status == VBA_FAIL
    assert cmp.ok is False
    assert any(VBA_SEMANTIC_UNAVAILABLE in r for r in cmp.fail_reasons)
    reason = next(r for r in cmp.fail_reasons if VBA_SEMANTIC_UNAVAILABLE in r)
    # se conserva evidencia (presencia + hashes) en el mensaje legible
    assert "sha-a" in reason and "sha-b" in reason
    assert "presente_antes=True" in reason and "presente_después=True" in reason


def test_semantic_unavailable_but_binary_stable_still_fails_closed():
    before = VbaProject(
        present=True, payload_sha256="same", modules=(),
        extraction_status="UNAVAILABLE_NO_OLEFILE",
    )
    cmp = compare_vba_projects(before, dataclasses.replace(before))
    assert cmp.status == VBA_FAIL
    assert cmp.ok is False


def test_semantic_unavailable_still_fails_if_vba_disappears():
    before = VbaProject(
        present=True, payload_sha256="sha-a", modules=(),
        extraction_status=EXTRACTION_UNAVAILABLE_PARSE_ERROR,
    )
    after = VbaProject(
        present=False, payload_sha256=None, modules=(), extraction_status="ABSENT"
    )
    cmp = compare_vba_projects(before, after)
    assert cmp.status == VBA_FAIL
    assert cmp.ok is False
