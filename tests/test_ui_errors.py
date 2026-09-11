"""Sprint 3.10 — traducción de errores del backend a mensajes humanos."""

from __future__ import annotations

from remasep.core.errors import RemasepError
from remasep.services.excel_writer import (
    STATUS_EXCEL_UNAVAILABLE,
    STATUS_OUTPUT_EXISTS,
    STATUS_TEMPLATE_INCOMPATIBLE,
    ExcelWriterError,
)
from remasep.services.runtime_assets import RuntimeAssetError
from remasep.ui.errors import humanize_error, humanize_generation_status


def test_runtime_asset_missing_is_human():
    h = humanize_error(RuntimeAssetError("RUNTIME_ASSET_MISSING", "detalle interno"))
    assert "instalación" in h.title.lower()
    assert "RuntimeAssetError" not in h.title and "RuntimeAssetError" not in h.detail
    assert "RUNTIME_ASSET_MISSING" in h.technical  # sólo en el detalle técnico


def test_excel_unavailable_status_is_human():
    h = humanize_generation_status(STATUS_EXCEL_UNAVAILABLE)
    assert "Microsoft Excel" in h.title
    assert "Traceback" not in h.detail


def test_template_incompatible_status_offers_template_action():
    h = humanize_generation_status(STATUS_TEMPLATE_INCOMPATIBLE)
    assert "compatible" in h.title.lower()
    assert "plantilla" in h.action_label.lower()


def test_output_exists_status_is_human():
    h = humanize_generation_status(STATUS_OUTPUT_EXISTS)
    assert "ya existe" in h.title.lower()


def test_medinet_read_error_maps_to_choose_another_file():
    exc = RemasepError("no se pudo leer 'x.xlsx': BadZipFile")
    h = humanize_error(exc)
    assert "leer" in h.title.lower()
    assert h.action_label == "Elegir otro archivo"


def test_wrong_format_error_maps_to_choose_another_file():
    exc = RemasepError("La hoja de detalle no expone columnas para: DIA_CITA")
    h = humanize_error(exc)
    assert "formato" in h.title.lower() or "Medinet" in h.title
    assert h.action_label == "Elegir otro archivo"


def test_template_incompatible_excel_writer_error_is_human():
    h = humanize_error(ExcelWriterError("TEMPLATE_INCOMPATIBLE esperado=stf:a actual=stf:b"))
    assert "compatible" in h.title.lower()
    assert "TEMPLATE_INCOMPATIBLE" in h.technical


def test_unknown_error_has_generic_human_message():
    h = humanize_error(ValueError("weird"))
    assert h.title and "ValueError" in h.technical
    assert "Traceback" not in h.title and "Traceback" not in h.detail
