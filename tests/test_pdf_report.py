"""Sprint 3.10 — export del resumen mensual a PDF."""

from __future__ import annotations

import os
import stat
from datetime import date

import pytest
from PySide6.QtGui import QPainter

from remasep.services.common import Period
from remasep.services.medinet_summary import build_monthly_medinet_summary
from remasep.ui.pdf_report import (
    PdfExportError,
    export_summary_pdf,
    report_content,
    suggested_pdf_name,
)

pytestmark = pytest.mark.usefixtures("qapp")

_PII_SENTINELS = (b"12345678-9", b"PACIENTE DE PRUEBA", b"correo@ejemplo.cl")


def _row(*, estado="Atendido", sexo="Mujer"):
    return [
        date(2026, 7, 10), date(2016, 1, 1), sexo, "Centro", "Ortodoncia",
        "CONSULTA", "prestacion", estado, "Presencial", "",
    ]


@pytest.fixture
def summary(make_medinet):
    rows = (
        [_row(estado="Atendido", sexo="Hombre") for _ in range(20)]
        + [_row(estado="Cancelado") for _ in range(6)]
        + [_row(estado="En Sala de Espera", sexo="Mujer") for _ in range(4)]
    )
    return build_monthly_medinet_summary(make_medinet(rows), Period(7, 2026))


def test_suggested_name(summary):
    assert suggested_pdf_name(summary) == "Resumen_Medinet_2026_07.pdf"


def test_export_creates_a_pdf_file(tmp_path, summary):
    out = export_summary_pdf(summary, tmp_path / "resumen.pdf")
    assert out.is_file()
    raw = out.read_bytes()
    assert raw.startswith(b"%PDF-")
    assert raw.rstrip().endswith(b"%%EOF")
    assert len(raw) > 0


@pytest.mark.skipif(os.name == "nt", reason="chmod no representa ACL de Windows")
def test_export_to_non_writable_destination_fails_without_partial_file(tmp_path, summary):
    locked = tmp_path / "sin_permiso"
    locked.mkdir()
    locked.chmod(stat.S_IRUSR | stat.S_IXUSR)
    target = locked / "resumen.pdf"
    try:
        with pytest.raises(PdfExportError, match="iniciar|crear|escribir"):
            export_summary_pdf(summary, target)
        assert not target.exists()
    finally:
        locked.chmod(stat.S_IRWXU)


def test_painter_begin_failure_is_an_error_and_leaves_no_partial(
    tmp_path, summary, monkeypatch
):
    from remasep.ui import pdf_report

    class BeginFailurePainter:
        def begin(self, _device):
            return False

        def isActive(self):
            return False

        def end(self):
            raise AssertionError("end no corresponde si begin falló")

    monkeypatch.setattr(pdf_report, "QPainter", BeginFailurePainter)
    target = tmp_path / "fallido.pdf"
    with pytest.raises(PdfExportError, match="iniciar"):
        export_summary_pdf(summary, target)
    assert not target.exists()


def _painter_with_end_hook(*, after_end=None, force_end_result=None):
    """``QPainter`` real (mismo backend Qt de producción) cuyo ``end()`` se
    intercepta después de completar la escritura real, para forzar las ramas
    de post-verificación de ``export_summary_pdf`` sin reimplementar su lógica.
    """

    class _Painter(QPainter):
        def end(self):
            result = super().end()
            if after_end is not None:
                after_end()
            return force_end_result if force_end_result is not None else result

    return _Painter


def test_painter_end_failure_is_an_error_and_cleans_up_the_file(
    tmp_path, summary, monkeypatch
):
    from remasep.ui import pdf_report

    target = tmp_path / "sin_finalizar.pdf"
    monkeypatch.setattr(pdf_report, "QPainter", _painter_with_end_hook(force_end_result=False))
    with pytest.raises(PdfExportError, match="finalizar"):
        export_summary_pdf(summary, target)
    assert not target.exists()


def test_missing_final_file_after_successful_paint_is_an_error(
    tmp_path, summary, monkeypatch
):
    from remasep.ui import pdf_report

    target = tmp_path / "desaparecido.pdf"
    monkeypatch.setattr(
        pdf_report,
        "QPainter",
        _painter_with_end_hook(after_end=lambda: target.unlink(missing_ok=True)),
    )
    with pytest.raises(PdfExportError, match="dispositivo"):
        export_summary_pdf(summary, target)
    assert not target.exists()


def test_empty_final_file_is_an_error_and_gets_cleaned_up(tmp_path, summary, monkeypatch):
    from remasep.ui import pdf_report

    target = tmp_path / "vacio.pdf"
    monkeypatch.setattr(
        pdf_report,
        "QPainter",
        _painter_with_end_hook(after_end=lambda: target.write_bytes(b"")),
    )
    with pytest.raises(PdfExportError, match="vac"):
        export_summary_pdf(summary, target)
    assert not target.exists()


def test_truncated_final_file_is_an_error_and_gets_cleaned_up(tmp_path, summary, monkeypatch):
    from remasep.ui import pdf_report

    target = tmp_path / "truncado.pdf"

    def _truncate() -> None:
        raw = target.read_bytes()
        target.write_bytes(raw[: len(raw) // 2])

    monkeypatch.setattr(pdf_report, "QPainter", _painter_with_end_hook(after_end=_truncate))
    with pytest.raises(PdfExportError, match="incompleto"):
        export_summary_pdf(summary, target)
    assert not target.exists()


def test_export_does_not_overwrite_silently(tmp_path, summary):
    target = tmp_path / "resumen.pdf"
    export_summary_pdf(summary, target)
    with pytest.raises(FileExistsError):
        export_summary_pdf(summary, target)


def test_export_is_a_single_page(tmp_path, summary):
    pytest.importorskip("PySide6.QtPdf")
    from PySide6.QtPdf import QPdfDocument

    out = export_summary_pdf(summary, tmp_path / "una_pagina.pdf")
    doc = QPdfDocument()
    doc.load(str(out))
    assert doc.pageCount() == 1


def test_report_content_has_period_kpis_and_footer(summary):
    """El contenido del reporte (modelo, no bytes) lleva período, KPIs y pie."""
    content = report_content(summary)
    texts = content.all_text()
    assert "Julio 2026" in texts
    assert content.title == "REMASEP"
    assert "Fuente: Medinet" in content.footer_left
    assert "REMASEP Automation" in content.footer_right
    assert "no representa todavía la totalidad" in content.disclaimer
    joined = " ".join(texts)
    for label in ("Citas del período", "Consideradas para REMASEP",
                  "Excluidas por estado", "Porcentaje considerado"):
        assert label in joined
    assert str(summary.period_scope_records) in joined
    assert str(summary.included_records) in joined
    assert str(summary.excluded_records) in joined
    # 4 gráficos con títulos
    assert len(content.charts) == 4
    assert "Distribución por estado" in content.charts[0].title


def test_report_content_and_pdf_have_no_pii(tmp_path, summary):
    content = report_content(summary)
    joined = " ".join(content.all_text())
    for pii in _PII_SENTINELS:
        assert pii.decode() not in joined
    for forbidden in ("RUT", "apellido", "teléfono", "nacimiento", "@"):
        assert forbidden not in joined
    # el PDF resultante es un archivo válido no trivial
    out = export_summary_pdf(summary, tmp_path / "r.pdf")
    raw = out.read_bytes()
    assert raw.startswith(b"%PDF-")
    assert len(raw) > 4000
    for pii in _PII_SENTINELS:
        assert pii not in raw
