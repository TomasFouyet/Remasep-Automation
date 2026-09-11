"""Sprint 3.10 — export del resumen mensual a PDF."""

from __future__ import annotations

from datetime import date

import pytest

from remasep.services.common import Period
from remasep.services.medinet_summary import build_monthly_medinet_summary
from remasep.ui.pdf_report import export_summary_pdf, report_content, suggested_pdf_name

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
    head = out.read_bytes()[:5]
    assert head == b"%PDF-"


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
