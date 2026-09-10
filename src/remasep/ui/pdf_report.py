"""Export del **resumen mensual Medinet** a un PDF profesional de una página.

No es un screenshot: se compone un reporte limpio con ``QPdfWriter`` +
``QPainter``, reutilizando :func:`remasep.ui.charts.draw_bar_chart`. Consume el
mismo :class:`MonthlyMedinetSummary` que la UI.

Separación: :func:`report_content` arma **qué** texto/series van en la página
(testeable sin Qt de pantalla); :func:`export_summary_pdf` sólo lo **pinta**.
**Sin PII.** No sobrescribe: si el archivo existe lanza ``FileExistsError``.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

from PySide6.QtCore import QMarginsF, QRectF, Qt
from PySide6.QtGui import QFont, QPageLayout, QPageSize, QPainter, QPdfWriter

from remasep.services.medinet_summary import MonthlyMedinetSummary
from remasep.ui.charts import Bar, draw_bar_chart
from remasep.ui.styles import THEME, format_int

_FOOTER_APP = "Generado por REMASEP Automation"
_FOOTER_SOURCE = "Fuente: Medinet"
_DISCLAIMER = (
    "Este resumen corresponde únicamente a datos provenientes de Medinet y no "
    "representa todavía la totalidad del informe REMASEP."
)


@dataclass(frozen=True)
class _Kpi:
    label: str
    value: str
    hint: str


@dataclass(frozen=True)
class _ChartSpec:
    title: str
    bars: tuple[Bar, ...]
    legend: tuple[tuple[str, str], ...] = ()


@dataclass(frozen=True)
class ReportContent:
    title: str
    subtitle: str
    period: str
    kpis: tuple[_Kpi, ...]
    charts: tuple[_ChartSpec, ...]
    footer_left: str
    footer_right: str
    disclaimer: str

    def all_text(self) -> list[str]:
        out = [self.title, self.subtitle, self.period, self.footer_left,
               self.footer_right, self.disclaimer]
        for k in self.kpis:
            out += [k.label, k.value, k.hint]
        for c in self.charts:
            out.append(c.title)
            out += [b.label for b in c.bars]
        return out


def suggested_pdf_name(summary: MonthlyMedinetSummary) -> str:
    return f"Resumen_Medinet_{summary.period_year:04d}_{summary.period_month:02d}.pdf"


def _series_bars(items) -> tuple[Bar, ...]:
    return tuple(
        Bar(c.label, c.count, THEME.chart_series[i % len(THEME.chart_series)])
        for i, c in enumerate(items)
    )


def report_content(s: MonthlyMedinetSummary) -> ReportContent:
    return ReportContent(
        title="REMASEP",
        subtitle="Resumen mensual · Fuente: Medinet",
        period=s.period_label,
        kpis=(
            _Kpi("Citas del período", format_int(s.period_scope_records),
                 "Válidas dentro del mes"),
            _Kpi("Consideradas para REMASEP", format_int(s.included_records),
                 "Sólo estados confirmados"),
            _Kpi("Excluidas por estado", format_int(s.excluded_records),
                 "Canceladas, no presentadas, etc."),
            _Kpi("Porcentaje considerado", f"{s.included_percentage:g}%",
                 s.considered_ratio_label),
        ),
        charts=(
            _ChartSpec(
                "Distribución por estado (Medinet)",
                tuple(
                    Bar(e.label, e.count,
                        THEME.chart_included if e.included else THEME.chart_excluded)
                    for e in s.estado_distribution
                ),
                (("Consideradas", THEME.chart_included), ("Excluidas", THEME.chart_excluded)),
            ),
            _ChartSpec("Atenciones consideradas por especialidad",
                       _series_bars(s.service_distribution)),
            _ChartSpec("Consideradas por sexo", _series_bars(s.sex_distribution)),
            _ChartSpec("Consideradas por tramo de edad", _series_bars(s.age_distribution)),
        ),
        footer_left=_FOOTER_SOURCE,
        footer_right=_FOOTER_APP,
        disclaimer=_DISCLAIMER,
    )


def export_summary_pdf(summary: MonthlyMedinetSummary, path: str | Path) -> Path:
    """Escribe el PDF de una página. Devuelve la ruta. No sobrescribe."""
    out = Path(path)
    if out.exists():
        raise FileExistsError(f"ya existe un archivo en {out}")
    out.parent.mkdir(parents=True, exist_ok=True)

    content = report_content(summary)
    writer = QPdfWriter(str(out))
    writer.setResolution(150)
    writer.setPageSize(QPageSize(QPageSize.PageSizeId.A4))
    writer.setPageMargins(QMarginsF(16, 14, 16, 14), QPageLayout.Unit.Millimeter)
    writer.setTitle(f"Resumen mensual Medinet — {summary.period_label}")

    painter = QPainter(writer)
    try:
        _paint(painter, content)
    finally:
        painter.end()
    return out


def _font(painter: QPainter, px: int, *, bold: bool = False) -> QFont:
    font = QFont(painter.font())
    font.setPixelSize(px)
    font.setBold(bold)
    return font


def _text(painter: QPainter, rect: QRectF, value: str, *, flags=Qt.AlignmentFlag.AlignLeft) -> None:
    painter.drawText(rect, int(flags) | int(Qt.AlignmentFlag.AlignVCenter), value)


# Alturas (px @150dpi) de cada bloque de gráfico, para que nada se solape / recorte.
_CHART_HEADER = 34
_CHART_SUBTITLE = 22
_CHART_LEGEND = 24
_CHART_ROW = 30
_CHART_PAD = 26


def _chart_height(spec: _ChartSpec) -> float:
    h = _CHART_HEADER + len(spec.bars) * _CHART_ROW + _CHART_PAD
    if spec.legend:
        h += _CHART_LEGEND
    return float(h)


def _paint(painter: QPainter, c: ReportContent) -> None:
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)
    page = painter.viewport()
    w = float(page.width())
    x = 0.0
    y = 0.0

    # --- cabecera ---
    painter.setFont(_font(painter, 36, bold=True))
    painter.setPen(THEME.qcolor(THEME.primary))
    _text(painter, QRectF(x, y, w, 46), c.title)
    y += 46
    painter.setFont(_font(painter, 17))
    painter.setPen(THEME.qcolor(THEME.text_muted))
    _text(painter, QRectF(x, y, w, 24), c.subtitle)
    y += 26
    painter.setFont(_font(painter, 23, bold=True))
    painter.setPen(THEME.qcolor(THEME.text))
    _text(painter, QRectF(x, y, w, 32), c.period)
    y += 40
    painter.setPen(THEME.qcolor(THEME.border))
    painter.drawLine(int(x), int(y), int(w), int(y))
    y += 20

    # --- KPIs ---
    gap = 18
    kpi_w = (w - 3 * gap) / 4
    kpi_h = 132.0
    for i, kpi in enumerate(c.kpis):
        _kpi(painter, QRectF(x + i * (kpi_w + gap), y, kpi_w, kpi_h), kpi)
    y += kpi_h + 22

    # --- gráficos 2x2, alto por fila según su contenido ---
    chart_w = (w - gap) / 2
    for left, right in ((c.charts[0], c.charts[1]), (c.charts[2], c.charts[3])):
        row_h = max(_chart_height(left), _chart_height(right))
        _chart_block(painter, QRectF(x, y, chart_w, row_h), left)
        _chart_block(painter, QRectF(x + chart_w + gap, y, chart_w, row_h), right)
        y += row_h + 18

    # --- pie ---
    foot_y = float(page.height()) - 64
    painter.setPen(THEME.qcolor(THEME.border))
    painter.drawLine(int(x), int(foot_y), int(w), int(foot_y))
    painter.setFont(_font(painter, 13))
    painter.setPen(THEME.qcolor(THEME.text_faint))
    _text(painter, QRectF(x, foot_y + 8, w, 18), c.footer_left)
    _text(painter, QRectF(x, foot_y + 8, w, 18), c.footer_right, flags=Qt.AlignmentFlag.AlignRight)
    painter.drawText(
        QRectF(x, foot_y + 28, w, 34),
        int(Qt.AlignmentFlag.AlignLeft) | int(Qt.TextFlag.TextWordWrap),
        c.disclaimer,
    )


def _kpi(painter: QPainter, rect: QRectF, kpi: _Kpi) -> None:
    painter.save()
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(THEME.qcolor(THEME.surface_muted))
    painter.drawRoundedRect(rect, 8, 8)
    left = rect.left() + 14
    inner_w = rect.width() - 28

    painter.setFont(_font(painter, 15))
    painter.setPen(THEME.qcolor(THEME.text_muted))
    painter.drawText(
        QRectF(left, rect.top() + 12, inner_w, 34),
        int(Qt.AlignmentFlag.AlignLeft | Qt.AlignmentFlag.AlignTop) | int(Qt.TextFlag.TextWordWrap),
        kpi.label,
    )
    painter.setFont(_font(painter, 34, bold=True))
    painter.setPen(THEME.qcolor(THEME.text))
    _text(painter, QRectF(left, rect.top() + 52, inner_w, 40), kpi.value)
    painter.setFont(_font(painter, 12))
    painter.setPen(THEME.qcolor(THEME.text_faint))
    _text(painter, QRectF(left, rect.top() + 98, inner_w, 22), kpi.hint)
    painter.restore()


def _chart_block(painter: QPainter, rect: QRectF, spec: _ChartSpec) -> None:
    painter.save()
    painter.setPen(Qt.PenStyle.NoPen)
    painter.setBrush(THEME.qcolor(THEME.surface_muted))
    painter.drawRoundedRect(rect, 8, 8)
    inner = rect.adjusted(16, 14, -16, -14)

    painter.setFont(_font(painter, 16, bold=True))
    painter.setPen(THEME.qcolor(THEME.text))
    _text(painter, QRectF(inner.left(), inner.top(), inner.width(), 20), spec.title)
    y = inner.top() + _CHART_HEADER

    if spec.legend:
        painter.setFont(_font(painter, 12))
        lx = inner.left()
        for label, color in spec.legend:
            painter.setPen(Qt.PenStyle.NoPen)
            painter.setBrush(THEME.qcolor(color))
            painter.drawRoundedRect(QRectF(lx, y + 2, 12, 12), 3, 3)
            painter.setPen(THEME.qcolor(THEME.text_muted))
            tw = painter.fontMetrics().horizontalAdvance(label) + 6
            _text(painter, QRectF(lx + 17, y, tw, 16), label)
            lx += 17 + tw + 18
        y += _CHART_LEGEND

    painter.setFont(_font(painter, 13))  # etiquetas de barras: regular, no hereda el bold del título
    draw_bar_chart(
        painter,
        QRectF(inner.left(), y, inner.width(), rect.bottom() - 14 - y),
        list(spec.bars),
        label_width=inner.width() * 0.46,
        row_height=24.0,
        row_gap=6.0,
    )
    painter.restore()
