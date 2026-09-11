"""Gráficos de barras horizontales — QPainter puro, sin dependencias externas.

`draw_bar_chart` es una función pura que pinta sobre cualquier ``QPainter``
(pantalla o PDF). `BarChartCard` la envuelve en una tarjeta de la aplicación.
La UI y el export PDF usan **la misma** función de dibujo.
"""

from __future__ import annotations

from dataclasses import dataclass

from PySide6.QtCore import QRectF, QSize, Qt
from PySide6.QtGui import QColor, QFont, QPainter
from PySide6.QtWidgets import QSizePolicy, QWidget

from remasep.ui.styles import THEME, format_int


@dataclass(frozen=True)
class Bar:
    label: str
    value: int
    color: str = THEME.chart_series[0]


def draw_bar_chart(
    painter: QPainter,
    rect: QRectF,
    bars: list[Bar],
    *,
    label_width: float = 150.0,
    row_height: float = 26.0,
    row_gap: float = 8.0,
    value_suffix: str = "",
) -> float:
    """Pinta barras horizontales dentro de ``rect``. Devuelve el alto usado."""
    painter.save()
    painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
    painter.setRenderHint(QPainter.RenderHint.TextAntialiasing, True)

    max_value = max((b.value for b in bars), default=0) or 1
    bar_area_x = rect.left() + label_width + 10
    bar_area_w = max(rect.width() - label_width - 10 - 62, 30)

    label_font = QFont(painter.font())
    label_font.setPointSizeF(9.0)
    value_font = QFont(label_font)
    value_font.setBold(True)

    y = rect.top()
    for bar in bars:
        row = QRectF(rect.left(), y, rect.width(), row_height)

        painter.setFont(label_font)
        painter.setPen(THEME.qcolor(THEME.text_muted))
        painter.drawText(
            QRectF(rect.left(), y, label_width, row_height),
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            _elide(bar.label, 26),
        )

        track = QRectF(bar_area_x, y + row_height * 0.18, bar_area_w, row_height * 0.64)
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(THEME.qcolor(THEME.chart_grid))
        painter.drawRoundedRect(track, 3, 3)

        width = bar_area_w * (bar.value / max_value)
        filled = QRectF(bar_area_x, y + row_height * 0.18, max(width, 2.0), row_height * 0.64)
        painter.setBrush(QColor(bar.color))
        painter.drawRoundedRect(filled, 3, 3)

        painter.setFont(value_font)
        painter.setPen(THEME.qcolor(THEME.text))
        painter.drawText(
            QRectF(bar_area_x + bar_area_w + 8, y, 54, row_height),
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignRight,
            format_int(bar.value) + value_suffix,
        )
        y = row.bottom() + row_gap

    painter.restore()
    return (y - rect.top()) if bars else 0.0


class BarChartCard(QWidget):
    """Tarjeta con encabezado + subtítulo + gráfico de barras (repintado)."""

    def __init__(
        self,
        title: str,
        *,
        subtitle: str = "",
        legend: list[tuple[str, str]] | None = None,
        parent: QWidget | None = None,
    ) -> None:
        super().__init__(parent)
        self.setObjectName("card")
        self._title = title
        self._subtitle = subtitle
        self._legend = legend or []
        self._bars: list[Bar] = []
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.MinimumExpanding)
        self._recompute_min_height()

    def set_bars(self, bars: list[Bar]) -> None:
        self._bars = list(bars)
        self._recompute_min_height()
        self.updateGeometry()  # avisa al layout del nuevo alto
        self.update()

    def set_subtitle(self, text: str) -> None:
        self._subtitle = text
        self._recompute_min_height()
        self.updateGeometry()
        self.update()

    def _content_height(self) -> int:
        header = 58 + (16 if self._subtitle else 0) + (22 if self._legend else 0)
        rows = max(len(self._bars), 1) * 34 + 12
        return int(header + rows + 24)

    def _recompute_min_height(self) -> None:
        self.setMinimumHeight(self._content_height())

    def sizeHint(self) -> QSize:
        return QSize(320, self._content_height())

    def minimumSizeHint(self) -> QSize:
        return QSize(240, self._content_height())

    def paintEvent(self, event) -> None:
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        r = self.rect().adjusted(THEME.pad, THEME.pad, -THEME.pad, -THEME.pad)
        x, y, w = float(r.left()), float(r.top()), float(r.width())

        title_font = QFont(painter.font())
        title_font.setPointSizeF(11.0)
        title_font.setBold(True)
        painter.setFont(title_font)
        painter.setPen(THEME.qcolor(THEME.text))
        painter.drawText(QRectF(x, y, w, 22), Qt.AlignmentFlag.AlignLeft, self._title)
        y += 24

        if self._subtitle:
            sub_font = QFont(painter.font())
            sub_font.setPointSizeF(8.5)
            sub_font.setBold(False)
            painter.setFont(sub_font)
            painter.setPen(THEME.qcolor(THEME.text_faint))
            painter.drawText(QRectF(x, y, w, 16), Qt.AlignmentFlag.AlignLeft, self._subtitle)
            y += 18

        if self._legend:
            y += _draw_legend(painter, QRectF(x, y, w, 18), self._legend) + 6

        if self._bars:
            draw_bar_chart(painter, QRectF(x, y + 4, w, float(r.height()) - (y - r.top())), self._bars)
        else:
            painter.setPen(THEME.qcolor(THEME.text_faint))
            painter.drawText(QRectF(x, y + 8, w, 40), Qt.AlignmentFlag.AlignLeft, "Sin datos.")
        painter.end()


def _draw_legend(painter: QPainter, rect: QRectF, legend: list[tuple[str, str]]) -> float:
    painter.save()
    font = QFont(painter.font())
    font.setPointSizeF(8.5)
    font.setBold(False)
    painter.setFont(font)
    x = rect.left()
    for label, color in legend:
        painter.setPen(Qt.PenStyle.NoPen)
        painter.setBrush(QColor(color))
        painter.drawRoundedRect(QRectF(x, rect.top() + 3, 11, 11), 2, 2)
        painter.setPen(THEME.qcolor(THEME.text_muted))
        text_w = painter.fontMetrics().horizontalAdvance(label) + 6
        painter.drawText(
            QRectF(x + 16, rect.top(), text_w, rect.height()),
            Qt.AlignmentFlag.AlignVCenter | Qt.AlignmentFlag.AlignLeft,
            label,
        )
        x += 16 + text_w + 18
    painter.restore()
    return rect.height()


def _elide(text: str, limit: int) -> str:
    return text if len(text) <= limit else text[: limit - 1] + "…"
