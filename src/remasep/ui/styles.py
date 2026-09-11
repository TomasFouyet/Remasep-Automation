"""Tema visual central de la aplicación (Sprint 3.10).

Un único lugar para colores, tipografía, espaciado y QSS. Los gráficos QPainter
y el export PDF consumen las mismas constantes (:data:`THEME`), sin duplicar
estilos.
"""

from __future__ import annotations

from dataclasses import dataclass, field

from PySide6.QtGui import QColor
from PySide6.QtWidgets import QWidget


@dataclass(frozen=True)
class Theme:
    # --- superficie / texto ---
    bg: str = "#f4f6f8"
    surface: str = "#ffffff"
    surface_muted: str = "#f8fafc"
    border: str = "#e2e6ea"
    border_strong: str = "#cfd6dd"
    text: str = "#1c2530"
    text_muted: str = "#5b6875"
    text_faint: str = "#8a949f"

    # --- color institucional (azul sanitario sobrio) ---
    primary: str = "#1f5f8b"
    primary_dark: str = "#184d70"
    primary_soft: str = "#e8f0f6"

    # --- semántica de estado ---
    ok: str = "#1f7a4d"
    ok_soft: str = "#e7f4ec"
    warning: str = "#8a6100"
    warning_soft: str = "#fbf1dd"
    danger: str = "#b42318"
    danger_soft: str = "#fdeceb"
    neutral: str = "#64748b"

    # --- gráficos ---
    chart_included: str = "#1f5f8b"     # atenciones consideradas
    chart_excluded: str = "#c2ccd6"     # excluidas por estado
    chart_series: tuple[str, ...] = field(
        default_factory=lambda: (
            "#1f5f8b", "#2f8fae", "#5aa9a0", "#7cbf87", "#a8cf74",
            "#d7cf6b", "#e0a95f", "#d98060", "#c76a86", "#9d6ea3",
        )
    )
    chart_axis: str = "#c8d0d8"
    chart_grid: str = "#eef1f4"

    # --- medidas ---
    radius: int = 10
    radius_sm: int = 6
    gap: int = 12
    gap_lg: int = 20
    pad: int = 20

    def qcolor(self, hex_value: str) -> QColor:
        return QColor(hex_value)


THEME = Theme()


STYLESHEET = f"""
QWidget {{
    color: {THEME.text};
    font-size: 14px;
}}
QMainWindow, QStackedWidget, QStackedWidget > QWidget, QScrollArea, QScrollArea > QWidget > QWidget {{
    background: {THEME.bg};
}}
QScrollArea {{ border: none; }}
QLabel {{ background: transparent; }}

QLabel[role="h1"]       {{ font-size: 28px; font-weight: 600; }}
QLabel[role="h2"]       {{ font-size: 19px; font-weight: 600; }}
QLabel[role="h3"]       {{ font-size: 15px; font-weight: 600; }}
QLabel[role="subtitle"] {{ font-size: 15px; color: {THEME.text_muted}; }}
QLabel[role="muted"]    {{ color: {THEME.text_muted}; font-size: 12.5px; }}
QLabel[role="faint"]    {{ color: {THEME.text_faint}; font-size: 12px; }}
QLabel[role="section"]  {{
    font-size: 11px; font-weight: 700; color: {THEME.text_faint};
    letter-spacing: 1.2px;
}}

QLabel[role="kpi-value"]        {{ font-size: 30px; font-weight: 600; color: {THEME.text}; }}
QLabel[role="kpi-value-accent"] {{ font-size: 30px; font-weight: 600; color: {THEME.primary}; }}
QLabel[role="kpi-label"] {{ font-size: 12.5px; color: {THEME.text_muted}; }}
QLabel[role="kpi-hint"]  {{ font-size: 11.5px; color: {THEME.text_faint}; }}

QLabel[status="ok"]      {{ color: {THEME.ok}; }}
QLabel[status="warning"] {{ color: {THEME.warning}; }}
QLabel[status="error"]   {{ color: {THEME.danger}; }}
QLabel[status="neutral"] {{ color: {THEME.text_muted}; }}

QFrame#card {{
    background: {THEME.surface};
    border: 1px solid {THEME.border};
    border-radius: {THEME.radius}px;
}}
QFrame#cardMuted {{
    background: {THEME.surface_muted};
    border: 1px solid {THEME.border};
    border-radius: {THEME.radius}px;
}}

QFrame#banner       {{ background: {THEME.primary_soft}; border: 1px solid #cfe0ec; border-radius: {THEME.radius_sm}px; }}
QFrame#bannerOk     {{ background: {THEME.ok_soft};      border: 1px solid #bfe3ce; border-radius: {THEME.radius_sm}px; }}
QFrame#bannerWarn   {{ background: {THEME.warning_soft}; border: 1px solid #ecd9a8; border-radius: {THEME.radius_sm}px; }}
QFrame#bannerError  {{ background: {THEME.danger_soft};  border: 1px solid #f0c4bf; border-radius: {THEME.radius_sm}px; }}

QPushButton {{
    background: {THEME.surface};
    border: 1px solid {THEME.border_strong};
    border-radius: 8px;
    padding: 9px 18px;
    font-size: 14px;
}}
QPushButton:hover    {{ border-color: {THEME.text_faint}; }}
QPushButton:disabled {{ color: {THEME.text_faint}; background: {THEME.surface_muted}; border-color: {THEME.border}; }}

QPushButton[variant="primary"] {{
    background: {THEME.primary};
    border: 1px solid {THEME.primary};
    color: #ffffff;
    font-size: 14px;
    font-weight: 600;
    padding: 11px 22px;
}}
QPushButton[variant="primary"]:hover    {{ background: {THEME.primary_dark}; border-color: {THEME.primary_dark}; }}
QPushButton[variant="primary"]:disabled {{ background: #b9c7d2; border-color: #b9c7d2; color: #eef2f6; }}

QPushButton[variant="ghost"] {{ background: transparent; border: none; color: {THEME.primary}; padding: 9px 12px; }}
QPushButton[variant="ghost"]:hover {{ color: {THEME.primary_dark}; text-decoration: underline; }}

QLabel[badge="ok"]      {{ background: {THEME.ok_soft};      color: {THEME.ok};      border: 1px solid #bfe3ce; border-radius: 4px; padding: 2px 9px; font-size: 12px; font-weight: 600; }}
QLabel[badge="warning"] {{ background: {THEME.warning_soft}; color: {THEME.warning}; border: 1px solid #ecd9a8; border-radius: 4px; padding: 2px 9px; font-size: 12px; font-weight: 600; }}
QLabel[badge="pending"] {{ background: {THEME.surface_muted}; color: {THEME.text_faint}; border: 1px solid {THEME.border}; border-radius: 4px; padding: 2px 9px; font-size: 12px; font-weight: 600; }}
QLabel[badge="info"]    {{ background: {THEME.primary_soft};  color: {THEME.primary_dark}; border: 1px solid #cfe0ec; border-radius: 4px; padding: 2px 9px; font-size: 12px; font-weight: 700; letter-spacing: 0.5px; }}

QComboBox, QSpinBox {{
    background: {THEME.surface};
    border: 1px solid {THEME.border_strong};
    border-radius: 8px;
    padding: 7px 10px;
    min-height: 20px;
}}
QComboBox:focus, QSpinBox:focus {{ border-color: {THEME.primary}; }}

QProgressBar {{
    border: 1px solid {THEME.border};
    border-radius: 6px;
    background: {THEME.surface_muted};
    height: 8px;
    text-align: center;
}}
QProgressBar::chunk {{ background: {THEME.primary}; border-radius: 6px; }}

QLabel#stepActive   {{ background: {THEME.primary};      color: #ffffff; border-radius: 13px; padding: 4px 12px; font-size: 14px; font-weight: 600; }}
QLabel#stepInactive {{ background: {THEME.surface_muted}; color: {THEME.text_faint}; border: 1px solid {THEME.border}; border-radius: 13px; padding: 4px 12px; font-size: 14px; }}
QLabel#stepDone     {{ background: {THEME.ok_soft};      color: {THEME.ok};      border-radius: 13px; padding: 4px 12px; font-size: 14px; font-weight: 600; }}
"""


def refresh_style(widget: QWidget) -> None:
    """Re-aplica el QSS tras cambiar una propiedad dinámica en runtime."""
    style = widget.style()
    style.unpolish(widget)
    style.polish(widget)
    widget.update()


def format_int(value: int) -> str:
    """Miles con punto, estilo es-CL: 2006 -> '2.006'."""
    return f"{value:,}".replace(",", ".")


# Compatibilidad con imports antiguos.
BG = THEME.bg
SURFACE = THEME.surface
BORDER = THEME.border
TEXT = THEME.text
TEXT_MUTED = THEME.text_muted
ACCENT = THEME.primary
OK = THEME.ok
WARNING = THEME.warning
ERROR = THEME.danger
PENDING = THEME.text_faint
