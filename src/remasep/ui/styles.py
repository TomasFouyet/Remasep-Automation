"""Hoja de estilos central de la maqueta.

Sin fuentes ni paquetes externos: solo QSS sobre los widgets de PySide6.
"""

from __future__ import annotations

from PySide6.QtWidgets import QWidget

# Paleta sobria e institucional.
BG = "#f6f7f9"
SURFACE = "#ffffff"
BORDER = "#e4e7eb"
TEXT = "#1f2328"
TEXT_MUTED = "#6a737d"
ACCENT = "#1f6feb"
ACCENT_DARK = "#1a5fd0"
OK = "#1a7f37"
WARNING = "#9a6700"
WARNING_BG = "#fff8e6"
ERROR = "#cf222e"
PENDING = "#8c959f"

STYLESHEET = f"""
/* Sin 'background' global: dejarlo aquí pintaba una franja gris tras cada QLabel
   dentro de las tarjetas blancas. El fondo de página se pinta solo en los
   contenedores; los QLabel son transparentes salvo clases concretas (badges,
   pasos, banners). */
QWidget {{
    color: {TEXT};
    font-size: 14px;
}}

QMainWindow, QStackedWidget, QStackedWidget > QWidget {{
    background: {BG};
}}

QLabel {{
    background: transparent;
}}

QLabel[role="h1"] {{ font-size: 30px; font-weight: 600; }}
QLabel[role="h2"] {{ font-size: 20px; font-weight: 600; }}
QLabel[role="section"] {{
    font-size: 12px; font-weight: 700; color: {TEXT_MUTED};
    letter-spacing: 1px;
}}
QLabel[role="subtitle"] {{ font-size: 15px; color: {TEXT_MUTED}; }}
QLabel[role="muted"] {{ color: {TEXT_MUTED}; font-size: 12px; }}
QLabel[role="metric"] {{ font-size: 22px; font-weight: 600; }}

QLabel[status="ok"] {{ color: {OK}; }}
QLabel[status="warning"] {{ color: {WARNING}; }}
QLabel[status="error"] {{ color: {ERROR}; }}
QLabel[status="pending"] {{ color: {PENDING}; }}

QFrame#card {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 10px;
}}
QFrame#banner {{
    background: {WARNING_BG};
    border: 1px solid #f0dca0;
    border-radius: 8px;
}}
QFrame#bannerOk {{
    background: #eaf6ec;
    border: 1px solid #b7e0c0;
    border-radius: 8px;
}}

QPushButton {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 9px 18px;
    font-size: 14px;
}}
QPushButton:hover {{ border-color: #c9ced4; }}
QPushButton:disabled {{ color: {TEXT_MUTED}; background: #f0f1f3; }}

QPushButton[variant="primary"] {{
    background: {ACCENT};
    border: 1px solid {ACCENT};
    color: #ffffff;
    font-weight: 600;
    padding: 11px 22px;
}}
QPushButton[variant="primary"]:hover {{ background: {ACCENT_DARK}; border-color: {ACCENT_DARK}; }}
QPushButton[variant="primary"]:disabled {{
    background: #cdd7e6; border-color: #cdd7e6; color: #eef2f8;
}}

QLabel[pill="pending"] {{
    background: {WARNING_BG};
    color: {WARNING};
    border: 1px solid #f0dca0;
    border-radius: 10px;
    padding: 2px 10px;
    font-size: 12px;
    font-weight: 600;
}}

QLabel[badge="real"] {{
    background: #e6effd;
    color: {ACCENT_DARK};
    border: 1px solid #c3d7f7;
    border-radius: 4px;
    padding: 3px 10px;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 1px;
}}
QLabel[badge="demo"] {{
    background: {WARNING_BG};
    color: {WARNING};
    border: 1px solid #f0dca0;
    border-radius: 4px;
    padding: 3px 10px;
    font-size: 12px;
    font-weight: 700;
    letter-spacing: 1px;
}}

QComboBox, QSpinBox {{
    background: {SURFACE};
    border: 1px solid {BORDER};
    border-radius: 8px;
    padding: 7px 10px;
    min-height: 20px;
}}
QComboBox:focus, QSpinBox:focus {{ border-color: {ACCENT}; }}

QLabel#stepActive {{
    background: {ACCENT}; color: #ffffff;
    border-radius: 13px; padding: 4px 12px; font-weight: 600;
}}
QLabel#stepInactive {{
    background: #e9ebee; color: {TEXT_MUTED};
    border-radius: 13px; padding: 4px 12px;
}}
QLabel#stepDone {{
    background: #dff0e3; color: {OK};
    border-radius: 13px; padding: 4px 12px; font-weight: 600;
}}
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
