"""Traducción de errores/estados del backend a mensajes humanos para la UI.

El detalle técnico (código, tipo de excepción) queda disponible en
``HumanError.technical`` para logs de desarrollo; **no** se muestra al usuario
como texto principal ni se vuelca un traceback.
"""

from __future__ import annotations

import logging
from dataclasses import dataclass

from remasep.core.errors import RemasepError
from remasep.services import excel_writer as _xw
from remasep.services.runtime_assets import (
    RUNTIME_ASSET_INCOMPATIBLE,
    RUNTIME_ASSET_INVALID,
    RUNTIME_ASSET_MISSING,
    RuntimeAssetError,
)

_log = logging.getLogger("remasep.ui")

_ACTION_RETRY_FILE = "Elegir otro archivo"
_ACTION_RETRY_TEMPLATE = "Elegir otra plantilla"
_ACTION_BACK = "Volver"


@dataclass(frozen=True)
class HumanError:
    title: str
    detail: str = ""
    action_label: str = _ACTION_BACK
    technical: str = ""


# Códigos de RuntimeAssetError -> texto humano
_RUNTIME_ASSET_MESSAGES = {
    RUNTIME_ASSET_MISSING: HumanError(
        "La instalación está incompleta.",
        "Faltan archivos internos necesarios para generar el informe. "
        "Reinstala la aplicación o contacta a soporte.",
        _ACTION_BACK,
    ),
    RUNTIME_ASSET_INCOMPATIBLE: HumanError(
        "La instalación no es compatible.",
        "Los archivos internos no corresponden a esta versión del REMASEP.",
        _ACTION_BACK,
    ),
    RUNTIME_ASSET_INVALID: HumanError(
        "Archivos internos dañados.",
        "La instalación contiene archivos internos incompletos o alterados.",
        _ACTION_BACK,
    ),
}

# Estados de generación (excel_writer) -> texto humano
_GENERATION_STATUS_MESSAGES = {
    _xw.STATUS_EXCEL_UNAVAILABLE: HumanError(
        "Microsoft Excel Desktop es necesario.",
        "La generación del REMASEP requiere Microsoft Excel Desktop instalado en "
        "este equipo con Windows.",
        _ACTION_BACK,
    ),
    _xw.STATUS_TEMPLATE_INCOMPATIBLE: HumanError(
        "La plantilla no es compatible.",
        "La plantilla seleccionada no corresponde a la versión del REMASEP "
        "compatible con la aplicación.",
        _ACTION_RETRY_TEMPLATE,
    ),
    _xw.STATUS_OUTPUT_EXISTS: HumanError(
        "Ya existe un informe con ese nombre.",
        "Elige otro nombre o guarda el informe en otra carpeta.",
        "Elegir otro nombre",
    ),
    _xw.STATUS_ABORTED_PREFLIGHT: HumanError(
        "No pudimos preparar la generación.",
        "Revisa el período, el archivo de Medinet y la plantilla seleccionada.",
        _ACTION_BACK,
    ),
    _xw.STATUS_FAILED_INTEGRITY_CHECK: HumanError(
        "El archivo generado no pasó la verificación.",
        "La verificación de integridad detectó un problema y el informe no se "
        "considera válido. Vuelve a intentarlo.",
        "Reintentar",
    ),
    _xw.STATUS_GENERATION_FAILED: HumanError(
        "No pudimos generar el informe.",
        "Ocurrió un problema durante la generación. Vuelve a intentarlo.",
        "Reintentar",
    ),
}

# Fragmentos conocidos dentro del mensaje de un RemasepError genérico
_MESSAGE_FRAGMENTS: tuple[tuple[str, HumanError], ...] = (
    (
        "no existe el archivo",
        HumanError(
            "No encontramos el archivo seleccionado.",
            "Verifica que el archivo siga en su ubicación y vuelve a elegirlo.",
            _ACTION_RETRY_FILE,
        ),
    ),
    (
        "no expone columnas",
        HumanError(
            "El archivo de Medinet no tiene el formato esperado.",
            "Parece que no es un export de “Detalle de citas” de Medinet. "
            "Elige el archivo correcto.",
            _ACTION_RETRY_FILE,
        ),
    ),
    (
        "no se pudo leer",
        HumanError(
            "No pudimos leer el archivo de Medinet.",
            "El archivo puede estar dañado o abierto en otra aplicación. Ciérralo "
            "y vuelve a intentarlo.",
            _ACTION_RETRY_FILE,
        ),
    ),
    (
        "período inválido",
        HumanError(
            "El período no es válido.",
            "Selecciona un mes y un año correctos.",
            _ACTION_BACK,
        ),
    ),
    (
        "no hay registros",
        HumanError(
            "El archivo no tiene atenciones para el período elegido.",
            "El archivo seleccionado no parece corresponder al mes y año elegidos.",
            _ACTION_RETRY_FILE,
        ),
    ),
)

_FALLBACK_MEDINET = HumanError(
    "No pudimos procesar el archivo de Medinet.",
    "Revisa que el archivo sea el export “Detalle de citas” del período "
    "elegido y vuelve a intentarlo.",
    _ACTION_RETRY_FILE,
)


def humanize_generation_status(status: str, *, errors: list[str] | None = None) -> HumanError:
    """Traduce un estado de ``GenerationService`` a un :class:`HumanError`."""
    base = _GENERATION_STATUS_MESSAGES.get(status)
    technical = f"status={status}" + (f"; {' | '.join(errors)}" if errors else "")
    if base is None:
        base = HumanError(
            "No pudimos generar el informe.",
            "Ocurrió un problema inesperado durante la generación.",
            "Reintentar",
        )
    _log.warning("generation status not OK: %s", technical)
    return HumanError(base.title, base.detail, base.action_label, technical)


def humanize_error(exc: BaseException, *, fallback: HumanError | None = None) -> HumanError:
    """Traduce una excepción del backend a un :class:`HumanError`.

    Nunca expone el traceback; el detalle técnico va a ``.technical`` y al log.
    """
    technical = f"{exc.__class__.__name__}: {exc}"
    _log.warning("UI caught backend error: %s", technical)

    if isinstance(exc, RuntimeAssetError):
        base = _RUNTIME_ASSET_MESSAGES.get(
            exc.code,
            HumanError(
                "La instalación tiene un problema.",
                "No se pudieron cargar los archivos internos necesarios.",
                _ACTION_BACK,
            ),
        )
        return HumanError(base.title, base.detail, base.action_label, technical)

    if isinstance(exc, _xw.ExcelWriterError):
        text = str(exc)
        if "TEMPLATE_INCOMPATIBLE" in text:
            base = _GENERATION_STATUS_MESSAGES[_xw.STATUS_TEMPLATE_INCOMPATIBLE]
            return HumanError(base.title, base.detail, base.action_label, technical)
        base = _GENERATION_STATUS_MESSAGES[_xw.STATUS_GENERATION_FAILED]
        return HumanError(base.title, base.detail, base.action_label, technical)

    if isinstance(exc, RemasepError):
        low = str(exc).lower()
        for fragment, human in _MESSAGE_FRAGMENTS:
            if fragment in low:
                return HumanError(human.title, human.detail, human.action_label, technical)
        base = fallback or _FALLBACK_MEDINET
        return HumanError(base.title, base.detail, base.action_label, technical)

    base = fallback or HumanError(
        "Ocurrió un problema inesperado.",
        "Vuelve a intentarlo. Si el problema persiste, contacta a soporte.",
        _ACTION_BACK,
    )
    return HumanError(base.title, base.detail, base.action_label, technical)
