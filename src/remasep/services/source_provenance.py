"""Modelo de *provenance* de fuente para los cambios del REMASEP (Sprint 3.4).

Cuando comparamos dos versiones del REMASEP oficial (una incompleta y la final)
cada celda que cambia tiene un **origen**:

- ``change_origin`` — ¿el cambio es un **input directo** (celda sin fórmula) o la
  **propagación** de una fórmula cuyo resultado cacheado cambió?
- ``source`` — ¿de qué **fuente / proceso** proviene? (``EGRESOS``,
  ``RESOURCE_CALCULATION``, ``SURGICAL_TABLE``, ``CONTROL_METADATA``,
  ``MEDINET``, ``UNKNOWN_PENDING``, ``MIXED_DERIVED``).
- ``status`` — con qué firmeza se afirma esa fuente
  (``CLIENT_CONFIRMED`` / ``STRUCTURALLY_INFERRED`` /
  ``DERIVED_FROM_DEPENDENCIES`` / ``UNKNOWN``).

**Regla fundamental (spec §10 / §20): no inferir más de lo que la evidencia
permite.** El cliente confirmó grandes trazos (edad/sexo de cirugía → REMASEP B1;
códigos Qx → B2 ANEXO; previsión → REMASEP 01 Quirófano). Todo lo demás que no se
pueda asignar inequívocamente queda ``UNKNOWN_PENDING`` — nunca se "fuerza" una
fuente.

Este módulo es **puro**: no abre workbooks. Recibe el contexto semántico ya
extraído (hoja, section/row/column path, código de procedimiento) y devuelve un
`SourceProvenance`. Así la clasificación es testeable con datos sintéticos.
"""

from __future__ import annotations

import unicodedata
from collections.abc import Sequence
from dataclasses import dataclass, field

# --- fuentes ---------------------------------------------------------------

SRC_EGRESOS = "EGRESOS"
SRC_RESOURCE_CALCULATION = "RESOURCE_CALCULATION"
SRC_SURGICAL_TABLE = "SURGICAL_TABLE"
SRC_CONTROL_METADATA = "CONTROL_METADATA"
SRC_MEDINET = "MEDINET"
SRC_UNKNOWN_PENDING = "UNKNOWN_PENDING"
SRC_MIXED_DERIVED = "MIXED_DERIVED"

ALL_SOURCES = (
    SRC_EGRESOS,
    SRC_RESOURCE_CALCULATION,
    SRC_SURGICAL_TABLE,
    SRC_CONTROL_METADATA,
    SRC_MEDINET,
    SRC_UNKNOWN_PENDING,
    SRC_MIXED_DERIVED,
)

# --- estado de la afirmación --------------------------------------------

ST_CLIENT_CONFIRMED = "CLIENT_CONFIRMED"
ST_STRUCTURALLY_INFERRED = "STRUCTURALLY_INFERRED"
ST_DERIVED_FROM_DEPENDENCIES = "DERIVED_FROM_DEPENDENCIES"
ST_UNKNOWN = "UNKNOWN"

# --- origen del cambio -------------------------------------------------

ORIGIN_DIRECT_INPUT = "DIRECT_INPUT"
ORIGIN_FORMULA_PROPAGATION = "FORMULA_PROPAGATION"
# la propia expresión de la fórmula se reescribió (add / remove / modify): no es
# ni un input directo ni una simple propagación de resultado.
ORIGIN_FORMULA_EXPRESSION_CHANGE = "FORMULA_EXPRESSION_CHANGE"


@dataclass(frozen=True)
class SourceProvenance:
    source: str
    status: str
    change_origin: str
    evidence: tuple[str, ...] = ()
    needs_human_review: bool = False

    def evidence_summary(self) -> str:
        return " | ".join(self.evidence)


# ---------------------------------------------------------------------------
# Utilidades de texto
# ---------------------------------------------------------------------------


def _norm(value: object) -> str:
    """Mayúsculas, sin acentos, espacios colapsados — para *keyword matching*."""
    raw = unicodedata.normalize("NFD", str(value or "").upper())
    raw = "".join(ch for ch in raw if unicodedata.category(ch) != "Mn")
    return " ".join(raw.split())


def _joined(*parts: Sequence[str] | str | None) -> str:
    chunks: list[str] = []
    for part in parts:
        if part is None:
            continue
        if isinstance(part, str):
            chunks.append(part)
        else:
            chunks.extend(str(p) for p in part)
    return _norm(" :: ".join(chunks))


# ---------------------------------------------------------------------------
# Clasificación de un input DIRECTO
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class DeltaContext:
    """Contexto semántico mínimo de una celda que cambió."""

    sheet: str
    section_path: tuple[str, ...] = ()
    row_path: tuple[str, ...] = ()
    column_path: tuple[str, ...] = ()
    procedure_code: str = ""
    control_reason: str = ""
    control_sheet_without_data: str = ""
    extra_evidence: tuple[str, ...] = field(default_factory=tuple)


def classify_direct(ctx: DeltaContext) -> SourceProvenance:
    """Clasifica un ``DIRECT_INPUT`` según hoja + estructura del formulario.

    Conservador por diseño: sólo se afirma ``CLIENT_CONFIRMED`` donde el cliente
    lo confirmó explícitamente; si la columna de la Sección D no puede asignarse
    inequívocamente entre recursos y tabla quirúrgica, o si es la Sección E de
    suspensiones, queda ``UNKNOWN_PENDING``.
    """
    sheet = _norm(ctx.sheet)
    section = _joined(ctx.section_path)
    columns = _joined(ctx.column_path)
    rows = _joined(ctx.row_path)
    haystack = f"{section} :: {rows} :: {columns}"

    # --- CONTROL: causales de hojas sin datos -------------------------
    if sheet == "CONTROL":
        ev = ["hoja CONTROL: causal de hoja sin datos"]
        if ctx.control_sheet_without_data:
            ev.append(f"hoja sin datos = {ctx.control_sheet_without_data}")
        if ctx.control_reason:
            ev.append(f"causal = {ctx.control_reason}")
        return SourceProvenance(
            SRC_CONTROL_METADATA, ST_STRUCTURALLY_INFERRED, ORIGIN_DIRECT_INPUT,
            tuple(ev), needs_human_review=False,
        )

    # --- B2 ANEXO: códigos de intervención quirúrgica -> EGRESOS ------
    if sheet in {"B2 ANEXO", "B2_ANEXO"}:
        is_surgical = ctx.procedure_code != "" or "INTERVENCIONES QUIRURGICAS" in haystack
        if is_surgical:
            ev = ["cliente: códigos de intervenciones quirúrgicas -> B2 ANEXO"]
            if ctx.procedure_code:
                ev.append(f"procedure_code = {ctx.procedure_code}")
            return SourceProvenance(
                SRC_EGRESOS, ST_CLIENT_CONFIRMED, ORIGIN_DIRECT_INPUT,
                tuple(ev), needs_human_review=False,
            )
        return SourceProvenance(
            SRC_UNKNOWN_PENDING, ST_UNKNOWN, ORIGIN_DIRECT_INPUT,
            ("B2 ANEXO: fila sin código de procedimiento ni contexto quirúrgico",),
            needs_human_review=True,
        )

    # --- REMASEP B1: grilla de cirugías por edad / sexo -> EGRESOS ----
    if sheet in {"REMASEP B1", "REMASEP_B1"}:
        by_age_sex = (
            "GRUPO DE EDAD" in haystack
            or "POR SEXO" in haystack
            or "INTERVENCIONES QUIRURGICAS" in section
        )
        if by_age_sex:
            return SourceProvenance(
                SRC_EGRESOS, ST_CLIENT_CONFIRMED, ORIGIN_DIRECT_INPUT,
                ("cliente: edad + sexo de cirugía -> REMASEP B1",),
                needs_human_review=False,
            )
        return SourceProvenance(
            SRC_UNKNOWN_PENDING, ST_UNKNOWN, ORIGIN_DIRECT_INPUT,
            ("REMASEP B1: input directo fuera de la grilla edad/sexo confirmada",),
            needs_human_review=True,
        )

    # --- REMASEP 01 --------------------------------------------------
    if sheet in {"REMASEP 01", "REMASEP_01"}:
        if "SECCION E" in section or "CAUSAS DE SUSPENSION" in haystack:
            # Sección E: el cliente NO indicó la fuente. No asumir tabla quirúrgica.
            return SourceProvenance(
                SRC_UNKNOWN_PENDING, ST_UNKNOWN, ORIGIN_DIRECT_INPUT,
                ("REMASEP 01 Sección E (causas de suspensión): fuente no confirmada por el cliente",),
                needs_human_review=True,
            )
        if "SECCION D" in section or "CAPACIDAD INSTALADA Y UTILIZACION DE QUIROFANOS" in haystack:
            return _classify_section_d(columns)
        return SourceProvenance(
            SRC_UNKNOWN_PENDING, ST_UNKNOWN, ORIGIN_DIRECT_INPUT,
            ("REMASEP 01: input directo fuera de las secciones D/E analizadas",),
            needs_human_review=True,
        )

    return SourceProvenance(
        SRC_UNKNOWN_PENDING, ST_UNKNOWN, ORIGIN_DIRECT_INPUT,
        (f"hoja '{ctx.sheet}' sin regla de provenance",),
        needs_human_review=True,
    )


def _classify_section_d(columns_norm: str) -> SourceProvenance:
    """Sección D de REMASEP 01 — capacidad vs. utilización de quirófanos.

    El cliente confirmó el *proceso* (capacidad = días hábiles × 7–8 h/día;
    ocupación se revisa desde la tabla quirúrgica) pero **no** la regla por
    columna. Se usa el texto de los encabezados; lo que no calza queda
    ``UNKNOWN_PENDING``. Nada se marca ``CLIENT_CONFIRMED``.
    """
    occ = any(k in columns_norm for k in ("TABLA QUIRURGICA", "OCUPAD", "PROGRAMADAS"))
    cap = any(k in columns_norm for k in ("DOTACION", "HABILITAD", "EN TRABAJO"))
    if occ:
        ev = f"encabezado de columna indica tabla quirúrgica / horas ocupadas: '{columns_norm}'"
        return SourceProvenance(
            SRC_SURGICAL_TABLE, ST_STRUCTURALLY_INFERRED, ORIGIN_DIRECT_INPUT,
            (ev,), needs_human_review=True,
        )
    if cap:
        ev = (
            f"encabezado de columna indica capacidad instalada "
            f"(dotación / habilitados / horas hábiles): '{columns_norm}'"
        )
        return SourceProvenance(
            SRC_RESOURCE_CALCULATION, ST_STRUCTURALLY_INFERRED, ORIGIN_DIRECT_INPUT,
            (ev,), needs_human_review=True,
        )
    ev = (
        f"Sección D: columna no asignable inequívocamente entre recursos y tabla "
        f"quirúrgica: '{columns_norm}'"
    )
    return SourceProvenance(
        SRC_UNKNOWN_PENDING, ST_UNKNOWN, ORIGIN_DIRECT_INPUT, (ev,), needs_human_review=True,
    )


# ---------------------------------------------------------------------------
# Propagación a un cambio DERIVADO (resultado de fórmula)
# ---------------------------------------------------------------------------


def _base_sources(items: Sequence[SourceProvenance]) -> set[str]:
    """Fuentes 'de base' presentes aguas arriba (``MIXED_DERIVED`` se aplana)."""
    out: set[str] = set()
    for prov in items:
        if prov.source == SRC_MIXED_DERIVED:
            # una entrada ya mezclada aporta su propia evidencia de mezcla
            out.update(
                s for s in ALL_SOURCES
                if s not in (SRC_MIXED_DERIVED,) and s in prov.evidence_summary()
            )
            out.add(SRC_MIXED_DERIVED)
        else:
            out.add(prov.source)
    return out


def unknown_expression_change(before_formula: str, after_formula: str) -> SourceProvenance:
    """Provenance de una celda cuya **expresión de fórmula** se reescribió.

    Un cambio de expresión no puede tratarse como input manual ni como
    propagación: sin evidencia explícita, queda ``UNKNOWN_PENDING`` /
    ``needs_human_review``. Se registra el antes/después de la fórmula.
    """
    kind = (
        "MODIFIED" if before_formula and after_formula
        else "ADDED" if after_formula
        else "REMOVED"
    )
    ev = (
        f"expresión de fórmula {kind}: before={before_formula!r} after={after_formula!r} "
        f"— no es input directo ni propagación de resultado"
    )
    return SourceProvenance(
        SRC_UNKNOWN_PENDING, ST_UNKNOWN, ORIGIN_FORMULA_EXPRESSION_CHANGE,
        (ev,), needs_human_review=True,
    )


def combine_derived(
    upstream: Sequence[SourceProvenance],
    *,
    change_origin: str = ORIGIN_FORMULA_PROPAGATION,
) -> SourceProvenance:
    """Provenance de una fórmula a partir de sus celdas *cambiadas* aguas arriba.

    - sin upstream trazable            -> ``UNKNOWN_PENDING``
    - exactamente una fuente de base   -> esa fuente, ``DERIVED_FROM_DEPENDENCIES``
    - varias fuentes distintas         -> ``MIXED_DERIVED``
    """
    if not upstream:
        return SourceProvenance(
            SRC_UNKNOWN_PENDING, ST_UNKNOWN, change_origin,
            ("cambio de resultado de fórmula sin celda cambiada trazable aguas arriba",),
            needs_human_review=True,
        )

    sources = _base_sources(upstream)
    mixed = SRC_MIXED_DERIVED in sources
    sources.discard(SRC_MIXED_DERIVED)
    needs_review = mixed or any(p.needs_human_review for p in upstream) or (
        SRC_UNKNOWN_PENDING in sources
    )

    if mixed or len(sources) > 1:
        detail = ", ".join(sorted(sources)) or "múltiples"
        return SourceProvenance(
            SRC_MIXED_DERIVED, ST_DERIVED_FROM_DEPENDENCIES, change_origin,
            (f"fórmula alimentada por varias fuentes: {detail}",),
            needs_human_review=needs_review,
        )

    (only,) = tuple(sources)
    return SourceProvenance(
        only,
        ST_DERIVED_FROM_DEPENDENCIES,
        change_origin,
        (f"propagado desde {len(upstream)} celda(s) cambiada(s) de fuente {only}",),
        needs_human_review=needs_review,
    )
