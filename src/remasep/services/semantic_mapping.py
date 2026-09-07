"""Modelo semántico explícito y auditable de las métricas MEDINET (Sprint 3.2).

Toma el *contexto preliminar* del inventario de Sprint 3.1 (rótulos de fila /
columna / sección + evidencia de fórmula) y lo transforma en dimensiones
**explícitas** — sexo, edad, alcance de agregación, código de procedimiento —
**sólo cuando hay evidencia suficiente**.

Regla fundamental: **no inventar semántica**. Cada dimensión inferida conserva
`value`, `status` y `evidence`. Estados por dimensión:

    CONFIRMED · LABEL_ONLY · FORMULA_ONLY · NOT_APPLICABLE · UNRESOLVED · CONFLICT

Vocabulario explícito (rótulos de sexo, hojas → formulario, reglas de código de
procedimiento) vive versionado en ``config/semantic_mapping_2026/`` con
``status: preliminary_pending_functional_validation`` — **no** es vocabulario
oficial MINSAL.

Este mapping **no** es todavía un mapping a la plantilla oficial, ni una
validación funcional por el cliente.
"""

from __future__ import annotations

import re
from dataclasses import dataclass, field
from pathlib import Path

import yaml

from remasep.core.errors import RemasepError
from remasep.services.semantic_layout import (
    age_intervals_overlap,
    is_group_header_label,
    normalize_semantic_label,
)

# --- estados de dimensión --------------------------------------------

CONFIRMED = "CONFIRMED"
LABEL_ONLY = "LABEL_ONLY"
FORMULA_ONLY = "FORMULA_ONLY"
NOT_APPLICABLE = "NOT_APPLICABLE"
UNRESOLVED = "UNRESOLVED"
CONFLICT = "CONFLICT"

# --- estado de mapeo global (distinto de context_status de Sprint 3.1) ---

MAPPING_CONFIRMED = "CONFIRMED"
MAPPING_PARTIAL = "PARTIAL"
MAPPING_CONFLICT = "CONFLICT"

# --- valores internos --------------------------------------------------

SOURCE_MEDINET = "MEDINET"
SEX_MALE = "MALE"
SEX_FEMALE = "FEMALE"
SEX_BOTH = "BOTH"

SCOPE_DETAIL = "DETAIL"
SCOPE_TOTAL = "TOTAL"
SCOPE_SUBTOTAL = "SUBTOTAL"

PROC_EXPLICIT = "EXPLICIT"

_DEFAULT_CONFIG_DIR = Path(__file__).resolve().parents[3] / "config" / "semantic_mapping_2026"


class SemanticMappingError(RemasepError):
    """Configuración de mapeo semántico inválida."""


# ---------------------------------------------------------------------------
# Configuración versionada
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class MappingConfig:
    version: str
    sex_by_label: dict[str, str]  # rótulo normalizado -> MALE/FEMALE/BOTH
    sex_by_criterion: dict[str, str]  # substring minúscula -> MALE/FEMALE
    form_by_sheet: dict[str, str]
    procedure_min_digits: int
    procedure_max_digits: int


def _load_yaml(path: Path) -> dict:
    if not path.is_file():
        raise SemanticMappingError(f"config de mapeo no encontrada: {path}")
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def load_mapping_config(config_dir: str | Path | None = None) -> MappingConfig:
    base = Path(config_dir) if config_dir is not None else _DEFAULT_CONFIG_DIR
    sex_raw = _load_yaml(base / "sex_labels.yaml")
    form_raw = _load_yaml(base / "form_labels.yaml")
    proc_raw = _load_yaml(base / "procedure_codes.yaml")

    sex_by_label: dict[str, str] = {}
    for internal, labels in (sex_raw.get("labels") or {}).items():
        for label in labels:
            sex_by_label[normalize_semantic_label(label)] = internal
    sex_by_criterion: dict[str, str] = {}
    for internal, subs in (sex_raw.get("formula_criteria") or {}).items():
        for sub in subs:
            sex_by_criterion[str(sub).lower()] = internal

    form_by_sheet = {str(k): str(v) for k, v in (form_raw.get("forms") or {}).items()}

    return MappingConfig(
        version=str(sex_raw.get("version", "semantic_mapping_2026")),
        sex_by_label=sex_by_label,
        sex_by_criterion=sex_by_criterion,
        form_by_sheet=form_by_sheet,
        procedure_min_digits=int(proc_raw.get("min_digits", 4)),
        procedure_max_digits=int(proc_raw.get("max_digits", 8)),
    )


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class Evidence:
    dimension: str  # FORM / SEX / AGE / AGGREGATION_SCOPE / PROCEDURE
    evidence_type: str  # SHEET / COLUMN_LABEL / ROW_LABEL / FORMULA_CRITERION / FORMULA_BOUNDS
    raw_value: str
    normalized_value: str
    evidence_status: str  # SUPPORTS / CONFLICTS


@dataclass(frozen=True)
class DimensionConflict:
    dimension: str
    label_evidence: str
    formula_evidence: str
    conflict_type: str


@dataclass(frozen=True)
class MetricInput:
    """Lo que Sprint 3.2 necesita de cada métrica del inventario de Sprint 3.1."""

    metric_id: str
    sheet: str
    cell: str
    kind: str
    value: object
    section_path: tuple[str, ...]
    row_path: tuple[str, ...]
    column_path: tuple[str, ...]
    text_criteria: tuple[str, ...]
    formula: str
    formula_age: tuple[int | None, int | None] | None
    label_age: tuple[int | None, int | None] | None
    label_age_text: str
    semantic_signature: str
    context_status: str = ""
    # la celda de la métrica está sobre una fila-encabezado (sección/grupo) del
    # layout de Sprint 3.1 -> es el subtotal de ese encabezado.
    own_row_is_header: bool = False


@dataclass
class SemanticMetric:
    metric_id: str
    source: str
    form: str
    form_raw: str
    sheet: str
    cell: str
    kind: str
    value: object
    section_path_raw: tuple[str, ...]
    row_path_raw: tuple[str, ...]
    column_path_raw: tuple[str, ...]
    sex_value: str
    sex_status: str
    age_min_years: int | None
    age_max_years: int | None
    age_status: str
    aggregation_scope: str
    aggregation_scope_status: str
    procedure_code_raw: str
    procedure_label_raw: str
    procedure_status: str
    mapping_status: str
    semantic_signature: str
    evidence: tuple[Evidence, ...] = ()
    conflicts: tuple[DimensionConflict, ...] = ()


# ---------------------------------------------------------------------------
# Inferencia por dimensión
# ---------------------------------------------------------------------------


def infer_form(sheet: str, config: MappingConfig) -> tuple[str, str]:
    return config.form_by_sheet.get(sheet, sheet.replace(" ", "_")), sheet


def _sex_labels_present(column_path: tuple[str, ...], config: MappingConfig) -> list[tuple[str, str]]:
    hits: list[tuple[str, str]] = []
    for raw in column_path:
        internal = config.sex_by_label.get(normalize_semantic_label(raw))
        if internal:
            hits.append((raw, internal))
    return hits


_QUOTED_LITERAL_RE = re.compile(r'"([^"]*)"')


def _sex_criteria_of(inp: MetricInput, config: MappingConfig) -> list[tuple[str, str]]:
    """Criterios de sexo de la fórmula: literales con comodín del inventario de
    Sprint 3.1 **más** literales entre comillas del texto de la fórmula que
    contengan ``hombre``/``mujer`` (patrón ``$A8&"Hombre"``). No re-parsea la
    fórmula: sólo mira sus literales entre comillas."""
    candidates = list(inp.text_criteria)
    for literal in _QUOTED_LITERAL_RE.findall(inp.formula or ""):
        low = literal.lower()
        if any(sub in low for sub in config.sex_by_criterion):
            candidates.append(literal)
    hits: list[tuple[str, str]] = []
    seen: set[tuple[str, str]] = set()
    for crit in candidates:
        low = crit.lower()
        for sub, internal in config.sex_by_criterion.items():
            if sub in low and (crit, internal) not in seen:
                seen.add((crit, internal))
                hits.append((crit, internal))
    return hits


def infer_sex(
    inp: MetricInput, config: MappingConfig
) -> tuple[str, str, list[Evidence], DimensionConflict | None]:
    label_hits = _sex_labels_present(inp.column_path, config)
    crit_hits = _sex_criteria_of(inp, config)
    label_values = {v for _r, v in label_hits}
    crit_values = {v for _r, v in crit_hits}

    evidence: list[Evidence] = []
    conflict: DimensionConflict | None = None

    # rótulos de sexo internamente contradictorios (raro)
    if len(label_values) > 1:
        for raw, val in label_hits:
            evidence.append(Evidence("SEX", "COLUMN_LABEL", raw, normalize_semantic_label(raw), "CONFLICTS"))
        conflict = DimensionConflict(
            "SEX", " / ".join(sorted(label_values)), "", "SEX_LABELS_DISAGREE"
        )
        return "", CONFLICT, evidence, conflict

    label_value = next(iter(label_values), "")
    formula_value = next(iter(crit_values), "") if len(crit_values) == 1 else ""

    def _label_ev(status: str) -> None:
        for raw, _v in label_hits:
            evidence.append(
                Evidence("SEX", "COLUMN_LABEL", raw, normalize_semantic_label(raw), status)
            )

    def _formula_ev(status: str) -> None:
        for raw, _v in crit_hits:
            evidence.append(Evidence("SEX", "FORMULA_CRITERION", raw, raw.lower(), status))

    if label_value and formula_value:
        if label_value == formula_value:
            _label_ev("SUPPORTS")
            _formula_ev("SUPPORTS")
            return label_value, CONFIRMED, evidence, None
        _label_ev("SUPPORTS")
        _formula_ev("CONFLICTS")
        conflict = DimensionConflict(
            "SEX", label_value, "|".join(sorted(v for _r, v in crit_hits)),
            "SEX_LABEL_FORMULA_MISMATCH",
        )
        return label_value, CONFLICT, evidence, conflict

    if label_value:
        _label_ev("SUPPORTS")
        return label_value, LABEL_ONLY, evidence, None

    if len(crit_values) == 1:
        _formula_ev("SUPPORTS")
        return formula_value, FORMULA_ONLY, evidence, None
    if len(crit_values) > 1:
        _formula_ev("CONFLICTS")
        return "", UNRESOLVED, evidence, None

    return "", NOT_APPLICABLE, evidence, None


def infer_age(
    inp: MetricInput,
) -> tuple[int | None, int | None, str, list[Evidence], DimensionConflict | None]:
    label = inp.label_age
    formula = inp.formula_age
    evidence: list[Evidence] = []

    def _label_ev(status: str) -> None:
        if inp.label_age_text:
            evidence.append(
                Evidence(
                    "AGE", "COLUMN_LABEL", inp.label_age_text,
                    normalize_semantic_label(inp.label_age_text), status,
                )
            )

    def _formula_ev(status: str) -> None:
        if formula is not None:
            lo, hi = formula
            parts = []
            if lo is not None:
                parts.append(f">={lo}")
            if hi is not None:
                parts.append(f"<={hi}")
            evidence.append(
                Evidence("AGE", "FORMULA_BOUNDS", " ".join(parts) or "AF", "", status)
            )

    if label is not None and formula is not None:
        if age_intervals_overlap(label, formula):
            _label_ev("SUPPORTS")
            _formula_ev("SUPPORTS")
            return label[0], label[1], CONFIRMED, evidence, None
        _label_ev("SUPPORTS")
        _formula_ev("CONFLICTS")
        conflict = DimensionConflict(
            "AGE", inp.label_age_text, _bounds_text(formula), "AGE_LABEL_FORMULA_DISJOINT"
        )
        return label[0], label[1], CONFLICT, evidence, conflict

    if label is not None:
        _label_ev("SUPPORTS")
        return label[0], label[1], LABEL_ONLY, evidence, None
    if formula is not None:
        _formula_ev("SUPPORTS")
        return formula[0], formula[1], FORMULA_ONLY, evidence, None

    status = UNRESOLVED if inp.context_status == "AMBIGUOUS" else NOT_APPLICABLE
    return None, None, status, evidence, None


def _bounds_text(bounds: tuple[int | None, int | None]) -> str:
    lo, hi = bounds
    parts = []
    if lo is not None:
        parts.append(f">={lo}")
    if hi is not None:
        parts.append(f"<={hi}")
    return " ".join(parts) or "AF"


def infer_aggregation_scope(
    inp: MetricInput, *, has_age_or_sex: bool, has_procedure: bool
) -> tuple[str, str, list[Evidence]]:
    col_norm = [normalize_semantic_label(x) for x in inp.column_path]
    evidence: list[Evidence] = []

    if "TOTAL" in col_norm:
        raw = inp.column_path[col_norm.index("TOTAL")]
        evidence.append(Evidence("AGGREGATION_SCOPE", "COLUMN_LABEL", raw, "TOTAL", "SUPPORTS"))
        return SCOPE_TOTAL, CONFIRMED, evidence

    on_header = inp.own_row_is_header or (
        len(inp.row_path) == 1 and is_group_header_label(inp.row_path[0])
    )
    if on_header:
        raw = inp.row_path[0] if inp.row_path else ""
        evidence.append(
            Evidence(
                "AGGREGATION_SCOPE", "ROW_LABEL", raw, normalize_semantic_label(raw), "SUPPORTS"
            )
        )
        return SCOPE_SUBTOTAL, CONFIRMED, evidence

    if has_age_or_sex or has_procedure:
        return SCOPE_DETAIL, CONFIRMED, evidence
    return SCOPE_DETAIL, UNRESOLVED, evidence


def extract_procedure(
    row_path: tuple[str, ...], config: MappingConfig
) -> tuple[str, str, str]:
    """``(code_raw, label_raw, source_row_label)`` — código de procedimiento
    EXPLÍCITO. ``("", "", "")`` si no hay ninguno."""
    lo, hi = config.procedure_min_digits, config.procedure_max_digits
    inline = re.compile(rf"^(\d{{{lo},{hi}}})\s*-\s*(.+)$")
    bare = re.compile(rf"^(\d{{{lo},{hi}}})$")
    for i, raw in enumerate(row_path):
        token = raw.strip()
        m = inline.match(token)
        if m:
            return m.group(1), m.group(2).strip(), raw
        b = bare.match(token)
        if b:
            for nxt in row_path[i + 1 :]:
                stripped = nxt.strip()
                if not bare.match(stripped) and not inline.match(stripped):
                    return b.group(1), stripped, raw
            return b.group(1), "", raw
    return "", "", ""


def compute_mapping_status(
    sex_status: str, age_status: str, scope_status: str, has_conflict: bool
) -> str:
    """CONFLICT si alguna dimensión choca; CONFIRMED si sexo y edad están
    ``CONFIRMED`` o justificadamente ``NOT_APPLICABLE`` **y** el alcance de
    agregación está ``CONFIRMED``; ``PARTIAL`` en el resto. Nunca ``COMPLETE``
    (para no confundir con el ``context_status`` de Sprint 3.1)."""
    if has_conflict:
        return MAPPING_CONFLICT
    confirmed_ok = {CONFIRMED, NOT_APPLICABLE}
    if sex_status in confirmed_ok and age_status in confirmed_ok and scope_status == CONFIRMED:
        return MAPPING_CONFIRMED
    return MAPPING_PARTIAL


# ---------------------------------------------------------------------------
# Orquestación por métrica
# ---------------------------------------------------------------------------


def map_metric(inp: MetricInput, config: MappingConfig) -> SemanticMetric:
    form, form_raw = infer_form(inp.sheet, config)

    sex_value, sex_status, sex_ev, sex_conflict = infer_sex(inp, config)
    age_min, age_max, age_status, age_ev, age_conflict = infer_age(inp)

    procedure_code, procedure_label, procedure_row = extract_procedure(inp.row_path, config)
    has_age_or_sex = sex_status in (CONFIRMED, LABEL_ONLY, FORMULA_ONLY) or age_status in (
        CONFIRMED, LABEL_ONLY, FORMULA_ONLY,
    )
    scope_value, scope_status, scope_ev = infer_aggregation_scope(
        inp, has_age_or_sex=has_age_or_sex, has_procedure=bool(procedure_code)
    )

    if procedure_code:
        procedure_status = PROC_EXPLICIT
    elif scope_value in (SCOPE_TOTAL, SCOPE_SUBTOTAL):
        procedure_status = NOT_APPLICABLE
    else:
        procedure_status = UNRESOLVED

    evidence: list[Evidence] = [
        Evidence("FORM", "SHEET", inp.sheet, normalize_semantic_label(inp.sheet), "SUPPORTS"),
        *sex_ev,
        *age_ev,
        *scope_ev,
    ]
    if procedure_code:
        evidence.append(
            Evidence(
                "PROCEDURE", "ROW_LABEL", procedure_row,
                normalize_semantic_label(procedure_row), "SUPPORTS",
            )
        )

    conflicts = tuple(c for c in (sex_conflict, age_conflict) if c is not None)
    mapping_status = compute_mapping_status(
        sex_status, age_status, scope_status, bool(conflicts)
    )

    return SemanticMetric(
        metric_id=inp.metric_id,
        source=SOURCE_MEDINET,
        form=form,
        form_raw=form_raw,
        sheet=inp.sheet,
        cell=inp.cell,
        kind=inp.kind,
        value=inp.value,
        section_path_raw=inp.section_path,
        row_path_raw=inp.row_path,
        column_path_raw=inp.column_path,
        sex_value=sex_value,
        sex_status=sex_status,
        age_min_years=age_min,
        age_max_years=age_max,
        age_status=age_status,
        aggregation_scope=scope_value,
        aggregation_scope_status=scope_status,
        procedure_code_raw=procedure_code,
        procedure_label_raw=procedure_label,
        procedure_status=procedure_status,
        mapping_status=mapping_status,
        semantic_signature=inp.semantic_signature,
        evidence=tuple(evidence),
        conflicts=conflicts,
    )


@dataclass
class MappingResult:
    version: str
    metrics: list[SemanticMetric] = field(default_factory=list)
