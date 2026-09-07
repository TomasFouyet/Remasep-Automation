"""Alineación semántica entre métricas MEDINET y la plantilla oficial (Sprint 3.5).

`SemanticMetric` MEDINET (Sprints 3.1–3.3) describe *qué* mide cada celda del
workbook **generador** legacy: forma, sección, fila, columna, sexo, edad, alcance
de agregación, código de procedimiento, firma semántica.

La plantilla oficial REMASEP tiene **otro layout** (otras coordenadas, otra
estructura de filas/columnas). Este módulo intenta emparejar cada métrica
elegible con la(s) celda(s) de la plantilla que representan **la misma cosa**,
usando **evidencia por dimensión** — nunca la coordenada como única prueba.

Salidas por par: `match_status` ∈ ``EXACT_SEMANTIC_MATCH · STRONG_MATCH ·
AMBIGUOUS · NO_MATCH · CONFLICT``, con el desglose de score.

**No** escribe Excel, **no** usa COM, **no** resuelve EGRESOS ni
RESOURCE_CALCULATION, **no** toca el mapeo semántico de Sprint 3.2, **no** llama
al resultado "MINSAL validated". Reglas versionadas en
``config/official_template_alignment_2026/`` (``status:
preliminary_pending_functional_validation``).
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

from remasep.core.errors import RemasepError
from remasep.services.semantic_layout import (
    SheetLayout,
    normalize_semantic_label,
    parse_age_label,
)

# --- clasificación FÍSICA de la celda target (qué es la celda) -------

DIRECT_INPUT_TARGET = "DIRECT_INPUT_TARGET"
FORMULA_TARGET = "FORMULA_TARGET"
STRUCTURAL = "STRUCTURAL"
VALIDATION = "VALIDATION"
UNKNOWN_TARGET = "UNKNOWN"

# --- ROL de la celda en la alineación (para qué sirve) --------------
# Distinto de `target_source_expectation` (de qué fuente proviene).

ROLE_INPUT_TARGET = "INPUT_TARGET"        # celda de ingreso: puede recibir un mapping
ROLE_DERIVED_TARGET = "DERIVED_TARGET"    # fórmula: depende de inputs, no es punto de ingreso
ROLE_STRUCTURAL = "STRUCTURAL"            # rótulo / código: nunca es punto de alineación
ROLE_VALIDATION = "VALIDATION"            # celda de control / validación
ROLE_UNKNOWN = "UNKNOWN"

_ROLE_BY_KIND = {
    DIRECT_INPUT_TARGET: ROLE_INPUT_TARGET,
    FORMULA_TARGET: ROLE_DERIVED_TARGET,
    STRUCTURAL: ROLE_STRUCTURAL,
    VALIDATION: ROLE_VALIDATION,
    UNKNOWN_TARGET: ROLE_UNKNOWN,
}


def alignment_role(target_kind: str) -> str:
    return _ROLE_BY_KIND.get(target_kind, ROLE_UNKNOWN)

# --- estado del match ---------------------------------------------

EXACT_SEMANTIC_MATCH = "EXACT_SEMANTIC_MATCH"
STRONG_MATCH = "STRONG_MATCH"
AMBIGUOUS = "AMBIGUOUS"
NO_MATCH = "NO_MATCH"
CONFLICT = "CONFLICT"

# --- expectativa de fuente del target (Sprint 3.4) ---------------

EXP_MEDINET = "MEDINET"
EXP_EGRESOS = "EGRESOS"
EXP_RESOURCE_CALCULATION = "RESOURCE_CALCULATION"
EXP_SURGICAL_TABLE = "SURGICAL_TABLE"
EXP_CONTROL_METADATA = "CONTROL_METADATA"
EXP_UNKNOWN = "UNKNOWN"

# Fuentes explícitamente NO-MEDINET: si el mejor target de una métrica MEDINET cae
# en una de estas regiones, no se empareja (no se mezclan fuentes). Una
# expectativa ``UNKNOWN`` NO bloquea: un match fuerte de un source MEDINET es en
# sí evidencia de que el target pertenece al subgrafo MEDINET.
_NON_MEDINET_EXPECTATIONS = (
    EXP_EGRESOS, EXP_RESOURCE_CALCULATION, EXP_SURGICAL_TABLE, EXP_CONTROL_METADATA,
)

# --- evidencia por dimensión ------------------------------------

EV_MATCH = "MATCH"
EV_PARTIAL = "PARTIAL"
EV_MISMATCH = "MISMATCH"
EV_NOT_APPLICABLE = "NOT_APPLICABLE"
EV_MISSING = "MISSING"

_DIMENSIONS = (
    "FORM", "SECTION", "ROW_PATH", "COLUMN_PATH",
    "SEX", "AGE", "AGGREGATION_SCOPE", "PROCEDURE_CODE",
)
_CONFLICTING_DIMENSIONS = ("SEX", "AGE", "PROCEDURE_CODE")

_DEFAULT_CONFIG_DIR = (
    Path(__file__).resolve().parents[3] / "config" / "official_template_alignment_2026"
)

_SEX_TOKENS = {
    "HOMBRES": "MALE", "HOMBRE": "MALE", "MASCULINO": "MALE", "VARONES": "MALE",
    "MUJERES": "FEMALE", "MUJER": "FEMALE", "FEMENINO": "FEMALE",
    "AMBOS SEXO": "BOTH", "AMBOS SEXOS": "BOTH", "AMBOS": "BOTH", "TOTAL": "",
}
_SEX_LABEL_NORMS = set(_SEX_TOKENS) | {"AMBOS SEXO", "AMBOS SEXOS"}
_SCOPE_TOTAL = "TOTAL"
_SCOPE_SUBTOTAL = "SUBTOTAL"
_SCOPE_DETAIL = "DETAIL"


class TemplateAlignmentError(RemasepError):
    """Configuración de alineación inválida."""


# ---------------------------------------------------------------------------
# Configuración versionada
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class AlignmentPolicy:
    version: str
    weights: dict[str, float]
    exact_threshold: float
    strong_threshold: float
    tie_margin: float
    form_by_sheet: dict[str, str]
    required_dimensions: dict[str, tuple[str, ...]]

    def weight(self, dimension: str) -> float:
        return float(self.weights.get(dimension, 0.0))

    def canonical_form(self, sheet: str) -> str:
        norm = normalize_semantic_label(sheet)
        return self.form_by_sheet.get(norm, norm.replace(" ", "_"))


@dataclass(frozen=True)
class SourceRegion:
    """Región de la plantilla con una fuente esperada declarada (Sprint 3.4/3.5).

    Se identifica por evidencia **semántica** (uno o más rótulos de sección) y/o,
    para el workbook de referencia actual, por un rango de filas documentado y
    testeado. ``expected_source`` puede ser NO-MEDINET (``EGRESOS`` / recursos /
    tabla quirúrgica) **o** ``MEDINET`` (evidencia estructural positiva de que el
    target pertenece al subgrafo Medinet).
    """

    region_id: str
    sheet: str
    expected_source: str
    section_contains: tuple[str, ...]
    row_min: int | None
    row_max: int | None
    evidence: str

    def matches(self, sheet: str, row: int, section_norm: tuple[str, ...]) -> bool:
        if normalize_semantic_label(sheet) != normalize_semantic_label(self.sheet):
            return False
        has_rows = self.row_min is not None or self.row_max is not None
        if has_rows:
            # rango de filas autoritativo (evidencia de coordenada explícita); la
            # sección declarada queda como documentación del bloque.
            if self.row_min is not None and row < self.row_min:
                return False
            return not (self.row_max is not None and row > self.row_max)
        if self.section_contains:
            needles = [normalize_semantic_label(n) for n in self.section_contains]
            return any(n in s for n in needles for s in section_norm)
        return True  # región de hoja completa


def _load_yaml(path: Path) -> dict:
    if not path.is_file():
        raise TemplateAlignmentError(f"config de alineación no encontrada: {path}")
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def load_alignment_policy(config_dir: str | Path | None = None) -> AlignmentPolicy:
    base = Path(config_dir) if config_dir is not None else _DEFAULT_CONFIG_DIR
    raw = _load_yaml(base / "policy.yaml")
    scoring = raw.get("scoring") or {}
    thresholds = raw.get("thresholds") or {}
    return AlignmentPolicy(
        version=str(raw.get("version", "official_template_alignment_2026")),
        weights={str(k): float(v) for k, v in (scoring.get("weights") or {}).items()},
        exact_threshold=float(thresholds.get("exact", 1.0)),
        strong_threshold=float(thresholds.get("strong", 0.75)),
        tie_margin=float(thresholds.get("tie_margin", 0.05)),
        form_by_sheet={
            normalize_semantic_label(k): str(v)
            for k, v in (raw.get("form_by_sheet") or {}).items()
        },
        required_dimensions={
            str(form): tuple(dims)
            for form, dims in (raw.get("required_dimensions") or {}).items()
        },
    )


def load_source_regions(config_dir: str | Path | None = None) -> list[SourceRegion]:
    base = Path(config_dir) if config_dir is not None else _DEFAULT_CONFIG_DIR
    raw = _load_yaml(base / "source_regions.yaml")
    out: list[SourceRegion] = []
    for entry in raw.get("regions") or []:
        raw_sections = entry.get("section_contains", "")
        if isinstance(raw_sections, str):
            sections = (raw_sections,) if raw_sections else ()
        else:
            sections = tuple(str(s) for s in raw_sections if str(s))
        out.append(
            SourceRegion(
                region_id=str(entry.get("region_id", "")),
                sheet=str(entry.get("sheet", "")),
                expected_source=str(entry.get("expected_source", EXP_UNKNOWN)),
                section_contains=sections,
                row_min=entry.get("row_min"),
                row_max=entry.get("row_max"),
                evidence=str(entry.get("evidence", "")),
            )
        )
    return out


def region_expectation(
    regions: list[SourceRegion], sheet: str, row: int, section_norm: tuple[str, ...]
) -> tuple[str, str]:
    """(expected_source, evidence).

    Devuelve la fuente de la **primera** región que reclama la celda (las
    NO-MEDINET, más específicas, se listan antes que las MEDINET amplias). Si
    ninguna región aplica -> ``UNKNOWN``: no se inventa MEDINET sin evidencia
    estructural.
    """
    for region in regions:
        if region.matches(sheet, row, section_norm):
            return region.expected_source, f"{region.region_id}: {region.evidence}"
    return EXP_UNKNOWN, "sin región de fuente declarada"


# ---------------------------------------------------------------------------
# Inventario semántico de la plantilla (target)
# ---------------------------------------------------------------------------


@dataclass
class TargetMetricContext:
    target_sheet: str
    target_cell: str
    form: str
    row: int
    column: int
    section_path: tuple[str, ...]
    row_path: tuple[str, ...]
    column_path: tuple[str, ...]
    section_norm: tuple[str, ...]
    row_norm: tuple[str, ...]
    column_norm: tuple[str, ...]
    column_structure_norm: tuple[str, ...]
    target_kind: str
    target_alignment_role: str
    has_formula: bool
    formula: str
    is_blank: bool
    sex_value: str
    age_min: int | None
    age_max: int | None
    aggregation_scope: str
    procedure_code: str
    semantic_signature: str
    target_source_expectation: str
    target_source_evidence: str

    @property
    def key(self) -> tuple[str, str]:
        return (self.target_sheet, self.target_cell)


def _col_letter(index: int) -> str:
    letters = ""
    while index > 0:
        index, rem = divmod(index - 1, 26)
        letters = chr(65 + rem) + letters
    return letters


def sex_from_labels(labels_norm: tuple[str, ...]) -> str:
    for label in labels_norm:
        value = _SEX_TOKENS.get(label)
        if value:
            return value
    return ""


def age_from_labels(labels_raw: tuple[str, ...]) -> tuple[int | None, int | None]:
    for label in labels_raw:
        parsed = parse_age_label(label)
        if parsed is not None:
            return parsed
    return (None, None)


def procedure_from_labels(labels_raw: tuple[str, ...]) -> str:
    for label in labels_raw:
        for token in str(label).replace("-", " ").split():
            if token.isdigit() and 4 <= len(token) <= 9:
                return token
    return ""


def strip_sex_age_tokens(labels_norm: tuple[str, ...]) -> tuple[str, ...]:
    out: list[str] = []
    for label in labels_norm:
        if label in _SEX_LABEL_NORMS:
            continue
        if parse_age_label(label) is not None:
            continue
        out.append(label)
    return tuple(out)


def _scope_from_context(
    row_norm: tuple[str, ...], column_norm: tuple[str, ...], has_formula: bool
) -> str:
    joined = " ".join((*row_norm, *column_norm))
    if "TOTAL" in column_norm or joined.strip().endswith("TOTAL"):
        return _SCOPE_TOTAL
    if has_formula and any(
        r.replace(".", "").replace("-", "").strip().split()[:1] == [seg]
        for r in row_norm for seg in ("TOTAL", "SUBTOTAL")
    ):
        return _SCOPE_SUBTOTAL
    return _SCOPE_DETAIL


def _looks_structural_value(value: object) -> bool:
    """Una celda cuyo valor es un rótulo (código de prestación, texto) no es un
    slot de dato de ingreso."""
    if isinstance(value, str):
        stripped = value.strip()
        if stripped.isdigit() and 4 <= len(stripped) <= 9:
            return True  # celda de código (texto)
        if len(stripped) > 3 and not stripped.replace(".", "").replace(",", "").isdigit():
            return True  # celda de texto / glosa
    if isinstance(value, (int, float)) and not isinstance(value, bool):
        digits = str(int(value)) if float(value).is_integer() else ""
        if 4 <= len(digits) <= 9:
            return True  # código de prestación guardado como número (p.ej. 1103001)
    return False


def _classify_target(*, has_formula: bool, is_blank: bool, has_value: bool,
                     value: object, context_status: str,
                     grid_statuses: tuple[str, ...],
                     sheet_protected: bool, cell_locked: bool | None) -> str:
    """Clasifica una celda de la plantilla.

    `DIRECT_INPUT_TARGET` exige **evidencia estructural de ser una celda de
    ingreso**: en una hoja protegida, que la celda esté explícitamente
    **desbloqueada** (``protection.locked is False``) — el autor de la hoja sólo
    permite escribir ahí. "No tener fórmula" NO basta: rótulos, encabezados y
    códigos tampoco tienen fórmula.
    """
    if has_formula:
        return FORMULA_TARGET
    if has_value and _looks_structural_value(value):
        return STRUCTURAL
    in_grid = context_status in grid_statuses

    if sheet_protected:
        # Semántica de protección de Excel: en una hoja protegida el usuario sólo
        # puede escribir en las celdas DESBLOQUEADAS. Es la señal autoritativa.
        if cell_locked is False:
            return DIRECT_INPUT_TARGET if in_grid else UNKNOWN_TARGET
        # bloqueada -> rótulo / encabezado / contenido estático
        return STRUCTURAL if (in_grid or context_status in ("PARTIAL", "AMBIGUOUS")) else UNKNOWN_TARGET

    # Hoja sin proteger: `protection.locked` no es fiable. Sin evidencia de
    # ingreso -> UNKNOWN es preferible a INPUT_TARGET (revisión de cierre §5).
    if in_grid and is_blank:
        return UNKNOWN_TARGET
    if in_grid or context_status in ("PARTIAL", "AMBIGUOUS"):
        return STRUCTURAL if has_value else UNKNOWN_TARGET
    return UNKNOWN_TARGET


_GRID_STATUSES_DEFAULT = ("COMPLETE",)
_GRID_STATUSES_FLAT = ("COMPLETE", "PARTIAL", "AMBIGUOUS")
# formas sin eje de columna (catálogo plano): se aceptan filas con contexto
# parcial/ambiguo porque su grilla no tiene encabezados de columna limpios.
_FLAT_FORMS = ("B2_ANEXO",)


def build_target_inventory(
    ws_values,
    ws_formulas,
    *,
    policy: AlignmentPolicy,
    regions: list[SourceRegion],
) -> list[TargetMetricContext]:
    """Recorre la grilla de una hoja de la plantilla (sólo lectura).

    ``ws_values`` es la hoja abierta con ``data_only=True`` (las fórmulas dan su
    valor, no el texto ``=...``); ``ws_formulas`` la misma hoja con
    ``data_only=False`` (para saber si la celda tiene fórmula).
    """
    layout = SheetLayout(ws_values)
    form = policy.canonical_form(ws_values.title)
    grid_statuses = _GRID_STATUSES_FLAT if form in _FLAT_FORMS else _GRID_STATUSES_DEFAULT
    keep_statuses = set(grid_statuses)
    sheet_protected = bool(getattr(ws_formulas.protection, "sheet", False))
    out: list[TargetMetricContext] = []
    for r in range(1, layout.max_row + 1):
        for c in range(1, layout.max_col + 1):
            coord = f"{_col_letter(c)}{r}"
            ctx = layout.cell_context(coord)
            if ctx.context_status not in keep_statuses:
                continue
            if not ctx.row_labels_raw:
                continue
            fcell = ws_formulas.cell(row=r, column=c)
            has_formula = fcell.data_type == "f"
            raw_formula = fcell.value if has_formula and isinstance(fcell.value, str) else ""
            cell_locked = getattr(getattr(fcell, "protection", None), "locked", None)
            vcell = ws_values.cell(row=r, column=c)
            has_value = vcell.value not in (None, "") and not has_formula
            is_blank = vcell.value in (None, "") and not has_formula
            kind = _classify_target(
                has_formula=has_formula, is_blank=is_blank, has_value=has_value,
                value=vcell.value, context_status=ctx.context_status,
                grid_statuses=grid_statuses,
                sheet_protected=sheet_protected, cell_locked=cell_locked,
            )
            column_struct = strip_sex_age_tokens(ctx.column_labels_norm)
            expectation, evidence = region_expectation(
                regions, ws_values.title, r, ctx.section_labels_norm
            )
            out.append(
                TargetMetricContext(
                    target_sheet=ws_values.title,
                    target_cell=coord,
                    form=form,
                    row=r,
                    column=c,
                    section_path=ctx.section_labels_raw,
                    row_path=ctx.row_labels_raw,
                    column_path=ctx.column_labels_raw,
                    section_norm=ctx.section_labels_norm,
                    row_norm=ctx.row_labels_norm,
                    column_norm=ctx.column_labels_norm,
                    column_structure_norm=column_struct,
                    target_kind=kind,
                    target_alignment_role=alignment_role(kind),
                    has_formula=has_formula,
                    formula=raw_formula,
                    is_blank=is_blank,
                    sex_value=sex_from_labels(ctx.column_labels_norm),
                    age_min=age_from_labels(ctx.column_labels_raw)[0],
                    age_max=age_from_labels(ctx.column_labels_raw)[1],
                    aggregation_scope=_scope_from_context(
                        ctx.row_labels_norm, ctx.column_labels_norm, has_formula
                    ),
                    procedure_code=procedure_from_labels(ctx.row_labels_raw),
                    semantic_signature=ctx.semantic_signature(),
                    target_source_expectation=expectation,
                    target_source_evidence=evidence,
                )
            )
    return out


# ---------------------------------------------------------------------------
# Scoring por dimensión
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class DimensionMatch:
    dimension: str
    source_value: str
    target_value: str
    evidence_status: str


def _norm_tuple(values) -> tuple[str, ...]:
    return tuple(normalize_semantic_label(v) for v in values if str(v).strip())


def _seq_evidence(source: tuple[str, ...], target: tuple[str, ...]) -> str:
    if not source and not target:
        return EV_NOT_APPLICABLE
    if not source or not target:
        return EV_MISSING
    if source == target:
        return EV_MATCH
    same_leaf = source[-1] == target[-1]
    nested = set(source) <= set(target) or set(target) <= set(source)
    if same_leaf and nested:
        # misma hoja identifica la fila igual; la otra sólo muestra más (o menos)
        # niveles de ancestro -> diferencia de profundidad, no de identidad.
        return EV_MATCH
    if same_leaf or nested or (set(source) & set(target)):
        return EV_PARTIAL
    return EV_MISMATCH


def _sex_evidence(source_value: str, source_status: str, target_value: str,
                  applicable: bool) -> str:
    if not applicable:
        return EV_NOT_APPLICABLE
    if source_status not in ("CONFIRMED", "LABEL_ONLY", "FORMULA_ONLY") or not source_value:
        return EV_MISSING
    if not target_value:
        return EV_MISSING
    if source_value == target_value:
        return EV_MATCH
    if "BOTH" in (source_value, target_value):
        return EV_PARTIAL
    return EV_MISMATCH


def _age_evidence(s_min, s_max, s_status, t_min, t_max, applicable: bool) -> str:
    if not applicable:
        return EV_NOT_APPLICABLE
    has_source = s_status in ("CONFIRMED", "LABEL_ONLY", "FORMULA_ONLY") and (
        s_min is not None or s_max is not None
    )
    has_target = t_min is not None or t_max is not None
    if not has_source or not has_target:
        return EV_MISSING
    if (s_min, s_max) == (t_min, t_max):
        return EV_MATCH
    lo = max(s_min or 0, t_min or 0)
    hi = min(s_max if s_max is not None else 200, t_max if t_max is not None else 200)
    return EV_PARTIAL if lo <= hi else EV_MISMATCH


def _procedure_evidence(source_code: str, target_code: str) -> str:
    if not source_code and not target_code:
        return EV_NOT_APPLICABLE
    if not source_code or not target_code:
        return EV_MISSING
    return EV_MATCH if source_code == target_code else EV_MISMATCH


@dataclass(frozen=True)
class ScoredPair:
    score: float
    dimensions: tuple[DimensionMatch, ...]
    has_conflict: bool

    def evidence(self, dimension: str) -> str:
        for dim in self.dimensions:
            if dim.dimension == dimension:
                return dim.evidence_status
        return EV_MISSING


_EV_VALUE = {EV_MATCH: 1.0, EV_PARTIAL: 0.5, EV_MISMATCH: 0.0}


def score_pair(source, target: TargetMetricContext, policy: AlignmentPolicy) -> ScoredPair:
    """Score determinista y explicable de emparejar ``source`` con ``target``."""
    form_applicable = target.form in ("B2_ANEXO",)
    sex_age_applies = target.form not in ("B2_ANEXO",)

    dims: list[DimensionMatch] = []

    def add(name: str, sval: str, tval: str, ev: str) -> None:
        dims.append(DimensionMatch(name, sval, tval, ev))

    src_form = normalize_semantic_label(source.form)
    add("FORM", src_form, target.form, EV_MATCH if src_form == target.form else EV_MISMATCH)

    s_section = _norm_tuple(source.section_path_raw)
    s_row = _norm_tuple(source.row_path_raw)
    s_col_struct = strip_sex_age_tokens(_norm_tuple(source.column_path_raw))
    flat_form = target.form in _FLAT_FORMS
    add("SECTION", " :: ".join(s_section), " :: ".join(target.section_norm),
        _seq_evidence(s_section, target.section_norm))
    row_ev = _seq_evidence(s_row, target.row_norm)
    # Si ambos lados identifican la fila por el MISMO código de procedimiento, la
    # fila es la misma aunque el rótulo esté formateado distinto (código + glosa
    # vs. sección + código).
    src_code = source.procedure_code_raw
    if (
        src_code and target.procedure_code == src_code
        and src_code in s_row and src_code in target.row_norm
        and row_ev in (EV_MISMATCH, EV_PARTIAL)
    ):
        row_ev = EV_MATCH
    add("ROW_PATH", " :: ".join(s_row), " :: ".join(target.row_norm), row_ev)
    col_ev = EV_NOT_APPLICABLE if flat_form else _seq_evidence(
        s_col_struct, target.column_structure_norm
    )
    add("COLUMN_PATH", " :: ".join(s_col_struct), " :: ".join(target.column_structure_norm), col_ev)
    add("SEX", source.sex_value, target.sex_value,
        _sex_evidence(source.sex_value, source.sex_status, target.sex_value, sex_age_applies))
    add("AGE", _fmt_age(source.age_min_years, source.age_max_years),
        _fmt_age(target.age_min, target.age_max),
        _age_evidence(source.age_min_years, source.age_max_years, source.age_status,
                      target.age_min, target.age_max, sex_age_applies))
    add("AGGREGATION_SCOPE", source.aggregation_scope, target.aggregation_scope,
        EV_MATCH if source.aggregation_scope == target.aggregation_scope else EV_MISMATCH)
    add("PROCEDURE_CODE", source.procedure_code_raw, target.procedure_code,
        _procedure_evidence(source.procedure_code_raw, target.procedure_code))

    # FORM es una compuerta: si no coincide, no hay match posible.
    form_ev = next(d.evidence_status for d in dims if d.dimension == "FORM")
    if form_ev != EV_MATCH and not form_applicable:
        return ScoredPair(0.0, tuple(dims), has_conflict=False)

    numer = denom = 0.0
    for dim in dims:
        if dim.dimension == "FORM":
            continue
        if dim.evidence_status in (EV_NOT_APPLICABLE, EV_MISSING):
            continue
        w = policy.weight(dim.dimension)
        denom += w
        numer += w * _EV_VALUE.get(dim.evidence_status, 0.0)
    score = round(numer / denom, 4) if denom else 0.0

    has_conflict = any(
        d.dimension in _CONFLICTING_DIMENSIONS and d.evidence_status == EV_MISMATCH
        for d in dims
    )
    return ScoredPair(score, tuple(dims), has_conflict)


def _fmt_age(lo, hi) -> str:
    if lo is None and hi is None:
        return ""
    return f"{lo if lo is not None else ''}..{hi if hi is not None else ''}"


# ---------------------------------------------------------------------------
# Alineación
# ---------------------------------------------------------------------------

# --- motivos de source sin match --------------------------------
REASON_NO_TARGET = "NO_TARGET"
REASON_TARGET_NOT_MEDINET = "TARGET_SOURCE_NOT_MEDINET"
REASON_CONTEXT_INSUFFICIENT = "CONTEXT_INSUFFICIENT"
REASON_CONFLICT = "CONFLICT"

# --- motivos de source excluido (no elegible) ------------------
EXCL_BLOCKED_CONFLICT = "BLOCKED_CONFLICT"
EXCL_REVIEW_REQUIRED = "REVIEW_REQUIRED"
EXCL_NOT_APPLICABLE = "ROLLUP_NOT_INPUT"  # readiness NOT_APPLICABLE = TOTAL/SUBTOTAL roll-up
EXCL_NOT_MEDINET = "SOURCE_NOT_MEDINET"

# --- motivos de target sin match (taxonomía de rol) -----------
TGT_NO_MEDINET_SOURCE_FOUND = "NO_MEDINET_SOURCE_FOUND"        # genuino: faltó el source
TGT_NON_MEDINET_REGION = "NON_MEDINET_REGION"                  # nunca era MEDINET
TGT_DERIVED_NOT_INPUT = "DERIVED_TARGET_NOT_INPUT"             # fórmula, no punto de ingreso
TGT_STRUCTURAL_NOT_ALIGNMENT = "STRUCTURAL_NOT_ALIGNMENT_TARGET"
TGT_VALIDATION_NOT_ALIGNMENT = "VALIDATION_NOT_ALIGNMENT_TARGET"
TGT_UNKNOWN_SOURCE = "UNKNOWN_SOURCE"                          # input sin fuente demostrable


@dataclass(frozen=True)
class AlignmentRow:
    source_metric_id: str
    source_form: str
    source_kind: str
    source_readiness: str
    target_sheet: str
    target_cell: str
    target_kind: str
    match_status: str
    match_score: float
    source_semantic_signature: str
    target_semantic_signature: str
    sex_match: str
    age_match: str
    procedure_match: str
    row_match: str
    column_match: str
    section_match: str
    target_alignment_role: str
    target_source_expectation: str
    needs_human_review: bool


@dataclass(frozen=True)
class ExcludedSource:
    metric_id: str
    form: str
    readiness_status: str
    reason: str
    detail: str


@dataclass(frozen=True)
class AmbiguousMatch:
    source_metric_id: str
    candidate_count: int
    candidate_cells: tuple[str, ...]
    reason: str


@dataclass(frozen=True)
class UnmatchedSource:
    metric_id: str
    form: str
    row_path: str
    column_path: str
    reason: str


@dataclass(frozen=True)
class UnmatchedTarget:
    target_sheet: str
    target_cell: str
    target_kind: str
    target_alignment_role: str
    target_context: str
    expected_source: str
    reason: str
    is_genuine_unmatched_medinet_input: bool


@dataclass
class AlignmentReport:
    policy_version: str
    rows: list[AlignmentRow] = field(default_factory=list)
    evidence: list[tuple] = field(default_factory=list)  # (mid, sheet, cell, dim, sval, tval, ev)
    excluded_sources: list[ExcludedSource] = field(default_factory=list)
    ambiguous: list[AmbiguousMatch] = field(default_factory=list)
    unmatched_sources: list[UnmatchedSource] = field(default_factory=list)
    unmatched_targets: list[UnmatchedTarget] = field(default_factory=list)
    target_inventory: list[TargetMetricContext] = field(default_factory=list)
    eligible_source_ids: tuple[str, ...] = ()


def _readiness_eligibility(readiness) -> tuple[bool, str, str]:
    """(eligible, exclusion_reason, detail). Sólo AUTO_READY es elegible."""
    rs = readiness.readiness_status
    if rs == "AUTO_READY":
        return True, "", ""
    if rs == "BLOCKED_CONFLICT":
        return False, EXCL_BLOCKED_CONFLICT, "conflicto rótulo↔fórmula (Sprint 3.2/3.3)"
    if rs == "REVIEW_REQUIRED":
        return False, EXCL_REVIEW_REQUIRED, "|".join(readiness.review_reasons)
    if rs == "NOT_APPLICABLE":
        # Sprint 3.3: readiness NOT_APPLICABLE == TOTAL/SUBTOTAL roll-up (verificado:
        # los 315 son DOWNSTREAM_TOTAL con aggregation_scope TOTAL/SUBTOTAL, 0 excepciones).
        return False, EXCL_NOT_APPLICABLE, "roll-up TOTAL/SUBTOTAL; no es punto de ingreso"
    return False, "UNKNOWN_READINESS", rs


def _index_targets(
    targets_by_form: dict[str, list[TargetMetricContext]],
) -> dict[str, dict[str, dict]]:
    """Índice por forma: {'leaf': {row_leaf -> [t]}, 'section': {...}, 'proc': {...}}."""
    index: dict[str, dict[str, dict]] = {}
    for form, targets in targets_by_form.items():
        leaf: dict[str, list] = {}
        section: dict[str, list] = {}
        proc: dict[str, list] = {}
        for t in targets:
            if t.row_norm:
                leaf.setdefault(t.row_norm[-1], []).append(t)
            for s in t.section_norm:
                section.setdefault(s, []).append(t)
            if t.procedure_code:
                proc.setdefault(t.procedure_code, []).append(t)
        index[form] = {"leaf": leaf, "section": section, "proc": proc}
    return index


def _candidate_targets(metric, index: dict, targets_by_form: dict, form_key: str) -> list:
    forma = index.get(form_key)
    if forma is None:
        return targets_by_form.get(form_key, [])
    seen: dict[tuple[str, str], TargetMetricContext] = {}
    s_row = _norm_tuple(metric.row_path_raw)
    if s_row:
        for t in forma["leaf"].get(s_row[-1], []):
            seen[t.key] = t
    for s in _norm_tuple(metric.section_path_raw):
        for t in forma["section"].get(s, []):
            seen[t.key] = t
    if metric.procedure_code_raw:
        for t in forma["proc"].get(metric.procedure_code_raw, []):
            seen[t.key] = t
    return list(seen.values())


def align(
    source_metrics: list,
    readiness_by_id: dict,
    targets_by_form: dict[str, list[TargetMetricContext]],
    policy: AlignmentPolicy,
    regions: list[SourceRegion],
) -> AlignmentReport:
    report = AlignmentReport(policy_version=policy.version, target_inventory=[
        t for ts in targets_by_form.values() for t in ts
    ])
    index = _index_targets(targets_by_form)

    matched_target_keys: set[tuple[str, str]] = set()
    eligible_ids: list[str] = []

    for metric in source_metrics:
        readiness = readiness_by_id.get(metric.metric_id)
        if readiness is None:
            report.excluded_sources.append(ExcludedSource(
                metric.metric_id, metric.form, "UNKNOWN", "NO_READINESS", ""))
            continue
        if metric.source != "MEDINET":
            report.excluded_sources.append(ExcludedSource(
                metric.metric_id, metric.form, readiness.readiness_status,
                EXCL_NOT_MEDINET, metric.source))
            continue
        eligible, reason, detail = _readiness_eligibility(readiness)
        if not eligible:
            report.excluded_sources.append(ExcludedSource(
                metric.metric_id, metric.form, readiness.readiness_status, reason, detail))
            continue

        eligible_ids.append(metric.metric_id)
        form_key = normalize_semantic_label(metric.form).replace(" ", "_")
        if form_key not in targets_by_form and metric.form in targets_by_form:
            form_key = metric.form
        candidates = _candidate_targets(metric, index, targets_by_form, form_key)

        scored: list[tuple[float, TargetMetricContext, ScoredPair]] = []
        for target in candidates:
            if target.target_kind not in (DIRECT_INPUT_TARGET, FORMULA_TARGET):
                continue
            sp = score_pair(metric, target, policy)
            if sp.dimensions and sp.evidence("FORM") != EV_MATCH:
                continue
            scored.append((sp.score, target, sp))

        real = [x for x in scored if not x[2].has_conflict]
        confl = [x for x in scored if x[2].has_conflict]

        if not real and confl:
            best = max(confl, key=lambda x: x[0])
            report.rows.append(_row(metric, readiness, best[1], best[2], CONFLICT, policy))
            _record_evidence(report, metric.metric_id, best[1], best[2])
            report.unmatched_sources.append(UnmatchedSource(
                metric.metric_id, metric.form, " :: ".join(metric.row_path_raw),
                " :: ".join(metric.column_path_raw), REASON_CONFLICT))
            continue

        if not real:
            report.unmatched_sources.append(UnmatchedSource(
                metric.metric_id, metric.form, " :: ".join(metric.row_path_raw),
                " :: ".join(metric.column_path_raw), REASON_NO_TARGET))
            continue

        real.sort(key=lambda x: (-x[0], x[1].target_cell))
        best_score, best_target, best_sp = real[0]

        if best_score < policy.strong_threshold:
            report.unmatched_sources.append(UnmatchedSource(
                metric.metric_id, metric.form, " :: ".join(metric.row_path_raw),
                " :: ".join(metric.column_path_raw), REASON_CONTEXT_INSUFFICIENT))
            continue

        near = [x for x in real if best_score - x[0] <= policy.tie_margin]
        # target en región explícitamente NO-MEDINET (EGRESOS / recursos / tabla
        # quirúrgica): se excluye para no mezclar fuentes. Una expectativa UNKNOWN
        # NO bloquea: el propio match fuerte es evidencia de que es MEDINET.
        if best_target.target_source_expectation in _NON_MEDINET_EXPECTATIONS:
            report.unmatched_sources.append(UnmatchedSource(
                metric.metric_id, metric.form, " :: ".join(metric.row_path_raw),
                " :: ".join(metric.column_path_raw), REASON_TARGET_NOT_MEDINET))
            continue

        if len(near) > 1:
            report.ambiguous.append(AmbiguousMatch(
                metric.metric_id, len(near),
                tuple(f"{t.target_sheet}!{t.target_cell}" for _s, t, _p in near[:8]),
                f"{len(near)} targets dentro de {policy.tie_margin} del mejor score "
                f"({best_score})"))
            report.rows.append(_row(metric, readiness, best_target, best_sp, AMBIGUOUS, policy))
            _record_evidence(report, metric.metric_id, best_target, best_sp)
            continue

        exact = (
            best_score >= policy.exact_threshold
            and best_sp.evidence("ROW_PATH") == EV_MATCH
            and best_sp.evidence("COLUMN_PATH") in (EV_MATCH, EV_NOT_APPLICABLE)
            and best_sp.evidence("SEX") in (EV_MATCH, EV_NOT_APPLICABLE)
            and best_sp.evidence("AGE") in (EV_MATCH, EV_NOT_APPLICABLE)
        )
        status = EXACT_SEMANTIC_MATCH if exact else STRONG_MATCH
        # Evidencia: un match EXACT/STRONG de un source MEDINET *es* evidencia de
        # que el target pertenece al subgrafo MEDINET. Se sube la expectativa
        # ``UNKNOWN`` -> ``MEDINET`` (nunca se pisa una región NO-MEDINET: ésas ya
        # se excluyeron arriba).
        if best_target.target_source_expectation == EXP_UNKNOWN:
            best_target.target_source_expectation = EXP_MEDINET
            best_target.target_source_evidence = (
                f"resuelto por alineación: {status} con {metric.metric_id}"
            )
        report.rows.append(_row(metric, readiness, best_target, best_sp, status, policy))
        _record_evidence(report, metric.metric_id, best_target, best_sp)
        matched_target_keys.add(best_target.key)

    report.eligible_source_ids = tuple(eligible_ids)

    for target in report.target_inventory:
        if target.key in matched_target_keys:
            continue
        reason, genuine = _unmatched_target_reason(target)
        report.unmatched_targets.append(UnmatchedTarget(
            target_sheet=target.target_sheet,
            target_cell=target.target_cell,
            target_kind=target.target_kind,
            target_alignment_role=target.target_alignment_role,
            target_context=" :: ".join(
                (*target.section_path, *target.row_path, *target.column_path)
            ),
            expected_source=target.target_source_expectation,
            reason=reason,
            is_genuine_unmatched_medinet_input=genuine,
        ))

    return report


def _unmatched_target_reason(target: TargetMetricContext) -> tuple[str, bool]:
    """(reason, is_genuine_unmatched_medinet_input).

    Distingue "faltó encontrar un source" de "esta celda nunca debía tener uno".
    """
    if target.target_kind == STRUCTURAL:
        return TGT_STRUCTURAL_NOT_ALIGNMENT, False
    if target.target_kind == VALIDATION:
        return TGT_VALIDATION_NOT_ALIGNMENT, False
    if target.target_kind == FORMULA_TARGET:
        return TGT_DERIVED_NOT_INPUT, False
    if target.target_kind != DIRECT_INPUT_TARGET:  # UNKNOWN físico
        return TGT_UNKNOWN_SOURCE, False
    exp = target.target_source_expectation
    if exp in _NON_MEDINET_EXPECTATIONS:
        return TGT_NON_MEDINET_REGION, False
    if exp == EXP_MEDINET:
        return TGT_NO_MEDINET_SOURCE_FOUND, True  # genuino input Medinet sin candidato
    return TGT_UNKNOWN_SOURCE, False


def _row(metric, readiness, target: TargetMetricContext, sp: ScoredPair,
         status: str, policy: AlignmentPolicy) -> AlignmentRow:
    return AlignmentRow(
        source_metric_id=metric.metric_id,
        source_form=metric.form,
        source_kind=metric.kind,
        source_readiness=readiness.readiness_status,
        target_sheet=target.target_sheet,
        target_cell=target.target_cell,
        target_kind=target.target_kind,
        match_status=status,
        match_score=sp.score,
        source_semantic_signature=metric.semantic_signature,
        target_semantic_signature=target.semantic_signature,
        sex_match=sp.evidence("SEX"),
        age_match=sp.evidence("AGE"),
        procedure_match=sp.evidence("PROCEDURE_CODE"),
        row_match=sp.evidence("ROW_PATH"),
        column_match=sp.evidence("COLUMN_PATH"),
        section_match=sp.evidence("SECTION"),
        target_alignment_role=target.target_alignment_role,
        target_source_expectation=target.target_source_expectation,
        needs_human_review=status in (STRONG_MATCH, AMBIGUOUS, CONFLICT)
        or target.target_source_expectation in _NON_MEDINET_EXPECTATIONS,
    )


def _record_evidence(report: AlignmentReport, mid: str, target: TargetMetricContext,
                     sp: ScoredPair) -> None:
    for dim in sp.dimensions:
        report.evidence.append((
            mid, target.target_sheet, target.target_cell, dim.dimension,
            dim.source_value, dim.target_value, dim.evidence_status,
        ))
