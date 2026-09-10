"""Write readiness del mapeo semántico -> plantilla oficial (Sprint 3.6).

Capa por ENCIMA del semantic alignment (Sprint 3.5). Convierte los alignments
*técnicamente seguros* en un conjunto explícito, auditable y versionable de
instrucciones de escritura — **sin escribir Excel**.

`write_status` ∈ ``WRITE_READY · WRITE_REVIEW_REQUIRED · WRITE_BLOCKED ·
NOT_WRITABLE`` — **distinto** de `readiness` (Sprint 3.3) y de `match_status`
(Sprint 3.5). Las capas se mantienen separadas.

3.6 decide **DÓNDE** escribir, no **QUÉ** población clínica entra (el filtro por
ESTADO sigue `PENDING_FUNCTIONAL_CONFIRMATION` y no se aplica) ni **QUÉ** valor se
escribe (eso es el *value producer* del Sprint 3.7).
"""

from __future__ import annotations

import hashlib
from collections import Counter, defaultdict
from collections.abc import Iterable, Sequence
from dataclasses import dataclass, field, replace
from pathlib import Path
from typing import Protocol

import yaml

from remasep.core.errors import RemasepError

# --- write status --------------------------------------------------

WRITE_READY = "WRITE_READY"
WRITE_REVIEW_REQUIRED = "WRITE_REVIEW_REQUIRED"
WRITE_BLOCKED = "WRITE_BLOCKED"
NOT_WRITABLE = "NOT_WRITABLE"

# --- expected value type ----------------------------------------

EVT_INTEGER_COUNT = "INTEGER_COUNT"
EVT_NUMERIC = "NUMERIC"
EVT_TEXT = "TEXT"
EVT_UNKNOWN = "UNKNOWN"

# --- motivos ------------------------------------------------------
R_EXACT_ALL_MATCH = "EXACT_ALL_DIMENSIONS_MATCH"
R_STRONG_REQUIRED_MATCH = "STRONG_REQUIRED_DIMENSIONS_MATCH"
R_STRONG_MISSING_REQUIRED = "STRONG_MISSING_REQUIRED_DIMENSION"
R_STRONG_DIMENSION_MISMATCH = "STRONG_DIMENSION_MISMATCH"
R_TARGET_SOURCE_UNKNOWN = "TARGET_SOURCE_EXPECTATION_UNKNOWN"
R_EXPECTED_VALUE_TYPE_INCOMPATIBLE = "EXPECTED_VALUE_TYPE_INCOMPATIBLE"
R_MATCH_AMBIGUOUS = "MATCH_AMBIGUOUS"
R_DIMENSION_CONFLICT = "DIMENSION_CONFLICT"
R_SOURCE_BLOCKED_CONFLICT = "SOURCE_BLOCKED_CONFLICT"
R_TARGET_LOCKED = "TARGET_LOCKED"
R_TARGET_SOURCE_NOT_MEDINET = "TARGET_SOURCE_NOT_MEDINET"
R_WRITE_COLLISION = "WRITE_COLLISION"
R_SOURCE_REVIEW_REQUIRED = "SOURCE_REVIEW_REQUIRED"
R_ROLLUP_COMPUTED_BY_TEMPLATE = "ROLLUP_COMPUTED_BY_TEMPLATE"
R_SOURCE_NOT_MEDINET = "SOURCE_NOT_MEDINET"
R_NO_ALIGNMENT_FOUND = "NO_ALIGNMENT_FOUND"
R_TARGET_IS_FORMULA = "TARGET_IS_FORMULA"
R_TARGET_NOT_INPUT_CELL = "TARGET_NOT_INPUT_CELL"
R_TARGET_ROLE_NOT_INPUT = "TARGET_ROLE_NOT_INPUT"

# --- colisiones -----------------------------------------------
COLL_MULTI_SOURCE_SAME_TARGET = "MULTIPLE_SOURCES_SAME_TARGET"
COLL_SAME_SOURCE_MULTI_TARGET = "SAME_SOURCE_MULTIPLE_TARGETS"
COLL_DUPLICATE_INSTRUCTION_ID = "DUPLICATE_INSTRUCTION_IDENTITY"
COLL_INCOMPATIBLE_TARGET_REUSE = "INCOMPATIBLE_TARGET_REUSE"

# --- constantes de las capas inferiores (para no acoplar imports) ---
_MATCH_EXACT = "EXACT_SEMANTIC_MATCH"
_MATCH_STRONG = "STRONG_MATCH"
_MATCH_AMBIGUOUS = "AMBIGUOUS"
_MATCH_CONFLICT = "CONFLICT"
_KIND_DIRECT_INPUT = "DIRECT_INPUT_TARGET"
_KIND_FORMULA = "FORMULA_TARGET"
_ROLE_INPUT = "INPUT_TARGET"
_EXP_MEDINET = "MEDINET"
_NON_MEDINET_EXPECTATIONS = frozenset({
    "EGRESOS", "RESOURCE_CALCULATION", "SURGICAL_TABLE", "CONTROL_METADATA",
})
_EV_MATCH = "MATCH"
_EV_NOT_APPLICABLE = "NOT_APPLICABLE"
_EV_MISMATCH = "MISMATCH"
# motivos de source excluido (Sprint 3.5)
_EXCL_BLOCKED_CONFLICT = "BLOCKED_CONFLICT"
_EXCL_REVIEW_REQUIRED = "REVIEW_REQUIRED"
_EXCL_ROLLUP = "ROLLUP_NOT_INPUT"
_EXCL_NOT_MEDINET = "SOURCE_NOT_MEDINET"

_DEFAULT_CONFIG_DIR = (
    Path(__file__).resolve().parents[3] / "config" / "writable_target_mapping_2026"
)


class WritableTargetMappingError(RemasepError):
    """Configuración de write mapping inválida."""


# ---------------------------------------------------------------------------
# Value producer contract (Sprint 3.7) — sólo interfaz, sin implementación
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class MetricValue:
    """Valor agregado de una métrica MEDINET para un período. **Sin PII.**

    Lo produce el *value producer* del Sprint 3.7 a partir de
    ``processing_scope_records`` (registros válidos ∩ mes/año) — **no** de todos
    los registros válidos, y **sin** aplicar el filtro por ESTADO pendiente de
    confirmación.
    """

    source_metric_id: str
    value: object
    period: str
    producer_version: str


class ValueProducer(Protocol):
    """Interfaz conceptual del Sprint 3.7. 3.6 no la implementa."""

    def produce(self, source_metric_id: str) -> MetricValue: ...


# ---------------------------------------------------------------------------
# Política versionada
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class WriteReadinessPolicy:
    version: str
    allowed_readiness: frozenset[str]
    allowed_target_kind: frozenset[str]
    allowed_target_role: frozenset[str]
    allowed_target_locked: frozenset[bool]
    allowed_expectation: frozenset[str]
    exact_write_ready: bool
    strong_auto_ready: bool
    strong_no_mismatch: bool
    strong_required_default: tuple[str, ...]
    strong_required_by_form: dict[str, tuple[str, ...]]
    strong_acceptable_evidence: frozenset[str]
    evt_by_kind: dict[str, str]
    evt_default: str
    evt_on_incompatible: str
    zero_write_policy: str
    collision_status: dict[str, str]
    template_compat: dict

    def strong_required(self, form: str) -> tuple[str, ...]:
        return self.strong_required_by_form.get(form, self.strong_required_default)

    def expected_value_type(self, source_kind: str) -> str:
        return self.evt_by_kind.get(source_kind, self.evt_default)


def _load_yaml(path: Path) -> dict:
    if not path.is_file():
        raise WritableTargetMappingError(f"config de write mapping no encontrada: {path}")
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def load_write_policy(config_dir: str | Path | None = None) -> WriteReadinessPolicy:
    base = Path(config_dir) if config_dir is not None else _DEFAULT_CONFIG_DIR
    raw = _load_yaml(base / "policy.yaml")
    sources = raw.get("allowed_sources") or {}
    targets = raw.get("allowed_targets") or {}
    strong = raw.get("strong_match") or {}
    req = strong.get("required_dimensions") or {}
    evt = raw.get("expected_value_type") or {}
    return WriteReadinessPolicy(
        version=str(raw.get("version", "writable_target_mapping_2026")),
        allowed_readiness=frozenset(sources.get("readiness") or ["AUTO_READY"]),
        allowed_target_kind=frozenset(targets.get("target_kind") or [_KIND_DIRECT_INPUT]),
        allowed_target_role=frozenset(targets.get("target_alignment_role") or [_ROLE_INPUT]),
        allowed_target_locked=frozenset(bool(v) for v in (targets.get("target_locked") or [False])),
        allowed_expectation=frozenset(targets.get("target_source_expectation") or [_EXP_MEDINET]),
        exact_write_ready=bool((raw.get("exact_match") or {}).get("write_ready", True)),
        strong_auto_ready=bool(strong.get("auto_write_ready", True)),
        strong_no_mismatch=bool(strong.get("no_dimension_mismatch", True)),
        strong_required_default=tuple(req.get("default") or []),
        strong_required_by_form={
            str(f): tuple(dims) for f, dims in (req.get("by_form") or {}).items()
        },
        strong_acceptable_evidence=frozenset(
            strong.get("acceptable_evidence") or [_EV_MATCH, _EV_NOT_APPLICABLE]
        ),
        evt_by_kind={str(k): str(v) for k, v in (evt.get("by_source_kind") or {}).items()},
        evt_default=str(evt.get("default", EVT_INTEGER_COUNT)),
        evt_on_incompatible=str(evt.get("on_incompatible_target_value", WRITE_REVIEW_REQUIRED)),
        zero_write_policy=str(raw.get("zero_write_policy", "UNRESOLVED")),
        collision_status={
            str(k): str(v) for k, v in (raw.get("collision_policy") or {}).items()
        },
        template_compat=raw.get("template_compatibility") or {},
    )


# ---------------------------------------------------------------------------
# Fingerprint estructural de la plantilla
# ---------------------------------------------------------------------------


def _sha16(text: str) -> str:
    return hashlib.sha256(text.encode("utf-8")).hexdigest()[:16]


def structural_template_fingerprint(wb_formula, *, file_sha256: str = "") -> dict:
    """Fingerprint REPRODUCIBLE e independiente del SHA256 del archivo.

    Un re-guardado que altere VBA/hash pero preserve la estructura semántica
    (hojas, orden, layout de fórmulas, celdas desbloqueadas, merges relevantes)
    produce el MISMO ``structural_sha256``. Un cambio estructural, no.
    """
    per_sheet: dict[str, dict] = {}
    for ws in wb_formula.worksheets:
        formula_coords: list[str] = []
        unlocked_coords: list[str] = []
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    formula_coords.append(cell.coordinate)
                prot = getattr(cell, "protection", None)
                if getattr(prot, "locked", True) is False:
                    unlocked_coords.append(cell.coordinate)
        merged = sorted(str(mr) for mr in ws.merged_cells.ranges)
        per_sheet[ws.title] = {
            "formula_count": len(formula_coords),
            "formula_coords_sha16": _sha16(",".join(sorted(formula_coords))),
            "protected": bool(getattr(ws.protection, "sheet", False)),
            "unlocked_count": len(unlocked_coords),
            "unlocked_coords_sha16": _sha16(",".join(sorted(unlocked_coords))),
            "merged_ranges_count": len(merged),
            "merged_ranges_sha16": _sha16(",".join(merged)),
        }
    canonical = repr(
        [list(wb_formula.sheetnames), sorted(per_sheet.items())]
    )
    structural_sha = _sha16(canonical)
    return {
        "structural_sha256": structural_sha,
        "structural_template_fingerprint_id": f"stf:{structural_sha}",
        "file_sha256": file_sha256,
        "note": (
            "structural_template_fingerprint_id ata el manifiesto a la ESTRUCTURA "
            "semántica; file_sha256 es sólo informativo (un re-guardado que cambie "
            "VBA no invalida el manifiesto si la estructura permanece)."
        ),
        "sheet_names_ordered": list(wb_formula.sheetnames),
        "per_sheet": per_sheet,
    }


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class WriteInstruction:
    instruction_id: str
    source_metric_id: str
    source_semantic_signature: str
    target_sheet: str
    target_cell: str
    target_semantic_signature: str
    template_fingerprint_id: str
    source_form: str
    source_kind: str
    match_status: str
    write_status: str
    expected_value_type: str
    evidence: tuple[str, ...]
    policy_version: str
    needs_human_review: bool


@dataclass(frozen=True)
class WriteReadiness:
    source_metric_id: str
    source_semantic_signature: str
    source_readiness: str
    source_form: str
    source_kind: str
    match_status: str
    target_sheet: str
    target_cell: str
    target_semantic_signature: str
    target_kind: str
    target_alignment_role: str
    target_source_expectation: str
    target_locked: bool | None
    write_status: str
    write_reason: str
    expected_value_type: str
    missing_evidence: tuple[str, ...]
    evidence_summary: str
    needs_human_review: bool
    instruction: WriteInstruction | None = None


@dataclass(frozen=True)
class Collision:
    collision_type: str
    key: str
    involved: tuple[str, ...]
    detail: str


@dataclass
class WriteMappingReport:
    policy_version: str
    template_fingerprint: dict
    readiness: list[WriteReadiness] = field(default_factory=list)
    instructions: list[WriteInstruction] = field(default_factory=list)
    collisions: list[Collision] = field(default_factory=list)
    zero_write_policy: str = "UNRESOLVED"
    estado_filter_status: str = "PENDING_FUNCTIONAL_CONFIRMATION"


# ---------------------------------------------------------------------------
# Identidad
# ---------------------------------------------------------------------------

_UNIT = "\x1f"
_UNSET = object()


def instruction_id(
    source_signature: str, target_signature: str, template_fingerprint_id: str
) -> str:
    """ID estable derivado de la IDENTIDAD SEMÁNTICA, no de la coordenada."""
    raw = _UNIT.join((source_signature, target_signature, template_fingerprint_id))
    return "wi:" + hashlib.sha256(raw.encode("utf-8")).hexdigest()[:20]


# ---------------------------------------------------------------------------
# Evaluación por item
# ---------------------------------------------------------------------------

# dimensiones observables en un AlignmentRow -> nombre del atributo de evidencia
_DIM_ATTR = {
    "SECTION": "section_match",
    "ROW_PATH": "row_match",
    "COLUMN_PATH": "column_match",
    "SEX": "sex_match",
    "AGE": "age_match",
    "PROCEDURE_CODE": "procedure_match",
}


def _row_evidence(row, dim: str) -> str:
    return getattr(row, _DIM_ATTR[dim], "MISSING")


def _all_dimension_evidences(row) -> dict[str, str]:
    return {dim: getattr(row, attr) for dim, attr in _DIM_ATTR.items()}


def expected_value_type_for(source_kind: str, target_value: object,
                            policy: WriteReadinessPolicy) -> tuple[str, bool]:
    """(expected_value_type, target_value_is_compatible)."""
    evt = policy.expected_value_type(source_kind)
    if target_value in (None, ""):
        return evt, True  # objetivo vacío: compatible con cualquier tipo numérico
    if isinstance(target_value, bool):
        return evt, evt not in (EVT_INTEGER_COUNT, EVT_NUMERIC)
    if isinstance(target_value, (int, float)):
        return evt, evt in (EVT_INTEGER_COUNT, EVT_NUMERIC)
    # texto en un target donde esperamos un conteo -> incompatible
    return evt, evt == EVT_TEXT


def _strong_assessment(row, policy: WriteReadinessPolicy) -> tuple[str, str, tuple[str, ...]]:
    """(write_status, reason, missing_evidence) para un STRONG_MATCH."""
    evidences = _all_dimension_evidences(row)
    if policy.strong_no_mismatch and any(v == _EV_MISMATCH for v in evidences.values()):
        bad = tuple(d for d, v in evidences.items() if v == _EV_MISMATCH)
        return WRITE_REVIEW_REQUIRED, R_STRONG_DIMENSION_MISMATCH, bad
    required = policy.strong_required(row.source_form)
    missing = tuple(
        d for d in required if _row_evidence(row, d) not in (_EV_MATCH, _EV_NOT_APPLICABLE)
    )
    if missing:
        return WRITE_REVIEW_REQUIRED, R_STRONG_MISSING_REQUIRED, missing
    if not policy.strong_auto_ready:
        return WRITE_REVIEW_REQUIRED, R_STRONG_MISSING_REQUIRED, ()
    return WRITE_READY, R_STRONG_REQUIRED_MATCH, ()


def assess_alignment_row(
    row, *, source_readiness: str, target_value: object, policy: WriteReadinessPolicy,
    locked_override: bool | None | object = _UNSET,
) -> tuple[str, str, str, tuple[str, ...]]:
    """(write_status, write_reason, expected_value_type, missing_evidence).

    ``locked_override`` re-valida la protección de la celda-target contra el
    workbook (defensa en profundidad, §17): si se pasa, gana sobre
    ``row.target_locked``. NO evalúa colisiones (se hacen globalmente después).
    """
    evt = policy.expected_value_type(row.source_kind)
    locked = row.target_locked if locked_override is _UNSET else locked_override

    if source_readiness not in policy.allowed_readiness:
        return NOT_WRITABLE, R_SOURCE_REVIEW_REQUIRED, evt, ()

    if row.match_status == _MATCH_AMBIGUOUS:
        return WRITE_BLOCKED, R_MATCH_AMBIGUOUS, evt, ()
    if row.match_status == _MATCH_CONFLICT:
        return WRITE_BLOCKED, R_DIMENSION_CONFLICT, evt, ()

    # --- target: rol / tipo -------------------------------------
    if row.target_kind == _KIND_FORMULA:
        return NOT_WRITABLE, R_TARGET_IS_FORMULA, evt, ()
    if row.target_kind not in policy.allowed_target_kind:
        return NOT_WRITABLE, R_TARGET_NOT_INPUT_CELL, evt, ()
    if row.target_alignment_role not in policy.allowed_target_role:
        return NOT_WRITABLE, R_TARGET_ROLE_NOT_INPUT, evt, ()

    # --- protección (revalidación defensiva) --------------------
    if locked is not False:
        return WRITE_BLOCKED, R_TARGET_LOCKED, evt, ()

    # --- fuente esperada del target ---------------------------
    if row.target_source_expectation in _NON_MEDINET_EXPECTATIONS:
        return WRITE_BLOCKED, R_TARGET_SOURCE_NOT_MEDINET, evt, ()
    if row.target_source_expectation not in policy.allowed_expectation:
        # UNKNOWN: el alignment resuelto no promovió a MEDINET -> revisión (§8)
        return WRITE_REVIEW_REQUIRED, R_TARGET_SOURCE_UNKNOWN, evt, ()

    # --- tipo de valor esperado ------------------------------
    evt, compatible = expected_value_type_for(row.source_kind, target_value, policy)
    if not compatible:
        return policy.evt_on_incompatible, R_EXPECTED_VALUE_TYPE_INCOMPATIBLE, evt, ()

    # --- evidencia de match --------------------------------
    if row.match_status == _MATCH_EXACT:
        if policy.exact_write_ready:
            return WRITE_READY, R_EXACT_ALL_MATCH, evt, ()
        return WRITE_REVIEW_REQUIRED, R_EXACT_ALL_MATCH, evt, ()
    if row.match_status == _MATCH_STRONG:
        status, reason, missing = _strong_assessment(row, policy)
        return status, reason, evt, missing

    return WRITE_REVIEW_REQUIRED, "UNHANDLED_MATCH_STATUS", evt, ()


# ---------------------------------------------------------------------------
# Colisiones
# ---------------------------------------------------------------------------


def detect_collisions(candidates: Sequence[WriteReadiness]) -> list[Collision]:
    """Colisiones entre candidatos a escritura (WRITE_READY / WRITE_REVIEW_REQUIRED).

    Nunca "last write wins": cualquier colisión bloquea a TODOS los implicados.
    """
    by_target: dict[tuple[str, str], list[str]] = defaultdict(list)
    by_source: dict[str, set[tuple[str, str]]] = defaultdict(set)
    by_instruction: dict[str, list[str]] = defaultdict(list)
    for c in candidates:
        by_target[(c.target_sheet, c.target_cell)].append(c.source_metric_id)
        by_source[c.source_metric_id].add((c.target_sheet, c.target_cell))
        if c.instruction is not None:
            by_instruction[c.instruction.instruction_id].append(c.source_metric_id)

    out: list[Collision] = []
    for (sheet, cell), sources in sorted(by_target.items()):
        uniq = sorted(set(sources))
        if len(uniq) > 1:
            out.append(Collision(
                COLL_MULTI_SOURCE_SAME_TARGET, f"{sheet}!{cell}", tuple(uniq),
                f"{len(uniq)} sources apuntan a la misma celda",
            ))
    for source, targets in sorted(by_source.items()):
        if len(targets) > 1:
            out.append(Collision(
                COLL_SAME_SOURCE_MULTI_TARGET, source,
                tuple(f"{s}!{c}" for s, c in sorted(targets)),
                f"un source resuelto a {len(targets)} celdas distintas",
            ))
    for iid, sources in sorted(by_instruction.items()):
        if len(sources) > 1:
            out.append(Collision(
                COLL_DUPLICATE_INSTRUCTION_ID, iid, tuple(sorted(set(sources))),
                "misma identidad de instrucción para varios sources",
            ))
    return out


# ---------------------------------------------------------------------------
# Orquestación
# ---------------------------------------------------------------------------

_EXCLUSION_STATUS = {
    _EXCL_BLOCKED_CONFLICT: (WRITE_BLOCKED, R_SOURCE_BLOCKED_CONFLICT),
    _EXCL_REVIEW_REQUIRED: (NOT_WRITABLE, R_SOURCE_REVIEW_REQUIRED),
    _EXCL_ROLLUP: (NOT_WRITABLE, R_ROLLUP_COMPUTED_BY_TEMPLATE),
    _EXCL_NOT_MEDINET: (NOT_WRITABLE, R_SOURCE_NOT_MEDINET),
}


def _evidence_summary(row) -> str:
    return (
        f"score={row.match_score} section={row.section_match} row={row.row_match} "
        f"col={row.column_match} sex={row.sex_match} age={row.age_match} "
        f"proc={row.procedure_match} expect={row.target_source_expectation} "
        f"locked={row.target_locked}"
    )


def build_write_mapping(
    alignment_report,
    template_fingerprint: dict,
    policy: WriteReadinessPolicy,
    *,
    target_values: dict[tuple[str, str], object] | None = None,
    locked_by_target: dict[tuple[str, str], bool | None] | None = None,
) -> WriteMappingReport:
    """Construye el write mapping desde un `AlignmentReport` (Sprint 3.5).

    ``locked_by_target`` re-valida la protección de cada celda-target contra el
    workbook (defensa en profundidad, §17).
    """
    tfp_id = template_fingerprint["structural_template_fingerprint_id"]
    target_values = target_values or {}
    locked_by_target = locked_by_target or {}
    report = WriteMappingReport(
        policy_version=policy.version,
        template_fingerprint=template_fingerprint,
        zero_write_policy=policy.zero_write_policy,
    )

    ambiguous_candidates: dict[str, tuple[str, ...]] = {
        a.source_metric_id: a.candidate_cells for a in alignment_report.ambiguous
    }

    provisional: list[WriteReadiness] = []
    for row in alignment_report.rows:
        key = (row.target_sheet, row.target_cell)
        tv = target_values.get(key)
        locked = locked_by_target.get(key, _UNSET)
        status, reason, evt, missing = assess_alignment_row(
            row, source_readiness=row.source_readiness, target_value=tv, policy=policy,
            locked_override=locked,
        )
        effective_locked = row.target_locked if locked is _UNSET else locked
        instr = None
        if status in (WRITE_READY, WRITE_REVIEW_REQUIRED):
            instr = WriteInstruction(
                instruction_id=instruction_id(
                    row.source_semantic_signature, row.target_semantic_signature, tfp_id
                ),
                source_metric_id=row.source_metric_id,
                source_semantic_signature=row.source_semantic_signature,
                target_sheet=row.target_sheet,
                target_cell=row.target_cell,
                target_semantic_signature=row.target_semantic_signature,
                template_fingerprint_id=tfp_id,
                source_form=row.source_form,
                source_kind=row.source_kind,
                match_status=row.match_status,
                write_status=status,
                expected_value_type=evt,
                evidence=(reason, *missing),
                policy_version=policy.version,
                needs_human_review=status != WRITE_READY,
            )
        extra = ""
        if row.match_status == _MATCH_AMBIGUOUS:
            extra = " candidates=" + "|".join(
                ambiguous_candidates.get(row.source_metric_id, ())
            )
        provisional.append(WriteReadiness(
            source_metric_id=row.source_metric_id,
            source_semantic_signature=row.source_semantic_signature,
            source_readiness=row.source_readiness,
            source_form=row.source_form,
            source_kind=row.source_kind,
            match_status=row.match_status,
            target_sheet=row.target_sheet,
            target_cell=row.target_cell,
            target_semantic_signature=row.target_semantic_signature,
            target_kind=row.target_kind,
            target_alignment_role=row.target_alignment_role,
            target_source_expectation=row.target_source_expectation,
            target_locked=effective_locked,
            write_status=status,
            write_reason=reason,
            expected_value_type=evt,
            missing_evidence=missing,
            evidence_summary=_evidence_summary(row) + extra,
            needs_human_review=status != WRITE_READY,
            instruction=instr,
        ))

    # --- colisiones: bloquean a TODOS los implicados ------------
    candidate_pool = [w for w in provisional if w.write_status in (WRITE_READY, WRITE_REVIEW_REQUIRED)]
    collisions = detect_collisions(candidate_pool)
    report.collisions = collisions
    colliding_sources: set[str] = set()
    colliding_targets: set[tuple[str, str]] = set()
    for coll in collisions:
        if coll.collision_type == COLL_MULTI_SOURCE_SAME_TARGET:
            sheet, cell = coll.key.split("!", 1)
            colliding_targets.add((sheet, cell))
            colliding_sources.update(coll.involved)
        else:
            colliding_sources.update(coll.involved)

    final: list[WriteReadiness] = []
    for w in provisional:
        if (
            w.write_status in (WRITE_READY, WRITE_REVIEW_REQUIRED)
            and (w.source_metric_id in colliding_sources
                 or (w.target_sheet, w.target_cell) in colliding_targets)
        ):
            final.append(_downgrade(w, WRITE_BLOCKED, R_WRITE_COLLISION))
        else:
            final.append(w)

    # --- sources no alineados / excluidos -----------------------
    aligned_ids = {w.source_metric_id for w in final}
    for u in alignment_report.unmatched_sources:
        if u.metric_id in aligned_ids:
            continue
        final.append(_bare_readiness(
            u.metric_id, u.form, "", "AUTO_READY", "NO_MATCH",
            NOT_WRITABLE, R_NO_ALIGNMENT_FOUND, policy,
        ))
    for e in alignment_report.excluded_sources:
        status, reason = _EXCLUSION_STATUS.get(
            e.reason, (NOT_WRITABLE, "SOURCE_NOT_ELIGIBLE")
        )
        final.append(_bare_readiness(
            e.metric_id, e.form, "", e.readiness_status, "EXCLUDED",
            status, reason, policy,
        ))

    report.readiness = sorted(final, key=lambda w: w.source_metric_id)
    report.instructions = sorted(
        (w.instruction for w in report.readiness
         if w.instruction is not None and w.write_status == WRITE_READY),
        key=lambda i: i.instruction_id,
    )
    return report


def _downgrade(w: WriteReadiness, status: str, reason: str) -> WriteReadiness:
    instr = None
    if w.instruction is not None:
        instr = replace(w.instruction, write_status=status,
                        evidence=(reason, *w.instruction.evidence),
                        needs_human_review=True)
    return replace(
        w, write_status=status, write_reason=reason, needs_human_review=True,
        instruction=instr if status != WRITE_BLOCKED else None,
    )


def _bare_readiness(mid, form, sig, source_readiness, match_status, status, reason,
                    policy) -> WriteReadiness:
    return WriteReadiness(
        source_metric_id=mid, source_semantic_signature=sig,
        source_readiness=source_readiness, source_form=form, source_kind="",
        match_status=match_status, target_sheet="", target_cell="",
        target_semantic_signature="", target_kind="", target_alignment_role="",
        target_source_expectation="", target_locked=None,
        write_status=status, write_reason=reason,
        expected_value_type=policy.evt_default, missing_evidence=(),
        evidence_summary="", needs_human_review=status != NOT_WRITABLE,
    )


# ---------------------------------------------------------------------------
# Clusters de la review queue
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ReviewCluster:
    cluster_id: str
    form: str
    match_status: str
    write_reason: str
    missing_evidence: str
    metric_count: int
    example_metric_ids: tuple[str, ...]


def build_review_clusters(
    readiness: Iterable[WriteReadiness], *, max_examples: int = 5
) -> list[ReviewCluster]:
    groups: dict[tuple[str, str, str, str], list[str]] = defaultdict(list)
    for w in readiness:
        if w.write_status != WRITE_REVIEW_REQUIRED:
            continue
        key = (w.source_form, w.match_status, w.write_reason, "|".join(w.missing_evidence))
        groups[key].append(w.source_metric_id)
    out: list[ReviewCluster] = []
    for (form, match_status, reason, missing), ids in sorted(groups.items()):
        ordered = sorted(ids)
        out.append(ReviewCluster(
            cluster_id=f"{form}:{match_status}:{reason}:{missing or '-'}",
            form=form, match_status=match_status, write_reason=reason,
            missing_evidence=missing, metric_count=len(ordered),
            example_metric_ids=tuple(ordered[:max_examples]),
        ))
    return out


def coverage_counts(readiness: Iterable[WriteReadiness]) -> Counter:
    return Counter(w.write_status for w in readiness)
