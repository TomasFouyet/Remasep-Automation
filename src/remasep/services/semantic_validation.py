"""Capa de VALIDACIÓN / READINESS sobre el mapeo semántico (Sprint 3.3).

Sprint 3.2 produjo `SemanticMetric` con dimensiones (`sex`, `age`,
`aggregation_scope`, `procedure`) y `mapping_status` ∈
``CONFIRMED / PARTIAL / CONFLICT``. Sprint 3.3 **no cambia esas inferencias**:
añade una respuesta técnica y auditable por métrica —

    ¿puede usarse automáticamente?  ¿necesita revisión?  ¿debe bloquearse?

`readiness_status` ∈ ``AUTO_READY · REVIEW_REQUIRED · BLOCKED_CONFLICT ·
NOT_APPLICABLE`` (**distinto** de `mapping_status`). Reglas transparentes y
versionadas en ``config/semantic_validation_2026/policy.yaml``.

`AUTO_READY` = "suficiente evidencia técnica según la policy actual". **No**
significa "validado MINSAL" ni "aprobado por el cliente".

Los 5 conflictos reales del workbook (`SEX_LABEL_FORMULA_MISMATCH`) quedan
`BLOCKED_CONFLICT` y generan issues con `recommended_action = HUMAN_REVIEW`.
**Nunca** se propone "cambiar a MALE/FEMALE": no se toma partido.
"""

from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

import yaml

from remasep.core.errors import RemasepError
from remasep.services.semantic_mapping import SemanticMetric

# --- estados de readiness --------------------------------------------

AUTO_READY = "AUTO_READY"
REVIEW_REQUIRED = "REVIEW_REQUIRED"
BLOCKED_CONFLICT = "BLOCKED_CONFLICT"
NOT_APPLICABLE = "NOT_APPLICABLE"

# --- severidad de issue --------------------------------------------

SEV_BLOCKING = "BLOCKING"
SEV_REVIEW = "REVIEW"
SEV_INFO = "INFO"

RECOMMENDED_ACTION = "HUMAN_REVIEW"
ISSUE_STATUS_OPEN = "OPEN"

_DIM_STATUS_RESOLVED = "RESOLVED"
_DIM_STATUS_MISSING = "MISSING"
_BUCKET_DETAIL = "DETAIL"
_BUCKET_ROLLUP = "ROLLUP"
# dimensiones que se evalúan (orden estable para issues/razones)
_DIMENSIONS = ("form", "row_path", "aggregation_scope", "sex", "age", "procedure")

_DEFAULT_CONFIG_DIR = (
    Path(__file__).resolve().parents[3] / "config" / "semantic_validation_2026"
)


class SemanticValidationError(RemasepError):
    """Configuración de validación semántica inválida."""


# ---------------------------------------------------------------------------
# Configuración versionada
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ReadinessPolicy:
    version: str
    rollup_scopes: frozenset[str]
    required_default: dict[str, tuple[str, ...]]  # bucket -> dims
    required_by_form: dict[str, dict[str, tuple[str, ...]]]
    acceptable_status: dict[str, frozenset[str]]  # dimension -> estados aceptables
    severity: dict[str, str]

    def required_dimensions(self, form: str, bucket: str) -> tuple[str, ...]:
        by_form = self.required_by_form.get(form, {})
        return by_form.get(bucket, self.required_default.get(bucket, ()))

    def acceptable(self, dimension: str) -> frozenset[str]:
        return self.acceptable_status.get(dimension, self.acceptable_status["default"])


@dataclass(frozen=True)
class OverrideDecision:
    metric_signature: str
    cell: str
    decision: str
    reason: str
    evidence: str
    approved_by: str
    approved_date: str


def _load_yaml(path: Path) -> dict:
    if not path.is_file():
        raise SemanticValidationError(f"config de validación no encontrada: {path}")
    return yaml.safe_load(path.read_text(encoding="utf-8")) or {}


def load_readiness_policy(config_dir: str | Path | None = None) -> ReadinessPolicy:
    base = Path(config_dir) if config_dir is not None else _DEFAULT_CONFIG_DIR
    raw = _load_yaml(base / "policy.yaml")

    req = raw.get("required_dimensions") or {}
    required_default = {
        bucket: tuple(dims) for bucket, dims in (req.get("default") or {}).items()
    }
    required_by_form = {
        form: {bucket: tuple(dims) for bucket, dims in buckets.items()}
        for form, buckets in (req.get("by_form") or {}).items()
    }
    acceptable = {
        dim: frozenset(states)
        for dim, states in (raw.get("acceptable_status") or {}).items()
    }
    acceptable.setdefault("default", frozenset({"CONFIRMED", "NOT_APPLICABLE"}))

    return ReadinessPolicy(
        version=str(raw.get("version", "semantic_validation_2026")),
        rollup_scopes=frozenset(raw.get("rollup_scopes") or ["TOTAL", "SUBTOTAL"]),
        required_default=required_default,
        required_by_form=required_by_form,
        acceptable_status=acceptable,
        severity={str(k): str(v) for k, v in (raw.get("severity") or {}).items()},
    )


def load_overrides(config_dir: str | Path | None = None) -> dict[str, OverrideDecision]:
    """Decisiones humanas por ``metric_signature``. Vacío por diseño en 3.3."""
    base = Path(config_dir) if config_dir is not None else _DEFAULT_CONFIG_DIR
    raw = _load_yaml(base / "overrides.yaml")
    out: dict[str, OverrideDecision] = {}
    for entry in raw.get("overrides") or []:
        signature = str(entry.get("metric_signature", "")).strip()
        if not signature:
            raise SemanticValidationError("override sin `metric_signature`")
        out[signature] = OverrideDecision(
            metric_signature=signature,
            cell=str(entry.get("cell", "")),
            decision=str(entry.get("decision", "")),
            reason=str(entry.get("reason", "")),
            evidence=str(entry.get("evidence", "")),
            approved_by=str(entry.get("approved_by", "")),
            approved_date=str(entry.get("approved_date", "")),
        )
    return out


def override_key(metric: SemanticMetric) -> str:
    """Identidad preferida de un override: la firma semántica (no la coordenada)."""
    return metric.semantic_signature


# ---------------------------------------------------------------------------
# Modelo
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SemanticReviewIssue:
    issue_id: str
    metric_id: str
    form: str
    sheet: str
    cell: str
    dimension: str
    issue_type: str
    severity: str
    label_evidence: str
    formula_evidence: str
    recommended_action: str = RECOMMENDED_ACTION
    status: str = ISSUE_STATUS_OPEN

    def evidence_summary(self) -> str:
        parts = []
        if self.label_evidence:
            parts.append(f"label={self.label_evidence}")
        if self.formula_evidence:
            parts.append(f"formula={self.formula_evidence}")
        return " ".join(parts)


@dataclass
class SemanticReadiness:
    metric_id: str
    source: str
    form: str
    sheet: str
    cell: str
    kind: str
    mapping_status: str
    readiness_status: str
    review_reasons: tuple[str, ...]
    issues: tuple[SemanticReviewIssue, ...]
    sex_status: str
    age_status: str
    aggregation_scope_status: str
    procedure_status: str
    semantic_signature: str
    override_matched: bool = False


@dataclass
class ValidationResult:
    policy_version: str
    readiness: list[SemanticReadiness] = field(default_factory=list)
    override_signatures: tuple[str, ...] = ()


# ---------------------------------------------------------------------------
# Evaluación
# ---------------------------------------------------------------------------


def _dimension_status(metric: SemanticMetric) -> dict[str, str]:
    return {
        "form": _DIM_STATUS_RESOLVED if metric.form else _DIM_STATUS_MISSING,
        "row_path": _DIM_STATUS_RESOLVED if metric.row_path_raw else _DIM_STATUS_MISSING,
        "aggregation_scope": metric.aggregation_scope_status,
        "sex": metric.sex_status,
        "age": metric.age_status,
        "procedure": metric.procedure_status,
    }


def _conflict_evidence(metric: SemanticMetric, dimension: str) -> tuple[str, str]:
    for conflict in metric.conflicts:
        if conflict.dimension.upper() == dimension.upper():
            return conflict.label_evidence, conflict.formula_evidence
    return "", ""


def assess_readiness(
    metric: SemanticMetric,
    policy: ReadinessPolicy,
    overrides: dict[str, OverrideDecision] | None = None,
) -> SemanticReadiness:
    overrides = overrides or {}
    bucket = _BUCKET_ROLLUP if metric.aggregation_scope in policy.rollup_scopes else _BUCKET_DETAIL
    required = set(policy.required_dimensions(metric.form, bucket))
    dim_status = _dimension_status(metric)

    issues: list[SemanticReviewIssue] = []

    def _add(dimension: str, issue_type: str, severity: str, label: str, formula: str) -> None:
        dimension = dimension.upper()
        issues.append(
            SemanticReviewIssue(
                issue_id=f"{metric.form}::{metric.cell}::{dimension}",
                metric_id=metric.metric_id,
                form=metric.form,
                sheet=metric.sheet,
                cell=metric.cell,
                dimension=dimension,
                issue_type=issue_type,
                severity=severity,
                label_evidence=label,
                formula_evidence=formula,
            )
        )

    # 1. CONFLICT en cualquier dimensión -> siempre BLOCKING (nunca se corrige).
    for conflict in metric.conflicts:
        _add(
            conflict.dimension, conflict.conflict_type, SEV_BLOCKING,
            conflict.label_evidence, conflict.formula_evidence,
        )

    # 2. dimensiones no resueltas: REVIEW si son requeridas, INFO si no.
    for dimension in _DIMENSIONS:
        status = dim_status[dimension]
        if status == "CONFLICT":
            continue  # ya cubierto en (1)
        if status in policy.acceptable(dimension):
            continue
        if status not in ("LABEL_ONLY", "FORMULA_ONLY", "UNRESOLVED", _DIM_STATUS_MISSING):
            continue
        is_required = dimension in required
        severity = SEV_REVIEW if is_required else SEV_INFO
        label, formula = _conflict_evidence(metric, dimension)
        _add(dimension, f"{dimension.upper()}_{status}", severity, label, formula)

    blocking = [i for i in issues if i.severity == SEV_BLOCKING]
    review = [i for i in issues if i.severity == SEV_REVIEW]

    if blocking:
        readiness_status = BLOCKED_CONFLICT
    elif bucket == _BUCKET_ROLLUP:
        readiness_status = NOT_APPLICABLE
    elif review:
        readiness_status = REVIEW_REQUIRED
    else:
        readiness_status = AUTO_READY

    review_reasons = tuple(
        sorted({i.issue_type for i in issues if i.severity in (SEV_BLOCKING, SEV_REVIEW)})
    )

    # Overrides: sólo arquitectura en 3.3. Se registra la coincidencia por firma
    # semántica; la RESOLUCIÓN (cambiar el readiness) no se aplica todavía.
    override_matched = override_key(metric) in overrides

    return SemanticReadiness(
        metric_id=metric.metric_id,
        source=metric.source,
        form=metric.form,
        sheet=metric.sheet,
        cell=metric.cell,
        kind=metric.kind,
        mapping_status=metric.mapping_status,
        readiness_status=readiness_status,
        review_reasons=review_reasons,
        issues=tuple(issues),
        sex_status=metric.sex_status,
        age_status=metric.age_status,
        aggregation_scope_status=metric.aggregation_scope_status,
        procedure_status=metric.procedure_status,
        semantic_signature=metric.semantic_signature,
        override_matched=override_matched,
    )


def assess_all(
    metrics: list[SemanticMetric],
    policy: ReadinessPolicy,
    overrides: dict[str, OverrideDecision] | None = None,
) -> ValidationResult:
    overrides = overrides or {}
    return ValidationResult(
        policy_version=policy.version,
        readiness=[assess_readiness(m, policy, overrides) for m in metrics],
        override_signatures=tuple(sorted(overrides)),
    )


# ---------------------------------------------------------------------------
# Clusters de revisión
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ReviewCluster:
    cluster_id: str
    form: str
    dimension: str
    issue_type: str
    severity: str
    metric_count: int
    example_metric_ids: tuple[str, ...]


def build_review_clusters(
    result: ValidationResult, *, max_examples: int = 5
) -> list[ReviewCluster]:
    groups: dict[tuple[str, str, str, str], list[str]] = {}
    for readiness in result.readiness:
        for issue in readiness.issues:
            key = (issue.form, issue.dimension, issue.issue_type, issue.severity)
            groups.setdefault(key, []).append(issue.metric_id)
    clusters: list[ReviewCluster] = []
    for (form, dimension, issue_type, severity), ids in sorted(groups.items()):
        ordered = sorted(ids)
        clusters.append(
            ReviewCluster(
                cluster_id=f"{form}:{dimension}:{issue_type}:{severity}",
                form=form,
                dimension=dimension,
                issue_type=issue_type,
                severity=severity,
                metric_count=len(ordered),
                example_metric_ids=tuple(ordered[:max_examples]),
            )
        )
    return clusters
