"""Reconciliación del input de producción Medinet vs. la referencia legacy.

Contrato de input (ver `docs/MEDINET_INPUT_CONTRACT.md`):

- **PRODUCCIÓN**: el export directo de Medinet **"Detalle de citas"**
  (p.ej. ``detalle_citas - 2026-09-07T123630.940.xlsx``). Es el único input real
  de la aplicación.
- **REFERENCIA LEGACY**: ``GENERACION DATOS REMASEP.xlsx``. Sólo se usa para
  ingeniería inversa; la app **no** depende de él.

Este módulo compara, de forma **privacy-safe** (sólo campos semánticos, nunca
RUN / nombre / teléfono / id de paciente), el export directo filtrado a un
período contra la hoja de detalle legacy, para explicar la diferencia de
recuento (2114 vs 1364 en julio 2026).

Comparación por **multiset** (conteos), nunca asumiendo huella única: un mismo
paciente puede tener varias citas con la misma huella semántica.
"""

from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field
from pathlib import Path

import pandas as pd

from remasep.adapters.medinet import read_medinet
from remasep.services.common import Period

# Campos semánticos para la huella. NADA identificatorio.
FINGERPRINT_FIELDS: tuple[str, ...] = (
    "DIA_CITA", "FECHA_NACIMIENTO", "SEXO", "SUCURSAL", "ESPECIALIDAD",
    "TIPO_DE_CITA", "PRESTACION", "ESTADO", "MODALIDAD", "PRESTACION_REALIZADA",
)
FINGERPRINT_FIELDS_NO_ESTADO: tuple[str, ...] = tuple(
    f for f in FINGERPRINT_FIELDS if f != "ESTADO"
)
_DATE_FIELDS = ("DIA_CITA", "FECHA_NACIMIENTO")

# Hipótesis a comprobar (NO regla funcional): estados que el generador legacy
# conserva vs. descarta. Ver `estado_reconstruction`.
LEGACY_KEPT_STATES: tuple[str, ...] = (
    "Atendido", "Atención Pausada", "En Sala de Espera", "En Atención",
)
LEGACY_EXCLUDED_STATES: tuple[str, ...] = (
    "Cancelado", "No Se Presenta", "Agendado", "Confirmado", "Re-Agendado",
)


# ---------------------------------------------------------------------------
# Carga
# ---------------------------------------------------------------------------


@dataclass
class LoadedRecords:
    source_name: str
    sheet_name: str
    frame: pd.DataFrame          # sólo columnas semánticas
    structurally_valid: pd.Series  # bool: DIA_CITA parseable (no fila arrastrada)


def load_records(path: str | Path) -> LoadedRecords:
    medinet = read_medinet(path)
    frame = medinet.frame
    valid = frame["DIA_CITA"].notna()
    return LoadedRecords(
        source_name=Path(path).name,
        sheet_name=medinet.sheet_name,
        frame=frame,
        structurally_valid=valid.fillna(False).astype(bool),
    )


# ---------------------------------------------------------------------------
# Alcance de período
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class PeriodScopeBreakdown:
    physical_records: int
    structurally_valid_records: int
    structurally_invalid_records: int
    in_period_records: int
    out_of_period_records: int
    processing_scope_records: int

    def __post_init__(self) -> None:
        if self.structurally_valid_records + self.structurally_invalid_records != self.physical_records:
            raise ValueError("PeriodScopeBreakdown inconsistente (valid + invalid != physical)")
        if self.in_period_records + self.out_of_period_records != self.structurally_valid_records:
            raise ValueError("PeriodScopeBreakdown inconsistente (in + out != valid)")
        if self.processing_scope_records != self.in_period_records:
            raise ValueError("PeriodScopeBreakdown inconsistente (scope != in_period)")


def period_scope(records: LoadedRecords, period: Period) -> PeriodScopeBreakdown:
    frame = records.frame
    valid = records.structurally_valid
    dia = frame["DIA_CITA"]
    in_period = (valid & dia.dt.year.eq(period.year) & dia.dt.month.eq(period.month)).fillna(
        False
    ).astype(bool)
    valid_n = int(valid.sum())
    in_n = int(in_period.sum())
    return PeriodScopeBreakdown(
        physical_records=len(frame),
        structurally_valid_records=valid_n,
        structurally_invalid_records=len(frame) - valid_n,
        in_period_records=in_n,
        out_of_period_records=valid_n - in_n,
        processing_scope_records=in_n,
    )


def in_period_frame(records: LoadedRecords, period: Period) -> pd.DataFrame:
    frame = records.frame
    dia = frame["DIA_CITA"]
    mask = (
        records.structurally_valid
        & dia.dt.year.eq(period.year)
        & dia.dt.month.eq(period.month)
    ).fillna(False).astype(bool)
    return frame.loc[mask].reset_index(drop=True)


# ---------------------------------------------------------------------------
# Huellas privacy-safe
# ---------------------------------------------------------------------------


def _norm_value(value: object, is_date: bool) -> str:
    if value is None or (isinstance(value, float) and pd.isna(value)):
        return ""
    if is_date:
        if pd.isna(value):
            return ""
        return pd.Timestamp(value).date().isoformat()
    text = str(value).strip()
    return "" if text.lower() in {"", "nan", "nat", "none", "<na>"} else text


def fingerprint_series(frame: pd.DataFrame, fields: tuple[str, ...]) -> pd.Series:
    """Serie de huellas ``campo|campo|…`` (sin identificadores)."""
    parts = []
    for name in fields:
        is_date = name in _DATE_FIELDS
        parts.append(frame[name].map(lambda v, d=is_date: _norm_value(v, d)))
    joined = parts[0].astype(str)
    for extra in parts[1:]:
        joined = joined.str.cat(extra.astype(str), sep="|")
    return joined


# ---------------------------------------------------------------------------
# Comparación multiset
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class SubsetComparison:
    fields: tuple[str, ...]
    subset_records: int          # legacy
    superset_records: int        # export directo del período
    matched_subset_records: int  # legacy que aparece (multiset) en el export
    unmatched_subset_records: int  # legacy SIN correspondencia en el export
    superset_excess_records: int  # export del período que NO cubre el legacy
    is_exact_subset: bool         # ¿legacy ⊆ export directo (multiset)?


def multiset_subset_check(
    subset_frame: pd.DataFrame, superset_frame: pd.DataFrame, fields: tuple[str, ...]
) -> SubsetComparison:
    sub = Counter(fingerprint_series(subset_frame, fields))
    sup = Counter(fingerprint_series(superset_frame, fields))
    matched = sum((sub & sup).values())
    unmatched = sum((sub - sup).values())
    excess = sum((sup - sub).values())
    return SubsetComparison(
        fields=fields,
        subset_records=int(sub.total()),
        superset_records=int(sup.total()),
        matched_subset_records=matched,
        unmatched_subset_records=unmatched,
        superset_excess_records=excess,
        is_exact_subset=unmatched == 0,
    )


# ---------------------------------------------------------------------------
# Agregados privacy-safe de los "extras"
# ---------------------------------------------------------------------------


def categorical_aggregate(frame: pd.DataFrame, field_name: str) -> list[tuple[str, int]]:
    if field_name not in frame.columns:
        return []
    values = frame[field_name].map(lambda v: _norm_value(v, False))
    counts = Counter(v if v else "<vacío>" for v in values)
    return sorted(counts.items(), key=lambda kv: (-kv[1], kv[0]))


def direct_export_excess_frame(
    direct_period: pd.DataFrame, legacy_active: pd.DataFrame
) -> pd.DataFrame:
    """Filas del export directo del período cuya huella SIN estado no está
    representada (multiset) en el detalle legacy — los registros que el proceso
    antiguo dejó fuera."""
    legacy_fp = Counter(fingerprint_series(legacy_active, FINGERPRINT_FIELDS_NO_ESTADO))
    direct_fp = fingerprint_series(direct_period, FINGERPRINT_FIELDS_NO_ESTADO)
    remaining = Counter(legacy_fp)
    keep_row: list[bool] = []
    for fp in direct_fp:
        if remaining.get(fp, 0) > 0:
            remaining[fp] -= 1
            keep_row.append(False)  # empareja un registro legacy -> no es "extra"
        else:
            keep_row.append(True)
    return direct_period.loc[pd.Series(keep_row, index=direct_period.index)].reset_index(
        drop=True
    )


# ---------------------------------------------------------------------------
# Reconstrucción por ESTADO (hipótesis, NO regla)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class EstadoReconstruction:
    legacy_active_records: int
    kept_states: tuple[str, ...]
    excluded_states: tuple[str, ...]
    direct_period_by_estado: tuple[tuple[str, int], ...]
    direct_kept_records: int      # export directo del período con estado "kept"
    direct_excluded_records: int
    reproduces_legacy_count_exactly: bool
    status: str = "CURRENT_LEGACY_BEHAVIOR_PENDING_FUNCTIONAL_CONFIRMATION"


def estado_reconstruction(
    direct_period: pd.DataFrame,
    legacy_active_records: int,
    *,
    kept_states: tuple[str, ...] = LEGACY_KEPT_STATES,
    excluded_states: tuple[str, ...] = LEGACY_EXCLUDED_STATES,
) -> EstadoReconstruction:
    by_estado = categorical_aggregate(direct_period, "ESTADO")
    kept_norm = {s.strip().casefold() for s in kept_states}
    excl_norm = {s.strip().casefold() for s in excluded_states}
    kept = sum(c for v, c in by_estado if v.strip().casefold() in kept_norm)
    excluded = sum(c for v, c in by_estado if v.strip().casefold() in excl_norm)
    return EstadoReconstruction(
        legacy_active_records=legacy_active_records,
        kept_states=kept_states,
        excluded_states=excluded_states,
        direct_period_by_estado=tuple(by_estado),
        direct_kept_records=kept,
        direct_excluded_records=excluded,
        reproduces_legacy_count_exactly=kept == legacy_active_records,
    )


# ---------------------------------------------------------------------------
# Orquestación
# ---------------------------------------------------------------------------


@dataclass
class ReconciliationResult:
    period: Period
    direct_source: str
    legacy_source: str
    direct_scope: PeriodScopeBreakdown
    legacy_scope: PeriodScopeBreakdown
    subset_full_fp: SubsetComparison
    subset_no_estado_fp: SubsetComparison
    estado: EstadoReconstruction
    extra_aggregates: dict[str, list[tuple[str, int]]] = field(default_factory=dict)
    notes: tuple[str, ...] = ()


def reconcile(
    direct_export_path: str | Path,
    legacy_reference_path: str | Path,
    period: Period,
) -> ReconciliationResult:
    direct = load_records(direct_export_path)
    legacy = load_records(legacy_reference_path)

    direct_scope = period_scope(direct, period)
    legacy_scope = period_scope(legacy, period)

    direct_period = in_period_frame(direct, period)
    legacy_active = legacy.frame.loc[legacy.structurally_valid].reset_index(drop=True)

    full = multiset_subset_check(legacy_active, direct_period, FINGERPRINT_FIELDS)
    no_estado = multiset_subset_check(
        legacy_active, direct_period, FINGERPRINT_FIELDS_NO_ESTADO
    )
    estado = estado_reconstruction(direct_period, len(legacy_active))

    excess = direct_export_excess_frame(direct_period, legacy_active)
    aggregates = {
        "ESTADO": categorical_aggregate(excess, "ESTADO"),
        "TIPO_DE_CITA": categorical_aggregate(excess, "TIPO_DE_CITA"),
        "SUCURSAL": categorical_aggregate(excess, "SUCURSAL"),
        "MODALIDAD": categorical_aggregate(excess, "MODALIDAD"),
        "ESPECIALIDAD": categorical_aggregate(excess, "ESPECIALIDAD"),
    }

    notes = []
    if no_estado.is_exact_subset:
        notes.append(
            f"Los {no_estado.subset_records} registros activos del detalle legacy son "
            f"un subconjunto EXACTO (multiset, huella sin ESTADO) de los "
            f"{no_estado.superset_records} registros del export directo de "
            f"{period.label}."
        )
    if not full.is_exact_subset:
        notes.append(
            f"Con ESTADO en la huella, {full.unmatched_subset_records} registro(s) "
            "legacy no casan exactamente: el export directo es posterior al snapshot "
            "legacy y algún ESTADO progresó (p.ej. 'En Sala de Espera' -> 'Atendido')."
        )
    if estado.reproduces_legacy_count_exactly:
        notes.append(
            f"El subconjunto de estados {list(estado.kept_states)} del export directo "
            f"de {period.label} reproduce EXACTAMENTE los {estado.legacy_active_records} "
            "registros activos del detalle legacy. Hipótesis marcada "
            f"{estado.status}; NO se implementa todavía como regla."
        )

    return ReconciliationResult(
        period=period,
        direct_source=direct.source_name,
        legacy_source=legacy.source_name,
        direct_scope=direct_scope,
        legacy_scope=legacy_scope,
        subset_full_fp=full,
        subset_no_estado_fp=no_estado,
        estado=estado,
        extra_aggregates=aggregates,
        notes=tuple(notes),
    )
