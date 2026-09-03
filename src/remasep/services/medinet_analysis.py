"""Análisis REAL de un export Medinet (Sprint 2.1).

Lee un archivo Medinet, valida estructura y registros, detecta el período,
calcula edades con la semántica legacy y aplica las 27 reglas candidatas de
Sprint 1.3. **No** genera REMASEP, no escribe Excel, no usa COM.

Privacidad: los valores de las filas se procesan solo en memoria. Nada que salga
de aquí (problems, validations, diagnostics, notes) contiene RUN, nombre, fecha
de nacimiento individual ni ningún dato personal: solo fila, campo y tipo de
problema, más agregados categóricos de ESTADO / MODALIDAD.
"""

from __future__ import annotations

from dataclasses import dataclass
from pathlib import Path

import pandas as pd

from remasep.adapters.medinet import (
    OPTIONAL_FIELDS,
    read_medinet,
)
from remasep.services.common import Period, ValidationResult
from remasep.services.legacy_rules import LegacyRuleSet, load_legacy_rules

ANALYSIS_MODE_REAL = "REAL"

_MAX_PROBLEMS = 500
_MAX_DISTRIBUTION = 15
_MAX_CATEGORICAL_UNIQUE = 40


@dataclass(frozen=True)
class RecordProblem:
    row_number: int
    field: str
    error_code: str
    message: str


@dataclass(frozen=True)
class ColumnDiagnostic:
    field: str
    present: bool
    non_empty: int = 0
    unique_values: int = 0
    distribution: tuple[tuple[str, int], ...] = ()
    note: str = ""


@dataclass(frozen=True)
class MedinetAnalysisResult:
    analysis_mode: str
    source_name: str
    period: Period

    detected_columns: tuple[str, ...]
    missing_optional_columns: tuple[str, ...]
    sheet_name: str

    physical_rows_examined: int
    structural_empty_rows: int
    total_records: int
    valid_records: int
    invalid_records: int
    records_in_period: int
    records_outside_period: int
    min_service_date: str | None
    max_service_date: str | None

    legacy_matches_total: int
    non_target_records: int
    legacy_counts: dict[str, int]

    age_computed: int
    age_legacy_zero: int
    age_missing: int

    prestacion_empty_records: int

    validations: tuple[ValidationResult, ...]
    problems: tuple[RecordProblem, ...]
    diagnostics: tuple[ColumnDiagnostic, ...]

    legacy_rules_version: str
    legacy_rules_status: str
    legacy_category_labels: dict[str, str]
    notes: tuple[str, ...] = ()
    problems_truncated: bool = False

    def __post_init__(self) -> None:
        if self.physical_rows_examined != self.structural_empty_rows + self.total_records:
            raise ValueError(
                "MedinetAnalysisResult inconsistente: "
                f"physical_rows_examined={self.physical_rows_examined} != "
                f"structural_empty_rows={self.structural_empty_rows} + "
                f"total_records={self.total_records}"
            )
        if self.total_records != self.valid_records + self.invalid_records:
            raise ValueError(
                "MedinetAnalysisResult inconsistente: "
                f"total_records={self.total_records} != "
                f"valid_records={self.valid_records} + invalid_records={self.invalid_records}"
            )

    @property
    def period_label(self) -> str:
        return self.period.label


@dataclass
class _AgeBreakdown:
    computed: int = 0
    legacy_zero: int = 0
    missing: int = 0


def legacy_age_years(birth: object, service: object) -> int | None:
    """Años completos cumplidos a la fecha de atención (equivale a ``DATEDIF(b, s, "Y")``).

    Excepción de compatibilidad legacy: si ``birth > service`` la fórmula actual
    del workbook devuelve ``0``; aquí se reproduce ese comportamiento.
    Devuelve ``None`` si falta alguna de las dos fechas.
    """
    if birth is None or service is None or pd.isna(birth) or pd.isna(service):
        return None
    birth_ts = pd.Timestamp(birth)
    service_ts = pd.Timestamp(service)
    if birth_ts > service_ts:
        return 0
    years = service_ts.year - birth_ts.year
    if (service_ts.month, service_ts.day) < (birth_ts.month, birth_ts.day):
        years -= 1
    return int(years)


def _iso_date(value: object) -> str | None:
    if value is None or pd.isna(value):
        return None
    return pd.Timestamp(value).date().isoformat()


def _build_diagnostic(
    field_name: str,
    frame: pd.DataFrame,
    blank_masks: dict[str, pd.Series],
    *,
    categorical: bool,
) -> ColumnDiagnostic:
    if field_name not in frame.columns:
        return ColumnDiagnostic(field=field_name, present=False)

    blank = blank_masks[field_name]
    values = frame.loc[~blank, field_name].astype("string")
    unique = int(values.nunique())
    non_empty = len(values)

    distribution: tuple[tuple[str, int], ...] = ()
    note = ""
    if categorical:
        if unique <= _MAX_CATEGORICAL_UNIQUE:
            counts = values.value_counts().head(_MAX_DISTRIBUTION)
            distribution = tuple((str(k), int(v)) for k, v in counts.items())
        else:
            note = f"{unique} valores únicos: demasiados para una distribución categórica."
    else:
        note = "Solo se reporta presencia y conteo (texto libre)."

    return ColumnDiagnostic(
        field=field_name,
        present=True,
        non_empty=non_empty,
        unique_values=unique,
        distribution=distribution,
        note=note,
    )


class MedinetAnalysisService:
    """Servicio de análisis real. No contiene lógica de UI."""

    def __init__(self, legacy_rules_path: str | Path | None = None) -> None:
        self._rules: LegacyRuleSet = load_legacy_rules(legacy_rules_path)

    @property
    def legacy_rules(self) -> LegacyRuleSet:
        return self._rules

    def analyze(
        self,
        path: str | Path,
        period: Period,
        *,
        sheet_name: str | int | None = None,
    ) -> MedinetAnalysisResult:
        medinet = read_medinet(path, sheet_name=sheet_name)
        frame = medinet.frame
        masks = medinet.blank_masks
        physical_rows = len(frame)

        # --- filas estructuralmente vacías -----------------------------
        # Una fila es estructuralmente vacía solo si TODOS los campos semánticos
        # reconocidos están vacíos (fórmulas AC:AL del workbook legacy arrastradas
        # más allá de las atenciones reales). No se cuentan ni generan problemas.
        structural_empty = masks[medinet.detected_fields[0]].copy()
        for field_name in medinet.detected_fields[1:]:
            structural_empty &= masks[field_name]
        structural_empty = structural_empty.fillna(False).astype(bool)
        active = ~structural_empty
        structural_empty_rows = int(structural_empty.sum())
        total = int(active.sum())

        dia = frame["DIA_CITA"]
        fnac = frame["FECHA_NACIMIENTO"]

        dia_bad = active & dia.isna()
        fnac_blank = masks["FECHA_NACIMIENTO"]
        fnac_present_bad = active & (~fnac_blank) & fnac.isna()
        sexo_empty = active & masks["SEXO"]
        tipo_empty = active & masks["TIPO_DE_CITA"]
        prest_empty = active & masks["PRESTACION"]

        invalid_mask = dia_bad | fnac_present_bad | sexo_empty | tipo_empty
        valid_mask = active & ~invalid_mask
        birth_after = active & (~dia.isna()) & (~fnac_blank) & (~fnac.isna()) & (fnac > dia)

        # --- edad (semántica DATEDIF "Y" + excepción legacy) ---------------
        # Ver `legacy_age_years` para la definición por fila; aquí solo se
        # agregan conteos (el valor individual nunca sale del servicio).
        computable = active & (~dia.isna()) & (~fnac_blank) & (~fnac.isna())
        ages = _AgeBreakdown(
            computed=int((computable & ~birth_after).sum()),
            legacy_zero=int(birth_after.sum()),
        )
        ages.missing = total - ages.computed - ages.legacy_zero

        # --- período -----------------------------------------------------
        valid_dates = dia[valid_mask]
        in_period_mask = (valid_dates.dt.year == period.year) & (
            valid_dates.dt.month == period.month
        )
        records_in_period = int(in_period_mask.sum())
        records_outside_period = int(len(valid_dates) - records_in_period)
        min_date = _iso_date(valid_dates.min()) if len(valid_dates) else None
        max_date = _iso_date(valid_dates.max()) if len(valid_dates) else None

        # --- reglas legacy (solo registros válidos) ---------------------
        legacy_counts = {code: 0 for code in self._rules.codes}
        legacy_matches_total = 0
        non_target_records = 0
        for pos in frame.index[valid_mask]:
            codes = self._rules.classify(
                {
                    "TIPO_DE_CITA": frame.at[pos, "TIPO_DE_CITA"],
                    "PRESTACION": frame.at[pos, "PRESTACION"],
                }
            )
            if codes:
                legacy_matches_total += 1
                for code in codes:
                    legacy_counts[code] += 1
            else:
                non_target_records += 1

        # --- problemas por fila (sin valores personales) ---------------
        # Las filas estructuralmente vacías no generan problemas.
        problems: list[RecordProblem] = []
        truncated = False
        for pos in range(physical_rows):
            if not active.iat[pos]:
                continue
            row_number = pos + 2  # fila 1 = encabezados
            row_issues: list[RecordProblem] = []
            if dia_bad.iat[pos]:
                row_issues.append(
                    RecordProblem(
                        row_number, "DIA_CITA", "DIA_CITA_INVALID",
                        "DIA_CITA vacía o con formato de fecha no válido",
                    )
                )
            if fnac_present_bad.iat[pos]:
                row_issues.append(
                    RecordProblem(
                        row_number, "FECHA_NACIMIENTO", "FECHA_NACIMIENTO_INVALID",
                        "FECHA_NACIMIENTO presente pero con formato no válido",
                    )
                )
            if sexo_empty.iat[pos]:
                row_issues.append(
                    RecordProblem(row_number, "SEXO", "SEXO_EMPTY", "SEXO vacío")
                )
            if tipo_empty.iat[pos]:
                row_issues.append(
                    RecordProblem(
                        row_number, "TIPO_DE_CITA", "TIPO_DE_CITA_EMPTY", "TIPO_DE_CITA vacío"
                    )
                )
            if birth_after.iat[pos]:
                row_issues.append(
                    RecordProblem(
                        row_number, "FECHA_NACIMIENTO", "BIRTH_AFTER_SERVICE",
                        "FECHA_NACIMIENTO posterior a DIA_CITA; edad legacy = 0",
                    )
                )
            for issue in row_issues:
                if len(problems) >= _MAX_PROBLEMS:
                    truncated = True
                    break
                problems.append(issue)

        valid_records = int(valid_mask.sum())
        invalid_records = total - valid_records
        date_problem_count = int(dia_bad.sum() + fnac_present_bad.sum())

        # --- validaciones agregadas ----------------------------------
        validations = [
            ValidationResult("Archivo leído correctamente", "ok"),
            ValidationResult(
                "Estructura compatible",
                "ok",
                f"{len(medinet.detected_fields)} columnas semánticas reconocidas",
            ),
            self._period_validation(period, total, records_in_period, records_outside_period),
            (
                ValidationResult(
                    "Fechas válidas", "warning", f"{date_problem_count} fecha(s) con problemas"
                )
                if date_problem_count
                else ValidationResult("Fechas válidas", "ok")
            ),
            ValidationResult(
                "Clasificación legacy ejecutada",
                "ok",
                f"{self._rules.version} (pendiente de validación funcional)",
            ),
        ]

        # --- diagnósticos funcionales -------------------------------
        diagnostics = (
            _build_diagnostic("ESTADO", frame, masks, categorical=True),
            _build_diagnostic("MODALIDAD", frame, masks, categorical=True),
            _build_diagnostic("PRESTACION_REALIZADA", frame, masks, categorical=False),
        )

        notes = [
            *medinet.notes,
            (
                f"Las categorías legacy reproducen la lógica de {self._rules.source_workbook} "
                "y están pendientes de validación funcional."
            ),
        ]
        if truncated:
            notes.append(f"Lista de problemas truncada a los primeros {_MAX_PROBLEMS}.")
        prestacion_empty = int(prest_empty.sum())
        if structural_empty_rows:
            notes.append(
                f"{structural_empty_rows} fila(s) estructuralmente vacía(s) ignoradas "
                "(no cuentan como registros ni como inválidos)."
            )
        if prestacion_empty:
            notes.append(
                f"Diagnóstico informativo: {prestacion_empty} registro(s) con PRESTACION "
                "vacía (no es obligatoria; no genera problema)."
            )

        return MedinetAnalysisResult(
            analysis_mode=ANALYSIS_MODE_REAL,
            source_name=Path(path).name,
            period=period,
            detected_columns=tuple(medinet.detected_fields),
            missing_optional_columns=tuple(medinet.missing_optional_fields),
            sheet_name=medinet.sheet_name,
            physical_rows_examined=physical_rows,
            structural_empty_rows=structural_empty_rows,
            total_records=total,
            valid_records=valid_records,
            invalid_records=invalid_records,
            records_in_period=records_in_period,
            records_outside_period=records_outside_period,
            min_service_date=min_date,
            max_service_date=max_date,
            legacy_matches_total=legacy_matches_total,
            non_target_records=non_target_records,
            legacy_counts=legacy_counts,
            age_computed=ages.computed,
            age_legacy_zero=ages.legacy_zero,
            age_missing=ages.missing,
            prestacion_empty_records=prestacion_empty,
            validations=tuple(validations),
            problems=tuple(problems),
            diagnostics=diagnostics,
            legacy_rules_version=self._rules.version,
            legacy_rules_status=self._rules.status,
            legacy_category_labels={c.code: c.label for c in self._rules.categories},
            notes=tuple(notes),
            problems_truncated=truncated,
        )

    @staticmethod
    def _period_validation(
        period: Period, total: int, in_period: int, outside: int
    ) -> ValidationResult:
        if total and in_period == 0:
            return ValidationResult(
                "Período correcto",
                "error",
                f"Ningún registro válido corresponde a {period.label}",
            )
        if outside:
            return ValidationResult(
                "Período correcto",
                "warning",
                f"{outside} registro(s) válidos fuera de {period.label} (no se descartan)",
            )
        return ValidationResult("Período correcto", "ok", period.label)


# Campos opcionales expuestos para la UI/documentación.
DIAGNOSTIC_FIELDS = OPTIONAL_FIELDS
