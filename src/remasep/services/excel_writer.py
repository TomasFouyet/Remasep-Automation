"""Núcleo del Excel writer — contratos independientes de plataforma (Sprint 3.7B).

Este módulo **no** importa ``win32com`` ni abre Excel: define el contrato
(:class:`WorkbookWriter`), el *preflight*, la compatibilidad de plantilla, la
verificación posterior (integridad de fórmulas / VBA / celdas destino), el
parseo de la hoja ``CONTROL`` y el orquestador :func:`generate`.

La escritura real la hace un :class:`WorkbookWriter` concreto — en producción
``ExcelComWorkbookWriter`` (Windows), en tests ``FakeWorkbookWriter``. openpyxl
se usa **sólo** para inspección/verificación estática, nunca para guardar el
REMASEP oficial.
"""

from __future__ import annotations

import shutil
import zipfile
from collections.abc import Callable, Iterable, Mapping, Sequence
from contextlib import AbstractContextManager
from dataclasses import dataclass, field
from pathlib import Path
from typing import Protocol

import openpyxl

from remasep.core.errors import RemasepError
from remasep.services.metric_value_producer import (
    EXPECTED_INTEGER_COUNT,
    PendingWrite,
)
from remasep.services.workbook_delta import compute_sha256
from remasep.services.writable_target_mapping import structural_template_fingerprint

_VBA_ENTRY = "xl/vbaProject.bin"

# --- modos / estados -----------------------------------------------------
MODE_DIAGNOSTIC_REFERENCE = "DIAGNOSTIC_REFERENCE"
MODE_PRODUCTION = "PRODUCTION"
_MODES = (MODE_DIAGNOSTIC_REFERENCE, MODE_PRODUCTION)

STATUS_GENERATED_DRAFT = "GENERATED_DRAFT"
STATUS_GENERATED_DIAGNOSTIC = "GENERATED_DIAGNOSTIC_REFERENCE"
STATUS_ABORTED_PREFLIGHT = "ABORTED_PREFLIGHT"
STATUS_TEMPLATE_INCOMPATIBLE = "TEMPLATE_INCOMPATIBLE"
STATUS_EXCEL_UNAVAILABLE = "EXCEL_UNAVAILABLE"
STATUS_GENERATION_FAILED = "GENERATION_FAILED"
STATUS_FAILED_INTEGRITY_CHECK = "GENERATION_FAILED_INTEGRITY_CHECK"
STATUS_OUTPUT_EXISTS = "OUTPUT_ALREADY_EXISTS"

WRITER_INTEGRITY_PASS = "WRITER_INTEGRITY_PASS"
WRITER_INTEGRITY_FAIL = "WRITER_INTEGRITY_FAIL"

CONTROL_PASS = "PASS_INTERNAL_VALIDATION"
CONTROL_FAIL = "FAIL_INTERNAL_VALIDATION"
CONTROL_UNAVAILABLE = "CONTROL_UNAVAILABLE"

SUBMISSION_LABEL = "NOT_FOR_SUBMISSION"

TARGET_FORMULA_CONFLICT = "TARGET_FORMULA_CONFLICT"

# tipos de cambio de fórmula reutilizados conceptualmente del workbook_delta
FORMULA_EXPRESSION_CHANGE = "FORMULA_EXPRESSION_CHANGE"


class ExcelWriterError(RemasepError):
    """Error del Excel writer con mensaje claro (sin traceback al usuario)."""


def describe_com_error(exc: BaseException) -> str:
    """Resume una excepción COM a algo legible (``clase(0xHRESULT)``).

    Sirve para reportar fallos de Excel sin volcar un traceback ni depender de
    ``pythoncom`` a nivel de módulo.
    """
    hresult = getattr(exc, "hresult", None)
    if hresult is None and getattr(exc, "args", None):
        first = exc.args[0]
        if isinstance(first, int):
            hresult = first
    if isinstance(hresult, int):
        return f"{exc.__class__.__name__}(0x{hresult & 0xFFFFFFFF:08X})"
    return exc.__class__.__name__


# ---------------------------------------------------------------------------
# Contrato del writer
# ---------------------------------------------------------------------------


class WorkbookWriter(Protocol):
    """Operaciones mínimas sobre un workbook abierto. Implementaciones: COM
    (Windows) y ``FakeWorkbookWriter`` (tests). El núcleo nunca abre Excel."""

    def sheet_names(self) -> list[str]: ...

    def has_sheet(self, sheet: str) -> bool: ...

    def cell_has_formula(self, sheet: str, cell: str) -> bool: ...

    def cell_in_incompatible_merge(self, sheet: str, cell: str) -> bool: ...

    def cell_is_writable(self, sheet: str, cell: str) -> bool: ...

    def read_cell(self, sheet: str, cell: str) -> object: ...

    def write_value2(self, sheet: str, cell: str, value: object) -> None: ...

    def recalculate(self) -> None: ...

    def save(self) -> None: ...

    def snapshot(self) -> WorkbookSnapshot: ...


# ---------------------------------------------------------------------------
# Snapshot estático (openpyxl) para verificación
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class WorkbookSnapshot:
    path: str
    sheet_names: tuple[str, ...]
    formula_map: Mapping[tuple[str, str], str]
    values: Mapping[tuple[str, str], object]
    vba_present: bool
    vba_payload_sha256: str | None
    structural_fingerprint_id: str
    file_sha256: str | None = None

    def formula_count(self) -> int:
        return len(self.formula_map)


def vba_payload_info(path: str | Path) -> tuple[bool, str | None]:
    """``(present, payload_sha256)`` del proyecto VBA embebido, de forma estable."""
    p = Path(path)
    if not zipfile.is_zipfile(p):
        return False, None
    with zipfile.ZipFile(p) as archive:
        if _VBA_ENTRY not in archive.namelist():
            return False, None
        import hashlib

        return True, hashlib.sha256(archive.read(_VBA_ENTRY)).hexdigest()


def snapshot_from_path(path: str | Path, *, with_values: bool = True) -> WorkbookSnapshot:
    """Construye un :class:`WorkbookSnapshot` por inspección estática (openpyxl)."""
    p = Path(path)
    keep_vba = p.suffix.lower() in {".xlsm", ".xltm", ".xlsb"}
    wb_f = openpyxl.load_workbook(p, data_only=False, read_only=False, keep_vba=keep_vba)
    formula_map: dict[tuple[str, str], str] = {}
    for ws in wb_f.worksheets:
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == "f":
                    formula_map[(ws.title, cell.coordinate)] = str(cell.value)
    fingerprint = structural_template_fingerprint(wb_f)["structural_template_fingerprint_id"]
    sheet_names = tuple(wb_f.sheetnames)
    wb_f.close()

    values: dict[tuple[str, str], object] = {}
    if with_values:
        wb_v = openpyxl.load_workbook(p, data_only=True, read_only=True)
        for ws in wb_v.worksheets:
            for row in ws.iter_rows():
                for cell in row:
                    if cell.value is not None:
                        values[(ws.title, cell.coordinate)] = cell.value
        wb_v.close()

    present, payload_sha = vba_payload_info(p)
    return WorkbookSnapshot(
        path=str(p),
        sheet_names=sheet_names,
        formula_map=formula_map,
        values=values,
        vba_present=present,
        vba_payload_sha256=payload_sha,
        structural_fingerprint_id=fingerprint,
        file_sha256=compute_sha256(p),
    )


# ---------------------------------------------------------------------------
# Preflight (§4)
# ---------------------------------------------------------------------------


@dataclass
class PreflightReport:
    ok: bool
    errors: list[str] = field(default_factory=list)
    checks: dict[str, bool] = field(default_factory=dict)

    def fail(self, code: str) -> None:
        self.ok = False
        self.errors.append(code)


def _is_within(path: Path, ancestor: Path) -> bool:
    try:
        path.resolve().relative_to(ancestor.resolve())
        return True
    except ValueError:
        return False


def run_preflight(
    *,
    template_path: Path,
    output_path: Path,
    pending_writes: Sequence[PendingWrite],
    manifest_instruction_ids: Sequence[str],
    zero_write_policy: str,
    zero_write_policy_accepted: Sequence[str],
    forbidden_dirs: Sequence[Path] = (),
) -> PreflightReport:
    report = PreflightReport(ok=True)

    def check(name: str, passed: bool, code: str) -> None:
        report.checks[name] = passed
        if not passed:
            report.fail(code)

    check("template_exists", template_path.is_file(), "TEMPLATE_NOT_FOUND")
    check(
        "template_is_xlsm",
        template_path.suffix.lower() == ".xlsm",
        "TEMPLATE_NOT_XLSM",
    )
    same = template_path.resolve() == output_path.resolve()
    check("output_is_not_template", not same, "OUTPUT_EQUALS_TEMPLATE")
    check("output_is_xlsm", output_path.suffix.lower() == ".xlsm", "OUTPUT_NOT_XLSM")
    check(
        "output_not_in_forbidden_dir",
        not any(_is_within(output_path, d) for d in forbidden_dirs),
        "OUTPUT_IN_FORBIDDEN_DIR",
    )
    check(
        "output_does_not_exist",
        not output_path.exists(),
        "OUTPUT_ALREADY_EXISTS",
    )

    ids = [pw.instruction_id for pw in pending_writes]
    check("instruction_ids_unique", len(ids) == len(set(ids)), "DUPLICATE_INSTRUCTION_ID")
    targets = [(pw.target_sheet, pw.target_cell) for pw in pending_writes]
    check("target_cells_unique", len(targets) == len(set(targets)), "DUPLICATE_TARGET_CELL")

    type_ok = True
    values_ok = True
    for pw in pending_writes:
        if pw.expected_value_type == EXPECTED_INTEGER_COUNT:
            if isinstance(pw.value, bool) or not isinstance(pw.value, int):
                type_ok = False
            elif pw.value < 0:
                values_ok = False
        else:  # sólo INTEGER_COUNT está soportado por el writer en 3.7B
            type_ok = False
    check("value_types_ok", type_ok, "VALUE_TYPE_MISMATCH")
    check("values_non_negative", values_ok, "NEGATIVE_COUNT")
    check("no_missing_values", all(pw.value is not None for pw in pending_writes), "MISSING_VALUE")

    manifest_set = set(manifest_instruction_ids)
    covered = set(ids)
    check(
        "no_missing_instructions",
        manifest_set.issubset(covered),
        "MISSING_WRITE_INSTRUCTIONS",
    )
    check(
        "no_orphan_instructions",
        covered.issubset(manifest_set) if manifest_set else True,
        "ORPHAN_WRITE_INSTRUCTIONS",
    )

    check(
        "zero_write_policy_compatible",
        zero_write_policy in set(zero_write_policy_accepted),
        "ZERO_WRITE_POLICY_INCOMPATIBLE",
    )

    return report


# ---------------------------------------------------------------------------
# Compatibilidad de plantilla (§5)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class TemplateCompatibility:
    compatible: bool
    expected_fingerprint_id: str
    actual_fingerprint_id: str
    file_sha256: str


def check_template_compatibility(
    template_path: Path, expected_fingerprint_id: str
) -> TemplateCompatibility:
    snapshot = snapshot_from_path(template_path, with_values=False)
    actual = snapshot.structural_fingerprint_id
    return TemplateCompatibility(
        compatible=actual == expected_fingerprint_id,
        expected_fingerprint_id=expected_fingerprint_id,
        actual_fingerprint_id=actual,
        file_sha256=snapshot.file_sha256 or "",
    )


# ---------------------------------------------------------------------------
# Integridad de fórmulas / VBA (§18 / §19)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class FormulaIntegrityReport:
    formula_count_before: int
    formula_count_after: int
    expression_changes: tuple[tuple[str, str], ...]
    added: tuple[tuple[str, str], ...]
    removed: tuple[tuple[str, str], ...]

    @property
    def ok(self) -> bool:
        return not (self.expression_changes or self.added or self.removed)


def compare_formula_integrity(
    before: WorkbookSnapshot, after: WorkbookSnapshot
) -> FormulaIntegrityReport:
    before_keys = set(before.formula_map)
    after_keys = set(after.formula_map)
    changed = tuple(
        sorted(
            key
            for key in before_keys & after_keys
            if _norm_formula(before.formula_map[key]) != _norm_formula(after.formula_map[key])
        )
    )
    return FormulaIntegrityReport(
        formula_count_before=len(before_keys),
        formula_count_after=len(after_keys),
        expression_changes=changed,
        added=tuple(sorted(after_keys - before_keys)),
        removed=tuple(sorted(before_keys - after_keys)),
    )


def _norm_formula(text: str) -> str:
    return text.strip().removeprefix("=").replace(" ", "").casefold()


@dataclass(frozen=True)
class VbaIntegrityReport:
    present_before: bool
    present_after: bool
    payload_sha256_before: str | None
    payload_sha256_after: str | None

    @property
    def ok(self) -> bool:
        if not (self.present_before and self.present_after):
            return self.present_before == self.present_after
        if self.payload_sha256_before and self.payload_sha256_after:
            return self.payload_sha256_before == self.payload_sha256_after
        return True

    @property
    def payload_stable(self) -> bool | None:
        if self.payload_sha256_before and self.payload_sha256_after:
            return self.payload_sha256_before == self.payload_sha256_after
        return None


def compare_vba_integrity(
    before: WorkbookSnapshot, after: WorkbookSnapshot
) -> VbaIntegrityReport:
    return VbaIntegrityReport(
        present_before=before.vba_present,
        present_after=after.vba_present,
        payload_sha256_before=before.vba_payload_sha256,
        payload_sha256_after=after.vba_payload_sha256,
    )


# ---------------------------------------------------------------------------
# Verificación de celdas destino (§17)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class TargetCheck:
    instruction_id: str
    source_metric_id: str
    target_sheet: str
    target_cell: str
    expected_value: object
    actual_value: object
    status: str  # OK | VALUE_MISMATCH | MISSING | BECAME_FORMULA


@dataclass(frozen=True)
class TargetVerification:
    checks: tuple[TargetCheck, ...]

    @property
    def written_ok(self) -> int:
        return sum(1 for c in self.checks if c.status == "OK")

    @property
    def failures(self) -> tuple[TargetCheck, ...]:
        return tuple(c for c in self.checks if c.status != "OK")

    @property
    def ok(self) -> bool:
        return not self.failures


def verify_targets(
    after: WorkbookSnapshot, pending_writes: Sequence[PendingWrite]
) -> TargetVerification:
    checks: list[TargetCheck] = []
    for pw in pending_writes:
        key = (pw.target_sheet, pw.target_cell)
        if key in after.formula_map:
            status, actual = "BECAME_FORMULA", after.formula_map[key]
        elif key not in after.values:
            # 0 puede no aparecer en `values` si openpyxl lo omite; se maneja aparte
            actual = after.values.get(key, 0 if _is_zero(pw.value) else None)
            status = "OK" if _values_equal(actual, pw.value) else "MISSING"
        else:
            actual = after.values[key]
            status = "OK" if _values_equal(actual, pw.value) else "VALUE_MISMATCH"
        checks.append(
            TargetCheck(
                instruction_id=pw.instruction_id,
                source_metric_id=pw.source_metric_id,
                target_sheet=pw.target_sheet,
                target_cell=pw.target_cell,
                expected_value=pw.value,
                actual_value=actual,
                status=status,
            )
        )
    return TargetVerification(checks=tuple(checks))


def _is_zero(value: object) -> bool:
    return isinstance(value, (int, float)) and not isinstance(value, bool) and value == 0


def _values_equal(a: object, b: object) -> bool:
    if isinstance(a, bool) or isinstance(b, bool):
        return a is b
    if isinstance(a, (int, float)) and isinstance(b, (int, float)):
        return float(a) == float(b)
    return a == b


# ---------------------------------------------------------------------------
# CONTROL (§21 / §22)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class ControlRowMap:
    sheet: str
    label_cell: str
    errors_cell: str
    data_status_cell: str | None
    causal_cell: str | None


@dataclass(frozen=True)
class ControlMap:
    sheet: str
    total_errors_cell: str
    rows: tuple[ControlRowMap, ...]
    data_status_ok_value: str
    data_status_missing_value: str
    modules_not_yet_produced: tuple[str, ...]


def load_control_map(raw: Mapping) -> ControlMap:
    rows = tuple(
        ControlRowMap(
            sheet=str(r["sheet"]),
            label_cell=str(r["label_cell"]),
            errors_cell=str(r["errors_cell"]),
            data_status_cell=(str(r["data_status_cell"]) if r.get("data_status_cell") else None),
            causal_cell=(str(r["causal_cell"]) if r.get("causal_cell") else None),
        )
        for r in raw.get("rows", [])
    )
    return ControlMap(
        sheet=str(raw.get("sheet", "CONTROL")),
        total_errors_cell=str(raw.get("total_errors_cell", "E16")),
        rows=rows,
        data_status_ok_value=str(raw.get("data_status_ok_value", "OK")),
        data_status_missing_value=str(raw.get("data_status_missing_value", "SIN DATOS")),
        modules_not_yet_produced=tuple(raw.get("modules_not_yet_produced", [])),
    )


@dataclass(frozen=True)
class ControlSheetRow:
    sheet: str
    errors: object
    data_status: object
    causal: object


@dataclass(frozen=True)
class ControlResult:
    status: str
    total_errors: object
    rows: tuple[ControlSheetRow, ...]
    errors_from_pending_modules: int
    reason: str = ""

    def as_dict(self) -> dict:
        return {
            "status": self.status,
            "total_errors": self.total_errors,
            "errors_from_pending_modules": self.errors_from_pending_modules,
            "reason": self.reason,
            "rows": [
                {
                    "sheet": r.sheet,
                    "errors": r.errors,
                    "data_status": r.data_status,
                    "causal": r.causal,
                }
                for r in self.rows
            ],
        }


def parse_control(after: WorkbookSnapshot, control_map: ControlMap) -> ControlResult:
    if control_map.sheet not in after.sheet_names:
        return ControlResult(
            status=CONTROL_UNAVAILABLE,
            total_errors=None,
            rows=(),
            errors_from_pending_modules=0,
            reason="CONTROL sheet ausente",
        )

    def value_at(cell: str | None) -> object:
        if cell is None:
            return None
        return after.values.get((control_map.sheet, cell))

    rows: list[ControlSheetRow] = []
    pending_module_errors = 0
    for row_map in control_map.rows:
        errors = value_at(row_map.errors_cell)
        rows.append(
            ControlSheetRow(
                sheet=row_map.sheet,
                errors=errors,
                data_status=value_at(row_map.data_status_cell),
                causal=value_at(row_map.causal_cell),
            )
        )
        if row_map.sheet in control_map.modules_not_yet_produced and isinstance(
            errors, (int, float)
        ):
            pending_module_errors += int(errors)

    total = value_at(control_map.total_errors_cell)
    if total is None:
        status = CONTROL_UNAVAILABLE
        reason = f"celda de total {control_map.total_errors_cell} vacía (workbook sin recalcular?)"
    elif isinstance(total, (int, float)) and total == 0:
        status = CONTROL_PASS
        reason = ""
    else:
        status = CONTROL_FAIL
        reason = f"CONTROL reporta {total} error(es); parte proviene de módulos aún no producidos"
    return ControlResult(
        status=status,
        total_errors=total,
        rows=tuple(rows),
        errors_from_pending_modules=pending_module_errors,
        reason=reason,
    )


# ---------------------------------------------------------------------------
# Nombres de salida (§16)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class OutputPlan:
    final_path: Path
    working_path: Path


def plan_output_paths(
    output_dir: Path,
    *,
    year: int,
    month: int,
    draft: bool,
    working_suffix: str,
    draft_name_template: str,
    final_name_template: str,
    explicit_output: Path | None = None,
) -> OutputPlan:
    if explicit_output is not None:
        final_path = explicit_output
    else:
        template = draft_name_template if draft else final_name_template
        final_path = output_dir / template.format(year=year, month=month)
    working_path = final_path.with_name(
        final_path.stem + working_suffix + final_path.suffix
    )
    return OutputPlan(final_path=final_path, working_path=working_path)


# ---------------------------------------------------------------------------
# Orquestador (§6 / §11 / §12 / §17 / §18)
# ---------------------------------------------------------------------------


@dataclass(frozen=True)
class GenerationRequest:
    mode: str
    template_path: Path
    output_path: Path
    period_year: int
    period_month: int
    expected_fingerprint_id: str
    zero_write_policy: str
    zero_write_policy_accepted: tuple[str, ...]
    manifest_instruction_ids: tuple[str, ...]
    forbidden_dirs: tuple[Path, ...] = ()
    historical_reference_cache_status: str = "KNOWN_STALE_FOR_142_WRITE_READY_VALUES"


@dataclass
class GenerationResult:
    status: str
    mode: str
    submission_label: str = SUBMISSION_LABEL
    output_path: str | None = None
    working_path: str | None = None
    written_cells: int = 0
    writer_integrity_status: str = ""
    control_status: str = ""
    preflight: PreflightReport | None = None
    template_compatibility: TemplateCompatibility | None = None
    formula_integrity: FormulaIntegrityReport | None = None
    vba_integrity: VbaIntegrityReport | None = None
    target_verification: TargetVerification | None = None
    control_result: ControlResult | None = None
    template_sha256_before: str | None = None
    template_sha256_after: str | None = None
    template_unchanged: bool | None = None
    historical_reference_cache_status: str = ""
    warnings: list[str] = field(default_factory=list)
    errors: list[str] = field(default_factory=list)

    @property
    def succeeded(self) -> bool:
        return self.status in (STATUS_GENERATED_DRAFT, STATUS_GENERATED_DIAGNOSTIC)


def generate(
    request: GenerationRequest,
    pending_writes: Sequence[PendingWrite],
    *,
    open_writer: Callable[[Path], AbstractContextManager[WorkbookWriter]],
    control_map: ControlMap,
    inspect: Callable[[Path], WorkbookSnapshot] = snapshot_from_path,
) -> GenerationResult:
    """Copia la plantilla, escribe los ``PendingWrite`` en la copia, recalcula,
    verifica y —sólo si todo pasa— promueve la copia de trabajo a salida final.

    ``open_writer`` construye el :class:`WorkbookWriter` (COM real o fake).
    ``inspect`` lee un :class:`WorkbookSnapshot` de un archivo (openpyxl por
    defecto); se inyecta en tests.
    """
    if request.mode not in _MODES:
        raise ExcelWriterError(f"modo de generación no soportado: {request.mode!r}")

    result = GenerationResult(
        status=STATUS_GENERATION_FAILED,
        mode=request.mode,
        historical_reference_cache_status=request.historical_reference_cache_status,
    )

    # --- preflight (antes de tocar nada) -------------------------------
    preflight = run_preflight(
        template_path=request.template_path,
        output_path=request.output_path,
        pending_writes=pending_writes,
        manifest_instruction_ids=request.manifest_instruction_ids,
        zero_write_policy=request.zero_write_policy,
        zero_write_policy_accepted=request.zero_write_policy_accepted,
        forbidden_dirs=request.forbidden_dirs,
    )
    result.preflight = preflight
    if not preflight.ok:
        result.status = (
            STATUS_OUTPUT_EXISTS
            if preflight.errors == ["OUTPUT_ALREADY_EXISTS"]
            else STATUS_ABORTED_PREFLIGHT
        )
        result.errors = list(preflight.errors)
        return result

    # --- snapshot BEFORE + sha de la plantilla -----------------------
    result.template_sha256_before = compute_sha256(request.template_path)
    before = inspect(request.template_path)

    # --- compatibilidad de plantilla (fingerprint estructural) ------
    compat = TemplateCompatibility(
        compatible=before.structural_fingerprint_id == request.expected_fingerprint_id,
        expected_fingerprint_id=request.expected_fingerprint_id,
        actual_fingerprint_id=before.structural_fingerprint_id,
        file_sha256=before.file_sha256 or result.template_sha256_before or "",
    )
    result.template_compatibility = compat
    if not compat.compatible:
        result.status = STATUS_TEMPLATE_INCOMPATIBLE
        result.errors = [
            (
                f"TEMPLATE_INCOMPATIBLE esperado={compat.expected_fingerprint_id} "
                f"actual={compat.actual_fingerprint_id}"
            )
        ]
        return result

    working_path = request.output_path.with_name(
        request.output_path.stem + ".__working__" + request.output_path.suffix
    )
    result.working_path = str(working_path)
    working_path.parent.mkdir(parents=True, exist_ok=True)
    _safe_unlink(working_path)

    # --- copy-first + escritura --------------------------------------
    try:
        shutil.copy2(request.template_path, working_path)
        with open_writer(working_path) as writer:
            _assert_targets_writable(writer, pending_writes)
            for pw in pending_writes:
                writer.write_value2(pw.target_sheet, pw.target_cell, pw.value)
            writer.recalculate()
            writer.save()
        after = inspect(working_path)
    except ExcelWriterError as exc:
        _safe_unlink(working_path)
        result.status = (
            STATUS_FAILED_INTEGRITY_CHECK
            if TARGET_FORMULA_CONFLICT in str(exc)
            else STATUS_GENERATION_FAILED
        )
        result.errors = [str(exc)]
        result.template_sha256_after = compute_sha256(request.template_path)
        result.template_unchanged = (
            result.template_sha256_after == result.template_sha256_before
        )
        return result
    except Exception as exc:  # noqa: BLE001 - se resume a un motivo legible
        _safe_unlink(working_path)
        result.status = STATUS_GENERATION_FAILED
        result.errors = [f"{exc.__class__.__name__}: {exc}"]
        result.template_sha256_after = compute_sha256(request.template_path)
        result.template_unchanged = (
            result.template_sha256_after == result.template_sha256_before
        )
        return result

    # --- verificación posterior ------------------------------------
    formula_integrity = compare_formula_integrity(before, after)
    vba_integrity = compare_vba_integrity(before, after)
    target_verification = verify_targets(after, pending_writes)
    control_result = parse_control(after, control_map)

    result.formula_integrity = formula_integrity
    result.vba_integrity = vba_integrity
    result.target_verification = target_verification
    result.control_result = control_result
    result.control_status = control_result.status
    result.written_cells = target_verification.written_ok

    fingerprint_still_ok = after.structural_fingerprint_id == request.expected_fingerprint_id
    sheets_ok = set(before.sheet_names).issubset(set(after.sheet_names))

    integrity_ok = (
        formula_integrity.ok
        and vba_integrity.ok
        and target_verification.ok
        and fingerprint_still_ok
        and sheets_ok
    )
    result.writer_integrity_status = (
        WRITER_INTEGRITY_PASS if integrity_ok else WRITER_INTEGRITY_FAIL
    )

    if not integrity_ok:
        if not formula_integrity.ok:
            result.errors.append("GENERATION_FAILED_INTEGRITY_CHECK: fórmulas alteradas")
        if not vba_integrity.ok:
            result.errors.append("GENERATION_FAILED_INTEGRITY_CHECK: VBA alterado/ausente")
        if not target_verification.ok:
            result.errors.append(
                "GENERATION_FAILED_INTEGRITY_CHECK: "
                f"{len(target_verification.failures)} celda(s) destino no verificadas"
            )
        if not fingerprint_still_ok:
            result.errors.append("GENERATION_FAILED_INTEGRITY_CHECK: fingerprint estructural cambió")
        if not sheets_ok:
            result.errors.append("GENERATION_FAILED_INTEGRITY_CHECK: faltan hojas")
        _safe_unlink(working_path)
        result.status = STATUS_FAILED_INTEGRITY_CHECK
        result.template_sha256_after = compute_sha256(request.template_path)
        result.template_unchanged = (
            result.template_sha256_after == result.template_sha256_before
        )
        return result

    # --- promoción atómica a salida final --------------------------
    if request.output_path.exists():
        _safe_unlink(working_path)
        result.status = STATUS_OUTPUT_EXISTS
        result.errors = ["OUTPUT_ALREADY_EXISTS"]
        return result
    working_path.replace(request.output_path)
    result.output_path = str(request.output_path)
    result.working_path = None

    result.template_sha256_after = compute_sha256(request.template_path)
    result.template_unchanged = result.template_sha256_after == result.template_sha256_before
    if not result.template_unchanged:
        result.warnings.append("TEMPLATE_SHA256_CHANGED")

    if control_result.status == CONTROL_FAIL:
        result.warnings.append(
            "CONTROL reporta errores; esperable en la primera generación sólo-MEDINET "
            "(faltan EGRESOS / recursos / tabla quirúrgica / metadatos). "
            "writer_integrity_status es independiente de control_status."
        )

    result.status = (
        STATUS_GENERATED_DIAGNOSTIC
        if request.mode == MODE_DIAGNOSTIC_REFERENCE
        else STATUS_GENERATED_DRAFT
    )
    return result


def _assert_targets_writable(
    writer: WorkbookWriter, pending_writes: Sequence[PendingWrite]
) -> None:
    """Defense-in-depth por celda antes de escribir (§9)."""
    for pw in pending_writes:
        if not writer.has_sheet(pw.target_sheet):
            raise ExcelWriterError(f"hoja destino ausente: {pw.target_sheet!r}")
        if writer.cell_has_formula(pw.target_sheet, pw.target_cell):
            raise ExcelWriterError(
                f"{TARGET_FORMULA_CONFLICT}: {pw.target_sheet}!{pw.target_cell} contiene fórmula"
            )
        if writer.cell_in_incompatible_merge(pw.target_sheet, pw.target_cell):
            raise ExcelWriterError(
                f"celda destino en merge incompatible: {pw.target_sheet}!{pw.target_cell}"
            )
        if not writer.cell_is_writable(pw.target_sheet, pw.target_cell):
            raise ExcelWriterError(
                f"celda destino no escribible (protegida): {pw.target_sheet}!{pw.target_cell}"
            )


def _safe_unlink(path: Path) -> None:
    with_suffix = path.suffix.lower() in {".xlsm", ".xlsx", ".xltm"}
    if path.exists() and with_suffix and "__working__" in path.name:
        path.unlink()


# ---------------------------------------------------------------------------
# Audit sin PII (§23)
# ---------------------------------------------------------------------------


def build_write_audit(
    pending_writes: Sequence[PendingWrite], verification: TargetVerification | None
) -> list[dict]:
    status_by_id = (
        {c.instruction_id: c.status for c in verification.checks} if verification else {}
    )
    return [
        {
            "instruction_id": pw.instruction_id,
            "source_metric_id": pw.source_metric_id,
            "target_sheet": pw.target_sheet,
            "target_cell": pw.target_cell,
            "written_value": pw.value,
            "write_status": status_by_id.get(pw.instruction_id, "NOT_VERIFIED"),
        }
        for pw in pending_writes
    ]


def iter_forbidden_dirs(*names: str) -> Iterable[Path]:
    for name in names:
        yield Path(name)
