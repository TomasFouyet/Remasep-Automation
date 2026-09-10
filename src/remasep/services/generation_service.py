"""Servicio de generación del REMASEP (Sprint 3.7B).

Orquesta: capacidad de plataforma → política/manifest → :func:`excel_writer.generate`
→ artefactos de verificación. Es la API que consumirá el botón de la UI en el
sprint siguiente (aún **no** conectado).

No importa ``win32com`` salvo, de forma perezosa, cuando efectivamente hay que
generar en Windows.
"""

from __future__ import annotations

import csv
import json
from collections.abc import Callable, Sequence
from contextlib import AbstractContextManager
from dataclasses import dataclass, field
from datetime import UTC, datetime
from pathlib import Path

import yaml

from remasep.core.errors import RemasepError
from remasep.services.excel_capability import ExcelCapability, detect_excel_capability
from remasep.services.excel_writer import (
    MODE_DIAGNOSTIC_REFERENCE,
    MODE_PRODUCTION,
    STATUS_EXCEL_UNAVAILABLE,
    SUBMISSION_LABEL,
    ControlMap,
    GenerationRequest,
    GenerationResult,
    WorkbookSnapshot,
    WorkbookWriter,
    build_write_audit,
    load_control_map,
    plan_output_paths,
    snapshot_from_path,
)
from remasep.services.metric_value_producer import PendingWrite

_DEFAULT_CONFIG_DIR = Path("config/excel_writer_2026")
_DEFAULT_ARTIFACTS_DIR = Path("artifacts/excel_writer")
_FALLBACK_FINGERPRINT = "stf:dc624775927d4d4d"


class GenerationServiceError(RemasepError):
    pass


@dataclass
class GenerationServiceResult:
    result: GenerationResult
    capability: ExcelCapability
    mode: str
    submission_label: str
    artifacts_dir: str | None = None
    artifact_files: list[str] = field(default_factory=list)

    # atajos que la UI espera (§31)
    @property
    def status(self) -> str:
        return self.result.status

    @property
    def output_path(self) -> str | None:
        return self.result.output_path

    @property
    def written_cells(self) -> int:
        return self.result.written_cells

    @property
    def control_status(self) -> str:
        return self.result.control_status

    @property
    def warnings(self) -> list[str]:
        return self.result.warnings

    @property
    def errors(self) -> list[str]:
        return self.result.errors


class GenerationService:
    def __init__(
        self,
        *,
        config_dir: Path = _DEFAULT_CONFIG_DIR,
        artifacts_dir: Path = _DEFAULT_ARTIFACTS_DIR,
    ) -> None:
        self._config_dir = Path(config_dir)
        self._artifacts_dir = Path(artifacts_dir)

    # -- configuración ------------------------------------------------
    def _policy(self) -> dict:
        path = self._config_dir / "policy.yaml"
        if not path.is_file():
            raise GenerationServiceError(f"falta la política del writer: {path}")
        return yaml.safe_load(path.read_text(encoding="utf-8")) or {}

    def _control_map(self, policy: dict) -> ControlMap:
        rel = policy.get("control_map") or str(self._config_dir / "control_map.yaml")
        path = Path(rel)
        if not path.is_file():
            raise GenerationServiceError(f"falta el mapa de CONTROL: {path}")
        return load_control_map(yaml.safe_load(path.read_text(encoding="utf-8")) or {})

    def capability(self, *, probe_com: bool = True) -> ExcelCapability:
        return detect_excel_capability(probe_com=probe_com)

    # -- generación -------------------------------------------------
    def generate(
        self,
        *,
        mode: str,
        template_path: str | Path,
        pending_writes: Sequence[PendingWrite],
        manifest_instruction_ids: Sequence[str],
        period_year: int,
        period_month: int,
        zero_write_policy: str,
        output_path: str | Path | None = None,
        expected_fingerprint_id: str | None = None,
        open_writer: Callable[[Path], AbstractContextManager[WorkbookWriter]] | None = None,
        inspect: Callable[[Path], WorkbookSnapshot] = snapshot_from_path,
        write_artifacts: bool = True,
        probe_com: bool = True,
    ) -> GenerationServiceResult:
        if mode not in (MODE_DIAGNOSTIC_REFERENCE, MODE_PRODUCTION):
            raise GenerationServiceError(f"modo no soportado: {mode!r}")

        policy = self._policy()
        control_map = self._control_map(policy)
        submission_label = (
            (policy.get("modes", {}).get(mode, {}) or {}).get("submission_label")
            or SUBMISSION_LABEL
        )
        accepted = tuple(policy.get("zero_write_policy_accepted") or ["WRITE_ZERO"])
        expected_fp = (
            expected_fingerprint_id
            or (policy.get("template_compatibility") or {}).get(
                "expected_fingerprint_id_fallback"
            )
            or _FALLBACK_FINGERPRINT
        )
        forbidden = tuple(Path(d) for d in ("data/local", "data", "inputs"))

        # --- resolver el open_writer real si no se inyectó uno ---------
        capability = self.capability(probe_com=probe_com)
        if open_writer is None:
            if not capability.can_generate:
                result = GenerationResult(
                    status=STATUS_EXCEL_UNAVAILABLE,
                    mode=mode,
                    submission_label=submission_label,
                    errors=[capability.reason],
                    warnings=[capability.user_message],
                )
                return GenerationServiceResult(
                    result=result,
                    capability=capability,
                    mode=mode,
                    submission_label=submission_label,
                )
            from remasep.adapters.excel_com import open_excel_com_writer  # lazy, Windows

            open_writer = open_excel_com_writer

        # --- ruta de salida -----------------------------------------
        out_cfg = policy.get("output") or {}
        if output_path is not None:
            final_path = Path(output_path)
        else:
            plan = plan_output_paths(
                Path(out_cfg.get("directory", "outputs")),
                year=period_year,
                month=period_month,
                draft=True,
                working_suffix=out_cfg.get("working_suffix", ".__working__"),
                draft_name_template=out_cfg.get(
                    "draft_name_template", "REMASEP_{year}_{month:02d}_DRAFT.xlsm"
                ),
                final_name_template=out_cfg.get(
                    "final_name_template", "REMASEP_{year}_{month:02d}.xlsm"
                ),
            )
            final_path = plan.final_path

        request = GenerationRequest(
            mode=mode,
            template_path=Path(template_path),
            output_path=final_path,
            period_year=period_year,
            period_month=period_month,
            expected_fingerprint_id=expected_fp,
            zero_write_policy=zero_write_policy,
            zero_write_policy_accepted=accepted,
            manifest_instruction_ids=tuple(manifest_instruction_ids),
            forbidden_dirs=forbidden,
            historical_reference_cache_status=str(
                policy.get(
                    "historical_reference_cache_status",
                    "KNOWN_STALE_FOR_142_WRITE_READY_VALUES",
                )
            ),
        )

        from remasep.services.excel_writer import generate as _generate

        result = _generate(
            request,
            list(pending_writes),
            open_writer=open_writer,
            control_map=control_map,
            inspect=inspect,
        )
        result.submission_label = submission_label

        service_result = GenerationServiceResult(
            result=result,
            capability=capability,
            mode=mode,
            submission_label=submission_label,
        )
        if write_artifacts:
            files = self._write_artifacts(service_result, pending_writes)
            service_result.artifacts_dir = str(self._artifacts_dir)
            service_result.artifact_files = files
        return service_result

    # -- artefactos (§24) ----------------------------------------
    def _write_artifacts(
        self, sr: GenerationServiceResult, pending_writes: Sequence[PendingWrite]
    ) -> list[str]:
        self._artifacts_dir.mkdir(parents=True, exist_ok=True)
        result = sr.result
        files: list[str] = []

        def write_csv(name: str, fieldnames: list[str], rows: list[dict]) -> None:
            with (self._artifacts_dir / name).open("w", encoding="utf-8", newline="") as fh:
                writer = csv.DictWriter(fh, fieldnames=fieldnames)
                writer.writeheader()
                writer.writerows(rows)
            files.append(name)

        audit = build_write_audit(pending_writes, result.target_verification)
        write_csv(
            "write_audit.csv",
            ["instruction_id", "source_metric_id", "target_sheet", "target_cell",
             "written_value", "write_status"],
            audit,
        )

        tv_rows = (
            [
                {
                    "instruction_id": c.instruction_id,
                    "source_metric_id": c.source_metric_id,
                    "target_sheet": c.target_sheet,
                    "target_cell": c.target_cell,
                    "expected_value": c.expected_value,
                    "actual_value": c.actual_value,
                    "status": c.status,
                }
                for c in result.target_verification.checks
            ]
            if result.target_verification
            else []
        )
        write_csv(
            "target_verification.csv",
            ["instruction_id", "source_metric_id", "target_sheet", "target_cell",
             "expected_value", "actual_value", "status"],
            tv_rows,
        )

        fi = result.formula_integrity
        fi_rows = []
        if fi:
            for sheet, cell in fi.expression_changes:
                fi_rows.append({"sheet": sheet, "cell": cell, "change": "EXPRESSION_CHANGE"})
            for sheet, cell in fi.added:
                fi_rows.append({"sheet": sheet, "cell": cell, "change": "ADDED"})
            for sheet, cell in fi.removed:
                fi_rows.append({"sheet": sheet, "cell": cell, "change": "REMOVED"})
        write_csv("formula_integrity.csv", ["sheet", "cell", "change"], fi_rows)

        cr = result.control_result
        cr_rows = (
            [
                {
                    "sheet": r.sheet,
                    "errors": r.errors,
                    "data_status": r.data_status,
                    "causal": r.causal,
                }
                for r in cr.rows
            ]
            if cr
            else []
        )
        write_csv("control_result.csv", ["sheet", "errors", "data_status", "causal"], cr_rows)

        template_verification = {
            "template_compatibility": _compat_dict(result),
            "template_sha256_before": result.template_sha256_before,
            "template_sha256_after": result.template_sha256_after,
            "template_unchanged": result.template_unchanged,
            "formula_count_before": fi.formula_count_before if fi else None,
            "formula_count_after": fi.formula_count_after if fi else None,
            "formula_integrity_ok": fi.ok if fi else None,
            "vba_present_before": result.vba_integrity.present_before if result.vba_integrity else None,
            "vba_present_after": result.vba_integrity.present_after if result.vba_integrity else None,
            "vba_payload_stable": (
                result.vba_integrity.payload_stable if result.vba_integrity else None
            ),
        }
        (self._artifacts_dir / "template_verification.json").write_text(
            json.dumps(template_verification, indent=2, ensure_ascii=False, default=str) + "\n",
            encoding="utf-8",
        )
        files.append("template_verification.json")

        summary = {
            "generated_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
            "mode": sr.mode,
            "submission_label": sr.submission_label,
            "status": result.status,
            "succeeded": result.succeeded,
            "output_path": result.output_path,
            "written_cells": result.written_cells,
            "writer_integrity_status": result.writer_integrity_status,
            "control_status": result.control_status,
            "control_total_errors": cr.total_errors if cr else None,
            "control_errors_from_pending_modules": (
                cr.errors_from_pending_modules if cr else None
            ),
            "historical_reference_cache_status": result.historical_reference_cache_status,
            "template_unchanged": result.template_unchanged,
            "pending_write_count": len(pending_writes),
            "preflight_ok": result.preflight.ok if result.preflight else None,
            "preflight_errors": result.preflight.errors if result.preflight else [],
            "template_compatible": (
                result.template_compatibility.compatible
                if result.template_compatibility
                else None
            ),
            "target_verification_ok": (
                result.target_verification.ok if result.target_verification else None
            ),
            "warnings": result.warnings,
            "errors": result.errors,
            "capability": sr.capability.as_dict(),
        }
        (self._artifacts_dir / "generation_summary.json").write_text(
            json.dumps(summary, indent=2, ensure_ascii=False, default=str) + "\n",
            encoding="utf-8",
        )
        files.append("generation_summary.json")

        (self._artifacts_dir / "README.md").write_text(_readme(summary), encoding="utf-8")
        files.append("README.md")
        return files


def _compat_dict(result: GenerationResult) -> dict | None:
    c = result.template_compatibility
    if c is None:
        return None
    return {
        "compatible": c.compatible,
        "expected_fingerprint_id": c.expected_fingerprint_id,
        "actual_fingerprint_id": c.actual_fingerprint_id,
        "file_sha256": c.file_sha256,
    }


def _readme(summary: dict) -> str:
    return f"""# Excel writer — Sprint 3.7B

Generación del REMASEP escribiendo `PendingWrite` (contrato del Sprint 3.7A) en
una **copia** de la plantilla oficial con Microsoft Excel Desktop (COM, Windows).

- estado: `{summary['status']}`  ·  modo: `{summary['mode']}`  ·
  etiqueta: `{summary['submission_label']}`
- writer_integrity_status: `{summary['writer_integrity_status']}`
- control_status: `{summary['control_status']}`
  (total_errors={summary['control_total_errors']};
  de módulos aún no producidos={summary['control_errors_from_pending_modules']})
- celdas escritas y verificadas: {summary['written_cells']} / {summary['pending_write_count']}
- plantilla intacta (SHA256): {summary.get('template_unchanged', 'n/a')}
- historical_reference_cache_status: `{summary['historical_reference_cache_status']}`

`writer_integrity_status` es **independiente** de `control_status`: la plantilla
también necesita EGRESOS / recursos / tabla quirúrgica / metadatos que este
sprint todavía no produce, así que CONTROL puede reportar errores sin que el
writer haya fallado.

`PASS_INTERNAL_VALIDATION` (hoja CONTROL) **no** equivale a `MINSAL_APPROVED`.
El archivo se marca `{summary['submission_label']}` mientras el filtro por ESTADO
siga pendiente de confirmación funcional.

## Archivos

`generation_summary.json`, `write_audit.csv`, `target_verification.csv`,
`formula_integrity.csv`, `control_result.csv`, `template_verification.json`.

## Capacidad de plataforma

```
{json.dumps(summary['capability'], indent=2, ensure_ascii=False)}
```
"""
