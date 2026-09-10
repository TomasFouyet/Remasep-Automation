"""CLI de generación del REMASEP con Microsoft Excel Desktop (COM).

Dos caminos claramente separados:

- ``--mode production`` (Sprint 3.8): construye los ``PendingWrite`` usando
  **sólo** el export Medinet + los assets de runtime versionados
  (``config/runtime_2026/``). **No** abre ``GENERACION DATOS REMASEP.xlsx`` ni
  lee ``artifacts/``.
- ``--mode diagnostic``: reproduce el legacy usando el workbook de
  reverse-engineering. Es un camino **dev-only**; su dependencia del workbook
  legacy vive detrás de un import perezoso y nunca se toca en producción.

En Linux/WSL no se intenta COM: se informa que se necesita Microsoft Excel
Desktop y se termina con código de salida controlado.

Uso:

    python scripts/generate_remasep.py \\
        --medinet "data/local/detalle_citas - 2026-09-07T123630.940.xlsx" \\
        --period 2026-07 \\
        --template "data/local/REMASEP_V1.4 Julio 2026.xlsm" \\
        --output outputs/REMASEP_2026_07_DRAFT.xlsm \\
        --mode production
"""

from __future__ import annotations

import argparse
import sys
from pathlib import Path

from remasep.core.errors import RemasepError
from remasep.services.common import Period, month_name  # noqa: F401 - month_name re-export
from remasep.services.excel_writer import (
    MODE_DIAGNOSTIC_REFERENCE,
    MODE_PRODUCTION,
    STATUS_EXCEL_UNAVAILABLE,
)
from remasep.services.generation_service import GenerationService
from remasep.services.production_pipeline import build_production_pending_writes
from remasep.services.runtime_assets import RuntimeAssetError, load_runtime_bundle

DEFAULT_TEMPLATE = "data/local/REMASEP_V1.4 Julio 2026.xlsm"
DEFAULT_MEDINET = "data/local/detalle_citas - 2026-09-07T123630.940.xlsx"
DEFAULT_GENERATOR = "data/local/GENERACION DATOS REMASEP.xlsx"  # sólo --mode diagnostic

_MODE_ALIASES = {
    "production": MODE_PRODUCTION,
    "diagnostic": MODE_DIAGNOSTIC_REFERENCE,
    "diagnostic_reference": MODE_DIAGNOSTIC_REFERENCE,
}


def _parse_period(text: str) -> Period:
    try:
        year_s, month_s = text.split("-", 1)
        return Period(month=int(month_s), year=int(year_s))
    except (ValueError, TypeError) as exc:
        raise RemasepError(f"período inválido {text!r}, se esperaba AAAA-MM") from exc


# ---------------------------------------------------------------------------
# Construcción de PendingWrites por camino
# ---------------------------------------------------------------------------


def _production_inputs(args, period: Period):
    """Camino de PRODUCCIÓN: sólo Medinet + runtime bundle."""
    bundle = load_runtime_bundle()
    result = build_production_pending_writes(args.medinet, period, bundle=bundle)
    if result.run.value_type_conflicts or result.run.unsupported:
        raise RemasepError(
            f"{len(result.run.value_type_conflicts)} conflicto(s) de tipo, "
            f"{len(result.run.unsupported)} métrica(s) sin soporte"
        )
    return {
        "pending": result.pending_writes,
        "manifest_instruction_ids": bundle.instruction_ids,
        "expected_fingerprint_id": bundle.template_fingerprint_id,
        "zero_write_policy": bundle.zero_write_policy,
        "completeness": result.completeness,
        "scope_records": result.scope_records,
    }


def _diagnostic_inputs(args, period: Period):
    """Camino DEV-ONLY: reproduce el legacy con el workbook de reverse-engineering.

    El import del motor legacy es **perezoso**: nunca entra al grafo de imports
    del camino de producción.
    """
    import csv as _csv

    import yaml as _yaml
    from build_metric_values import (
        _diagnostic_frame,
        build_legacy_reference,
        load_manifest,
        run_producer,
    )

    from remasep.services.medinet_analysis import processing_scope_frame
    from remasep.services.medinet_input_reconciliation import LEGACY_KEPT_STATES
    from remasep.services.metric_value_producer import join_metric_values_with_manifest

    reference = build_legacy_reference(args.generator)
    manifest = load_manifest(Path(args.manifest))
    frame = _diagnostic_frame(processing_scope_frame(args.medinet, period), LEGACY_KEPT_STATES)
    run = run_producer(frame, reference, manifest, period, mode="LEGACY_EQUIVALENCE_DIAGNOSTIC")
    if run.value_type_conflicts or run.unsupported:
        raise RemasepError(
            f"{len(run.value_type_conflicts)} conflicto(s) de tipo, "
            f"{len(run.unsupported)} sin soporte"
        )
    pending = join_metric_values_with_manifest(run.metric_values, manifest)
    with Path(args.manifest).open(encoding="utf-8", newline="") as _fh:
        fps = {r["template_fingerprint_id"] for r in _csv.DictReader(_fh)}
    zdoc = _yaml.safe_load(Path(args.zero_policy).read_text(encoding="utf-8")) or {}
    return {
        "pending": pending,
        "manifest_instruction_ids": [row.instruction_id for row in manifest],
        "expected_fingerprint_id": fps.pop() if len(fps) == 1 else "",
        "zero_write_policy": str(zdoc.get("resolution", "UNRESOLVED")),
        "completeness": None,
        "scope_records": run.scope_records,
    }


# ---------------------------------------------------------------------------
# CLI
# ---------------------------------------------------------------------------


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="generate_remasep.py",
        description=(
            "Genera el REMASEP escribiendo PendingWrite en una copia de la "
            "plantilla con Microsoft Excel Desktop. Requiere Windows + Excel."
        ),
    )
    parser.add_argument("--medinet", default=DEFAULT_MEDINET)
    parser.add_argument("--period", default="2026-07", help="AAAA-MM")
    parser.add_argument("--template", default=DEFAULT_TEMPLATE)
    parser.add_argument("--output", default=None, help="ruta .xlsm de salida (outputs/...)")
    parser.add_argument("--mode", choices=sorted(_MODE_ALIASES), default="production")
    parser.add_argument("--artifacts", default="artifacts/excel_writer")
    # sólo relevantes para --mode diagnostic (dev-only)
    parser.add_argument("--generator", default=DEFAULT_GENERATOR, help="[diagnostic] workbook legacy")
    parser.add_argument(
        "--manifest",
        default="artifacts/writable_target_mapping/write_manifest_ready.csv",
        help="[diagnostic] manifiesto de artifacts/",
    )
    parser.add_argument(
        "--zero-policy",
        default="config/metric_value_producer_2026/zero_write_policy.yaml",
        help="[diagnostic] política del cero",
    )
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    writer_mode = _MODE_ALIASES[args.mode]
    service = GenerationService(artifacts_dir=Path(args.artifacts))

    capability = service.capability(probe_com=True)
    print(f"plataforma: {capability.platform}  can_generate: {capability.can_generate}  "
          f"modo: {args.mode}")
    if not capability.can_generate:
        print(capability.user_message)
        print(f"  motivo: {capability.reason}")
        print(
            "\nPara la prueba real: ejecutar este comando en Windows con "
            "Python + pywin32 + Microsoft Excel Desktop instalados."
        )
        return 3

    try:
        period = _parse_period(args.period)
        if writer_mode == MODE_PRODUCTION:
            data = _production_inputs(args, period)
        else:
            data = _diagnostic_inputs(args, period)
    except RuntimeAssetError as exc:
        print(f"ERROR: {exc.user_message}", file=sys.stderr)
        print(f"  detalle técnico: {exc}", file=sys.stderr)
        return 2
    except RemasepError as exc:
        print(f"ERROR construyendo PendingWrite: {exc}", file=sys.stderr)
        return 2

    pending = data["pending"]
    try:
        sr = service.generate(
            mode=writer_mode,
            template_path=args.template,
            output_path=args.output,
            pending_writes=pending,
            manifest_instruction_ids=data["manifest_instruction_ids"],
            period_year=period.year,
            period_month=period.month,
            zero_write_policy=data["zero_write_policy"],
            expected_fingerprint_id=data["expected_fingerprint_id"],
        )
    except RemasepError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    r = sr.result
    print(f"\nestado: {r.status}  ({sr.submission_label})")
    print(f"  writer_integrity_status: {r.writer_integrity_status}")
    print(f"  control_status: {r.control_status}")
    print(f"  celdas escritas y verificadas: {r.written_cells} / {len(pending)}")
    print(f"  scope_records: {data['scope_records']}")
    print(f"  plantilla intacta: {r.template_unchanged}")
    if r.output_path:
        print(f"  salida: {r.output_path}")
    for w in r.warnings:
        print(f"  aviso: {w}")
    for e in r.errors:
        print(f"  error: {e}")
    if sr.artifact_files:
        print(f"artefactos en {sr.artifacts_dir}: {', '.join(sr.artifact_files)}")

    if r.status == STATUS_EXCEL_UNAVAILABLE:
        return 3
    return 0 if r.succeeded else 2


if __name__ == "__main__":
    raise SystemExit(main())
