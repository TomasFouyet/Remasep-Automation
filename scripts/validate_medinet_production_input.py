"""Production Medinet Input Contract & Period Scope Validation (hotfix).

Formaliza que el input de producción es el export directo **"Detalle de citas"**
de Medinet, valida el alcance de período (registros de OTROS meses no se
descartan, pero quedan fuera del REMASEP mensual) y reconcilia, de forma
privacy-safe, la diferencia de recuento contra la referencia legacy
``GENERACION DATOS REMASEP.xlsx``.

Sólo lectura: no modifica ni guarda ningún archivo; verifica el SHA256 de ambos
antes y después. **No** exporta filas individuales ni PII.

Uso:

    python scripts/validate_medinet_production_input.py \\
        "data/local/detalle_citas - 2026-09-07T123630.940.xlsx" \\
        "data/local/GENERACION DATOS REMASEP.xlsx" --month 7 --year 2026
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import json
import sys
from datetime import UTC, datetime
from pathlib import Path

from remasep.core.errors import SourceValidationError
from remasep.services import medinet_input_reconciliation as rec
from remasep.services.common import Period

DEFAULT_OUTPUT = "artifacts/medinet_production_input_validation"
DEFAULT_DIRECT = "data/local/detalle_citas - 2026-09-07T123630.940.xlsx"
DEFAULT_LEGACY = "data/local/GENERACION DATOS REMASEP.xlsx"


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _csv(path: Path, header: list[str], rows) -> None:
    with path.open("w", newline="", encoding="utf-8") as handle:
        writer = csv.writer(handle)
        writer.writerow(header)
        writer.writerows(rows)


_README = """# Medinet production input validation (hotfix)

Salida de `scripts/validate_medinet_production_input.py`.

## Contrato de input

| rol | archivo | uso |
| --- | --- | --- |
| **PRODUCCIÓN** | export directo de Medinet **"Detalle de citas"** (`detalle_citas - …xlsx`) | único input real de la app |
| **REFERENCIA LEGACY** | `GENERACION DATOS REMASEP.xlsx` | sólo ingeniería inversa; la app **no** depende de él |

## Alcance de período

Un export Medinet trae **varios meses**. El REMASEP mensual se calcula SOLO sobre
`processing_scope_records` = registros **estructuralmente válidos** ∩ **del mes/año
seleccionados**. Los registros de otros períodos **no se descartan** del archivo;
sólo quedan fuera del cálculo.

## Reconciliación (privacy-safe)

Comparación por **multiset** de huellas semánticas (`DIA_CITA · FECHA_NACIMIENTO ·
SEXO · SUCURSAL · ESPECIALIDAD · TIPO_DE_CITA · PRESTACION · ESTADO · MODALIDAD ·
PRESTACION_REALIZADA`). **Sin** RUN, nombre, teléfono, id de paciente.

## Hipótesis de ESTADO — NO es una regla

`estado_processing_hypothesis` es la reconstrucción del comportamiento observado
del generador legacy (`status: CURRENT_LEGACY_BEHAVIOR_PENDING_FUNCTIONAL_CONFIRMATION`).
**No** se implementa como filtro en la clasificación: queda documentada para
confirmación funcional del cliente.

## Archivos

- `summary.json` — todas las cantidades (§10 del hotfix).
- `period_scope_summary.csv` — desglose físico / válido / período / alcance, por archivo.
- `legacy_subset_comparison.csv` — multiset legacy ⊆ export directo, con y sin ESTADO.
- `extra_records_by_status.csv` / `_by_appointment_type.csv` / `_by_branch.csv` /
  `_by_modality.csv` — agregados privacy-safe de los registros que el proceso
  antiguo dejó fuera.
- `README.md`.

Ninguna fila individual, ningún dato personal.
"""


def run(direct_path: str, legacy_path: str, period: Period, output_dir: str) -> dict:
    direct, legacy = Path(direct_path), Path(legacy_path)
    for p in (direct, legacy):
        if not p.is_file():
            raise SourceValidationError(f"no es un archivo: {p}")
    sha_pre = (_sha256(direct), _sha256(legacy))

    result = rec.reconcile(direct, legacy, period)

    sha_post = (_sha256(direct), _sha256(legacy))
    if sha_pre != sha_post:
        raise SourceValidationError("¡el SHA256 de un archivo fuente cambió durante el análisis!")

    out = Path(output_dir)
    out.mkdir(parents=True, exist_ok=True)
    files: list[Path] = []

    def _scope_rows(label: str, s: rec.PeriodScopeBreakdown) -> list:
        return [
            label, s.physical_records, s.structurally_valid_records,
            s.structurally_invalid_records, s.in_period_records,
            s.out_of_period_records, s.processing_scope_records,
        ]

    p = out / "period_scope_summary.csv"
    _csv(p, [
        "file", "physical_records", "structurally_valid_records",
        "structurally_invalid_records", "in_period_records", "out_of_period_records",
        "processing_scope_records",
    ], [
        _scope_rows(f"direct_export ({result.direct_source})", result.direct_scope),
        _scope_rows(f"legacy_reference ({result.legacy_source})", result.legacy_scope),
    ])
    files.append(p)

    p = out / "legacy_subset_comparison.csv"
    _csv(p, [
        "fingerprint", "subset_records_legacy", "superset_records_direct_period",
        "matched_subset_records", "unmatched_subset_records", "superset_excess_records",
        "is_exact_multiset_subset",
    ], [
        ["FULL (incl ESTADO)", *_cmp_row(result.subset_full_fp)],
        ["NO_ESTADO", *_cmp_row(result.subset_no_estado_fp)],
    ])
    files.append(p)

    agg_files = {
        "ESTADO": "extra_records_by_status.csv",
        "TIPO_DE_CITA": "extra_records_by_appointment_type.csv",
        "SUCURSAL": "extra_records_by_branch.csv",
        "MODALIDAD": "extra_records_by_modality.csv",
        "ESPECIALIDAD": "extra_records_by_specialty.csv",
    }
    for dim, fname in agg_files.items():
        p = out / fname
        _csv(p, [dim, "record_count"], result.extra_aggregates.get(dim, []))
        files.append(p)

    summary = _summary(result, sha_pre)
    p = out / "summary.json"
    p.write_text(json.dumps(summary, indent=2, ensure_ascii=False), encoding="utf-8")
    files.append(p)

    p = out / "README.md"
    p.write_text(_README, encoding="utf-8")
    files.append(p)

    return {"result": result, "summary": summary, "files": files}


def _cmp_row(c: rec.SubsetComparison) -> list:
    return [
        c.subset_records, c.superset_records, c.matched_subset_records,
        c.unmatched_subset_records, c.superset_excess_records,
        "yes" if c.is_exact_subset else "no",
    ]


def _summary(result: rec.ReconciliationResult, sha_pre: tuple[str, str]) -> dict:
    d, lg = result.direct_scope, result.legacy_scope
    e = result.estado
    return {
        "generated_at": datetime.now(UTC).strftime("%Y-%m-%dT%H:%M:%SZ"),
        "period": result.period.label,
        "input_contract": {
            "production_input": "Medinet direct export 'Detalle de citas'",
            "production_file": result.direct_source,
            "legacy_reference": result.legacy_source,
            "note": "La aplicación NO depende del generador legacy para funcionar.",
        },
        "source_hash_verified_unchanged": True,
        "direct_export_sha256": sha_pre[0],
        "legacy_reference_sha256": sha_pre[1],
        # --- alcance de período (export directo de producción) ---
        "raw_file_records": d.physical_records,
        "structurally_valid_records": d.structurally_valid_records,
        "structurally_invalid_records": d.structurally_invalid_records,
        "july_records": d.in_period_records,
        "out_of_period_records": d.out_of_period_records,
        "processing_scope_records": d.processing_scope_records,
        # --- referencia legacy ---
        "legacy_reference_records": lg.structurally_valid_records,
        "legacy_reference_in_period_records": lg.in_period_records,
        # --- comparación multiset ---
        "legacy_matched_to_direct_export_full_fp": result.subset_full_fp.matched_subset_records,
        "legacy_unmatched_full_fp": result.subset_full_fp.unmatched_subset_records,
        "legacy_is_exact_subset_full_fp": result.subset_full_fp.is_exact_subset,
        "legacy_matched_to_direct_export_no_estado_fp":
            result.subset_no_estado_fp.matched_subset_records,
        "legacy_unmatched_no_estado_fp": result.subset_no_estado_fp.unmatched_subset_records,
        "legacy_is_exact_subset_no_estado_fp": result.subset_no_estado_fp.is_exact_subset,
        "direct_export_excess_records": result.subset_no_estado_fp.superset_excess_records,
        # --- extras por ESTADO ---
        "extra_by_status": dict(result.extra_aggregates.get("ESTADO", [])),
        # --- hipótesis ESTADO (NO regla) ---
        "estado_processing_hypothesis": {
            "status": e.status,
            "kept_states": list(e.kept_states),
            "excluded_states": list(e.excluded_states),
            "direct_period_by_estado": dict(e.direct_period_by_estado),
            "direct_kept_records": e.direct_kept_records,
            "direct_excluded_records": e.direct_excluded_records,
            "legacy_active_records": e.legacy_active_records,
            "kept_states_reproduce_legacy_count_exactly":
                e.reproduces_legacy_count_exactly,
        },
        "notes": list(result.notes),
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        prog="validate_medinet_production_input.py",
        description=(
            "Contrato de input Medinet + validación de alcance de período + "
            "reconciliación privacy-safe con la referencia legacy. Sólo lectura."
        ),
    )
    parser.add_argument("direct_export", nargs="?", default=DEFAULT_DIRECT)
    parser.add_argument("legacy_reference", nargs="?", default=DEFAULT_LEGACY)
    parser.add_argument("--month", type=int, default=7)
    parser.add_argument("--year", type=int, default=2026)
    parser.add_argument("--output", "-o", default=DEFAULT_OUTPUT)
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    period = Period(args.month, args.year)
    try:
        outcome = run(args.direct_export, args.legacy_reference, period, args.output)
    except SourceValidationError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2

    s = outcome["summary"]
    e = s["estado_processing_hypothesis"]
    print(f"Período: {s['period']}")
    print(f"  input de producción: {s['input_contract']['production_file']}")
    print(f"  referencia legacy:   {s['input_contract']['legacy_reference']}")
    print(f"  hash de ambos fuentes verificado sin cambios: "
          f"{s['source_hash_verified_unchanged']}")
    print("  --- alcance de período (export directo) ---")
    print(f"    raw_file_records          {s['raw_file_records']}")
    print(f"    structurally_valid        {s['structurally_valid_records']}")
    print(f"    july_records (in period)  {s['july_records']}")
    print(f"    out_of_period_records     {s['out_of_period_records']}")
    print(f"    processing_scope_records  {s['processing_scope_records']}")
    print("  --- reconciliación con legacy ---")
    print(f"    legacy_reference_records  {s['legacy_reference_records']}")
    print(f"    legacy matched (full fp)      {s['legacy_matched_to_direct_export_full_fp']} "
          f"(exact subset: {s['legacy_is_exact_subset_full_fp']})")
    print(f"    legacy matched (sin ESTADO)   {s['legacy_matched_to_direct_export_no_estado_fp']} "
          f"(exact subset: {s['legacy_is_exact_subset_no_estado_fp']})")
    print(f"    direct_export_excess          {s['direct_export_excess_records']}")
    print(f"    extra_by_status               {s['extra_by_status']}")
    print("  --- hipótesis ESTADO (NO regla) ---")
    print(f"    status: {e['status']}")
    print(f"    kept:     {e['kept_states']} -> {e['direct_kept_records']} registros")
    print(f"    excluded: {e['excluded_states']} -> {e['direct_excluded_records']} registros")
    print(f"    reproduce exactamente los {e['legacy_active_records']} legacy: "
          f"{e['kept_states_reproduce_legacy_count_exactly']}")
    print(f"Artefactos en {args.output}:")
    for f in outcome["files"]:
        print(f"  - {f}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
