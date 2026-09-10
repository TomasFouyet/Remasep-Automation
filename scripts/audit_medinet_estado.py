"""Sprint 3.9 — auditoría diagnóstica del campo ESTADO del export Medinet.

**Sólo lectura.** No modifica `production_pipeline`, los runtime assets, el
manifiesto, `MetricValueProducer`, el writer, la config ni ninguna regla. No
implementa ningún filtro por ESTADO: sólo *evalúa* la hipótesis legacy observada
para preparar evidencia y confirmarla con el cliente.

Compara dos alcances sobre el mismo Medinet / período:

    A. PERIOD_ONLY               = processing_scope_records (lo que hace producción hoy)
    B. LEGACY_STATE_HYPOTHESIS   = A ∩ {estados "incluidos" de la hipótesis legacy}

y mide el impacto de B sobre las 1122 métricas REMASEP write-ready.

Uso:

    python scripts/audit_medinet_estado.py \\
        --medinet "data/local/detalle_citas - 2026-09-07T123630.940.xlsx" \\
        --period 2026-07
"""

from __future__ import annotations

import argparse
import csv
import json
import sys
from dataclasses import dataclass
from pathlib import Path

from remasep.core.errors import RemasepError
from remasep.services.common import Period
from remasep.services.legacy_rules import load_legacy_rules
from remasep.services.medinet_analysis import processing_scope_frame
from remasep.services.medinet_input_reconciliation import (
    LEGACY_EXCLUDED_STATES,
    LEGACY_KEPT_STATES,
)
from remasep.services.metric_value_producer import (
    MODE_PRODUCTION,
    MetricValueProducer,
    build_rows,
    check_write_completeness,
    join_metric_values_with_manifest,
)
from remasep.services.runtime_assets import load_runtime_bundle

DEFAULT_MEDINET = "data/local/detalle_citas - 2026-09-07T123630.940.xlsx"
DEFAULT_OUT = "artifacts/medinet_estado_audit"
_ESTADO_FILTER_STATUS = "PENDING_FUNCTIONAL_CONFIRMATION"


def _parse_period(text: str) -> Period:
    year_s, month_s = text.split("-", 1)
    return Period(month=int(month_s), year=int(year_s))


@dataclass
class ScopeRun:
    label: str
    scope_records: int
    metric_values: int
    pending_writes: int
    nonzero: int
    completeness_ok: bool
    value_by_id: dict


def _run_scope(label: str, frame, bundle, period: Period, ruleset) -> ScopeRun:
    rows = build_rows(frame.to_dict("records"), bundle.detail_columns_map, ruleset)
    producer = MetricValueProducer(
        formula_index=bundle.formula_index,
        detail_sheet=bundle.detail_sheet,
        resolver_by_sheet=bundle.resolver_by_sheet(),
        rows=rows,
        period=period,
        scope_records=len(frame),
        mode=MODE_PRODUCTION,  # el modo sólo afecta el reporte, no el cálculo
    )
    ids = [w.source_metric_id for w in bundle.write_instructions]
    run = producer.produce_run(ids, expected_types=bundle.expected_value_types())
    pending = join_metric_values_with_manifest(run.metric_values, bundle.write_instructions)
    completeness = check_write_completeness(run.metric_values, bundle.write_instructions)
    values = {mv.source_metric_id: mv.value for mv in run.metric_values}
    return ScopeRun(
        label=label,
        scope_records=len(frame),
        metric_values=len(run.metric_values),
        pending_writes=len(pending),
        nonzero=sum(1 for v in values.values() if v),
        completeness_ok=completeness.ok,
        value_by_id=values,
    )


def audit(medinet_path: str | Path, period: Period) -> dict:
    bundle = load_runtime_bundle()
    ruleset = load_legacy_rules()

    frame = processing_scope_frame(medinet_path, period)
    total = len(frame)
    estado = frame["ESTADO"].astype("string")

    # --- 3. distribución ESTADO | cantidad | % -----------------------
    counts = estado.fillna("(vacío)").value_counts(dropna=False)
    distribution = [
        {
            "estado": str(name),
            "cantidad": int(n),
            "porcentaje": round(100.0 * int(n) / total, 2) if total else 0.0,
        }
        for name, n in counts.items()
    ]

    # --- 5-6. hipótesis legacy -------------------------------------
    norm = estado.str.strip().str.casefold()
    kept_norm = {s.strip().casefold() for s in LEGACY_KEPT_STATES}
    excl_norm = {s.strip().casefold() for s in LEGACY_EXCLUDED_STATES}
    mask_kept = norm.isin(kept_norm)
    kept_n = int(mask_kept.sum())
    excl_n = int(norm.isin(excl_norm).sum())
    other_states = sorted(
        {str(v) for v in estado.dropna().unique()}
        - {s for s in LEGACY_KEPT_STATES}
        - {s for s in LEGACY_EXCLUDED_STATES}
    )

    frame_b = frame.loc[mask_kept.fillna(False)].reset_index(drop=True)

    # --- 7. impacto sobre métricas -------------------------------
    run_a = _run_scope("A_PERIOD_ONLY", frame, bundle, period, ruleset)
    run_b = _run_scope("B_LEGACY_STATE_HYPOTHESIS", frame_b, bundle, period, ruleset)

    wi_by_id = {w.source_metric_id: w for w in bundle.write_instructions}
    changed = []
    for mid, a in run_a.value_by_id.items():
        b = run_b.value_by_id.get(mid)
        if a != b:
            changed.append(
                {
                    "source_metric_id": mid,
                    "target_sheet": wi_by_id[mid].target_sheet,
                    "target_cell": wi_by_id[mid].target_cell,
                    "value_A": a,
                    "value_B": b,
                    "delta_A_minus_B": a - b,
                }
            )
    changed.sort(key=lambda r: -abs(r["delta_A_minus_B"]))

    by_sheet: dict[str, dict] = {}
    for r in changed:
        s = by_sheet.setdefault(r["target_sheet"], {"metrics_changed": 0, "sum_abs_delta": 0})
        s["metrics_changed"] += 1
        s["sum_abs_delta"] += abs(r["delta_A_minus_B"])

    total_a = sum(run_a.value_by_id.values())
    total_b = sum(run_b.value_by_id.values())

    return {
        "generated_for": {
            "medinet": Path(medinet_path).name,
            "period": period.label,
            "estado_filter_status": _ESTADO_FILTER_STATUS,
            "note": (
                "Evidencia diagnóstica. NO se implementa ningún filtro por ESTADO; "
                "la hipótesis legacy NO se asume como la regla correcta."
            ),
        },
        "total_records": total,
        "total_is_2114": total == 2114,
        "estado_distribution": distribution,
        "hypothesis": {
            "included_states": list(LEGACY_KEPT_STATES),
            "excluded_states": list(LEGACY_EXCLUDED_STATES),
            "states_not_covered_by_hypothesis": other_states,
            "included_count": kept_n,
            "excluded_count": excl_n,
            "included_equals_1364": kept_n == 1364,
        },
        "scope_comparison": {
            "A_PERIOD_ONLY": {
                "scope_records": run_a.scope_records,
                "metric_values": run_a.metric_values,
                "pending_writes": run_a.pending_writes,
                "nonzero_metric_values": run_a.nonzero,
                "completeness_ok": run_a.completeness_ok,
                "sum_of_all_counts": total_a,
            },
            "B_LEGACY_STATE_HYPOTHESIS": {
                "scope_records": run_b.scope_records,
                "metric_values": run_b.metric_values,
                "pending_writes": run_b.pending_writes,
                "nonzero_metric_values": run_b.nonzero,
                "completeness_ok": run_b.completeness_ok,
                "sum_of_all_counts": total_b,
            },
            "metrics_changed_A_to_B": len(changed),
            "metrics_total": len(run_a.value_by_id),
            "all_changes_decrease_A_gt_B": all(r["delta_A_minus_B"] > 0 for r in changed),
            "sum_abs_delta": sum(abs(r["delta_A_minus_B"]) for r in changed),
            "sum_of_all_counts_delta_A_minus_B": total_a - total_b,
            "by_target_sheet": by_sheet,
            "changed_metrics": changed,
        },
    }


def _write_outputs(result: dict, out_dir: Path) -> list[str]:
    out_dir.mkdir(parents=True, exist_ok=True)
    files = []
    (out_dir / "summary.json").write_text(
        json.dumps(result, indent=2, ensure_ascii=False, default=str) + "\n", encoding="utf-8"
    )
    files.append("summary.json")
    with (out_dir / "estado_distribution.csv").open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(fh, fieldnames=["estado", "cantidad", "porcentaje"])
        w.writeheader()
        w.writerows(result["estado_distribution"])
    files.append("estado_distribution.csv")
    with (out_dir / "changed_metrics.csv").open("w", encoding="utf-8", newline="") as fh:
        w = csv.DictWriter(
            fh,
            fieldnames=[
                "source_metric_id", "target_sheet", "target_cell",
                "value_A", "value_B", "delta_A_minus_B",
            ],
        )
        w.writeheader()
        w.writerows(result["scope_comparison"]["changed_metrics"])
    files.append("changed_metrics.csv")
    return files


def _print_report(result: dict) -> None:
    r = result
    print(f"Medinet: {r['generated_for']['medinet']}   período: {r['generated_for']['period']}")
    print(f"estado_filter_status: {r['generated_for']['estado_filter_status']}\n")

    print("ESTADO | cantidad | % del total")
    print("-" * 44)
    for row in r["estado_distribution"]:
        print(f"  {row['estado']:<22} {row['cantidad']:>6}  {row['porcentaje']:>6.2f}%")
    print("-" * 44)
    print(f"  {'TOTAL':<22} {r['total_records']:>6}  100.00%")
    print(f"\n4. total == 2114 : {r['total_is_2114']}")

    h = r["hypothesis"]
    print("\n5-6. Hipótesis legacy observada")
    print(f"  incluidos {h['included_states']} -> {h['included_count']}")
    print(f"  excluidos {h['excluded_states']} -> {h['excluded_count']}")
    print(f"  estados no cubiertos por la hipótesis: {h['states_not_covered_by_hypothesis'] or 'ninguno'}")
    print(f"  incluidos == 1364 : {h['included_equals_1364']}")

    sc = r["scope_comparison"]
    print("\n7. Comparación de alcances (pipeline actual, sin modificar producción)")
    for key in ("A_PERIOD_ONLY", "B_LEGACY_STATE_HYPOTHESIS"):
        s = sc[key]
        print(
            f"  {key:<26} scope={s['scope_records']:>5}  MetricValues={s['metric_values']}  "
            f"PendingWrites={s['pending_writes']}  nonzero={s['nonzero_metric_values']}  "
            f"Σcounts={s['sum_of_all_counts']}  completeness_ok={s['completeness_ok']}"
        )
    print(
        f"  métricas que cambian A->B: {sc['metrics_changed_A_to_B']} / {sc['metrics_total']}  "
        f"(todas bajan A>B: {sc['all_changes_decrease_A_gt_B']})"
    )
    print(f"  Σ|Δ| = {sc['sum_abs_delta']}   Σcounts(A) - Σcounts(B) = {sc['sum_of_all_counts_delta_A_minus_B']}")
    print("  por hoja destino (métricas / Σ|Δ|):")
    for sheet, agg in sc["by_target_sheet"].items():
        print(f"    {sheet:<12} {agg['metrics_changed']:>4}  {agg['sum_abs_delta']:>6}")
    print("  top 15 |Δ|:")
    for row in sc["changed_metrics"][:15]:
        print(
            f"    {row['source_metric_id']:<28} {row['target_sheet']}!{row['target_cell']:<5} "
            f"A={row['value_A']:<4} B={row['value_B']:<4} Δ={row['delta_A_minus_B']}"
        )


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="audit_medinet_estado.py")
    parser.add_argument("--medinet", default=DEFAULT_MEDINET)
    parser.add_argument("--period", default="2026-07", help="AAAA-MM")
    parser.add_argument("--output", "-o", default=DEFAULT_OUT)
    parser.add_argument("--no-artifacts", action="store_true", help="no escribir CSV/JSON")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    if not Path(args.medinet).is_file():
        print(f"ERROR: no existe {args.medinet}", file=sys.stderr)
        return 2
    try:
        result = audit(args.medinet, _parse_period(args.period))
    except RemasepError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    _print_report(result)
    if not args.no_artifacts:
        files = _write_outputs(result, Path(args.output))
        print(f"\nartefactos en {args.output}: {', '.join(files)}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
