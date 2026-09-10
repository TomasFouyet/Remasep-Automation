"""Sprint 3.8 — congela los assets de runtime de producción (dev-only).

Lee el workbook legacy (``GENERACION DATOS REMASEP.xlsx``) y el manifiesto de
escritura del Sprint 3.6 (``artifacts/writable_target_mapping/…``) y **congela**
el conocimiento estructural necesario para producir las 1122 escrituras en
``config/runtime_2026/``:

    bundle.yaml            contrato + sha256 de cada asset
    write_manifest.csv     1122 instrucciones (contrato semántico, sin PII)
    metric_catalog.csv     fórmulas legacy por métrica (1122 + deps transitivas)
    detail_contract.yaml   hoja de detalle, mapa de columnas, refs de criterio
    zero_policy.yaml       resolución de la política del cero (evidence-derived 3.7A)

Este script es una herramienta de desarrollo: **su salida se versiona**, el
workbook legacy NO. Incluye un escaneo anti-PII y una auto-verificación de que
los assets reproducen exactamente los mismos PendingWrites que el path validado.

Uso:

    python scripts/build_runtime_assets.py \\
        --generator "data/local/GENERACION DATOS REMASEP.xlsx" \\
        --manifest artifacts/writable_target_mapping/write_manifest_ready.csv
"""

from __future__ import annotations

import argparse
import csv
import hashlib
import re
import sys
from pathlib import Path

import yaml
from close_legacy_aggregations import KIND_BASE, KIND_DERIVED, close_legacy_aggregations
from compare_legacy_aggregations import _make_resolver

from remasep.core.errors import RemasepError
from remasep.services.medinet_input_reconciliation import (
    LEGACY_EXCLUDED_STATES,
    LEGACY_KEPT_STATES,
)

DEFAULT_GENERATOR = "data/local/GENERACION DATOS REMASEP.xlsx"
DEFAULT_MANIFEST = "artifacts/writable_target_mapping/write_manifest_ready.csv"
DEFAULT_ZERO_POLICY = "config/metric_value_producer_2026/zero_write_policy.yaml"
DEFAULT_OUT = "config/runtime_2026"

# Sprint 3.9 fase 2 — regla ESTADO CONFIRMADA por el responsable funcional
# (Fundación Gantz / Jacqueline). Coincide con los estados observados del
# proceso legacy (`LEGACY_KEPT_STATES`); se cruza contra ese constante para
# detectar desalineación en tiempo de build.
ESTADO_FILTER_STATUS = "CONFIRMED"
ESTADO_INCLUDED_STATES = (
    "Atendido", "En Sala de Espera", "Atención Pausada", "En Atención",
)
ESTADO_EXCLUDED_STATES = (
    "Cancelado", "No Se Presenta", "Agendado", "Confirmado", "Re-Agendado",
)

_MANIFEST_COLUMNS = (
    "instruction_id", "source_metric_id", "source_semantic_signature",
    "target_sheet", "target_cell", "target_semantic_signature",
    "expected_value_type", "template_fingerprint_id", "policy_version", "match_status",
)
_CATALOG_COLUMNS = ("metric_id", "sheet", "cell", "kind", "value_ref_coords", "formula")

# Patrones que NUNCA deben aparecer en un asset de runtime.
_PII_PATTERNS = (
    re.compile(r"\b\d{7,8}-[\dkK]\b"),          # RUT
    re.compile(r"\b\d{1,2}/\d{1,2}/\d{4}\b"),   # fecha individual dd/mm/aaaa
    re.compile(r"\b\d{4}-\d{2}-\d{2}\b"),       # fecha ISO individual
    re.compile(r"[\w.+-]+@[\w-]+\.[\w.-]+"),    # email
)


def _sha256(path: Path) -> str:
    digest = hashlib.sha256()
    with path.open("rb") as handle:
        for chunk in iter(lambda: handle.read(1 << 20), b""):
            digest.update(chunk)
    return digest.hexdigest()


def _load_manifest(path: Path) -> list[dict]:
    with path.open(encoding="utf-8", newline="") as handle:
        reader = csv.DictReader(handle)
        rows = list(reader)
    missing = [c for c in _MANIFEST_COLUMNS if c not in (reader.fieldnames or ())]
    if missing:
        raise RemasepError(f"manifiesto sin columnas {missing}")
    return rows


def _build_catalog(generator: Path, manifest_ids: list[str]):
    closure = close_legacy_aggregations(generator)
    node_by_coord = {}
    node_by_id = {}
    for (sheet, cell), node in closure.nodes.items():
        if node.kind in (KIND_BASE, KIND_DERIVED) and node.medinet:
            node_by_coord[(sheet, cell)] = node
            node_by_id[node.metric_id] = node

    needed: set[str] = set()
    stack = list(manifest_ids)
    while stack:
        mid = stack.pop()
        if mid in needed:
            continue
        needed.add(mid)
        node = node_by_id.get(mid)
        if node is None:
            raise RemasepError(f"métrica del manifiesto ausente del cierre legacy: {mid}")
        for coord in node.value_ref_coords:
            dep = node_by_coord.get((node.sheet, coord))
            if dep is None:
                raise RemasepError(f"dependencia {node.sheet}!{coord} de {mid} no evaluable")
            stack.append(dep.metric_id)

    rows = []
    for mid in sorted(needed):
        node = node_by_id[mid]
        rows.append(
            {
                "metric_id": mid,
                "sheet": node.sheet,
                "cell": node.cell,
                "kind": node.kind,
                "value_ref_coords": "|".join(node.value_ref_coords),
                "formula": node.formula,
            }
        )

    # refs de celda usadas como CRITERIO dentro de COUNTIF -> congelar su valor
    from remasep.services.legacy_aggregation import parse_derived_formula

    resolvers = {s: _make_resolver(d) for s, d in closure.cached.items()}
    criterion_refs: dict[tuple[str, str], object] = {}
    for mid in needed:
        node = node_by_id[mid]
        parsed = parse_derived_formula(node.formula, closure.detail_sheet, resolvers[node.sheet])
        for ref in parsed.criterion_refs:
            key = (node.sheet, ref.replace("$", "").upper())
            criterion_refs[key] = resolvers[node.sheet](ref)

    detail_contract = {
        "detail_sheet": closure.detail_sheet,
        "detail_columns_map": dict(closure.detail_columns_map),
        "criterion_refs": [
            {"sheet": s, "cell": c, "value": v} for (s, c), v in sorted(criterion_refs.items())
        ],
    }
    return rows, detail_contract, closure


def _scan_pii(paths: list[Path]) -> list[str]:
    hits: list[str] = []
    for path in paths:
        text = path.read_text(encoding="utf-8")
        for pat in _PII_PATTERNS:
            for m in pat.finditer(text):
                hits.append(f"{path.name}: {pat.pattern!r} -> {m.group(0)!r}")
    return hits


def _write_csv(path: Path, columns: tuple[str, ...], rows: list[dict]) -> None:
    with path.open("w", encoding="utf-8", newline="") as handle:
        writer = csv.DictWriter(handle, fieldnames=list(columns))
        writer.writeheader()
        writer.writerows(rows)


def build(generator: Path, manifest: Path, zero_policy: Path, out_dir: Path) -> dict:
    out_dir.mkdir(parents=True, exist_ok=True)
    manifest_rows = _load_manifest(manifest)
    manifest_ids = [r["source_metric_id"] for r in manifest_rows]
    fps = {r["template_fingerprint_id"] for r in manifest_rows}
    if len(fps) != 1:
        raise RemasepError(f"el manifiesto tiene {len(fps)} fingerprints")
    fingerprint = fps.pop()

    catalog_rows, detail_contract, _closure = _build_catalog(generator, manifest_ids)

    # zero policy: se congela sólo la resolución (evidence-derived en 3.7A)
    zdoc = yaml.safe_load(zero_policy.read_text(encoding="utf-8")) or {}
    resolution = str(zdoc.get("resolution", "UNRESOLVED"))

    # --- escribir assets ------------------------------------------
    manifest_out = out_dir / "write_manifest.csv"
    _write_csv(
        manifest_out,
        _MANIFEST_COLUMNS,
        [{c: r[c] for c in _MANIFEST_COLUMNS} for r in manifest_rows],
    )
    catalog_out = out_dir / "metric_catalog.csv"
    _write_csv(catalog_out, _CATALOG_COLUMNS, catalog_rows)

    detail_out = out_dir / "detail_contract.yaml"
    detail_out.write_text(
        "# Sprint 3.8 — congelado por scripts/build_runtime_assets.py. NO editar a mano.\n"
        "# Contrato estructural de la hoja de detalle Medinet + refs de criterio.\n"
        + yaml.safe_dump(detail_contract, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    zero_out = out_dir / "zero_policy.yaml"
    zero_out.write_text(
        "# Sprint 3.8 — resolución congelada de la política del cero (evidence-derived 3.7A).\n"
        "# Fuente: config/metric_value_producer_2026/zero_write_policy.yaml\n"
        + yaml.safe_dump(
            {"version": "runtime_2026", "resolution": resolution,
             "provenance": "metric_value_producer_2026 / Sprint 3.7A"},
            allow_unicode=True, sort_keys=False,
        ),
        encoding="utf-8",
    )

    # --- regla ESTADO confirmada (Sprint 3.9 fase 2) ---------------
    if set(ESTADO_INCLUDED_STATES) != set(LEGACY_KEPT_STATES) or set(
        ESTADO_EXCLUDED_STATES
    ) != set(LEGACY_EXCLUDED_STATES):
        raise RemasepError(
            "la regla ESTADO confirmada no coincide con LEGACY_KEPT/EXCLUDED_STATES"
        )
    estado_out = out_dir / "estado_filter.yaml"
    estado_out.write_text(
        "# Sprint 3.9 fase 2 — regla funcional del campo ESTADO de Medinet.\n"
        "# CONFIRMADA por el responsable funcional (Fundación Gantz / Jacqueline).\n"
        "# NO editar a mano; regenerar con scripts/build_runtime_assets.py.\n"
        "# Normalización: str.strip().casefold() (espacios externos + minúsculas, con tildes).\n"
        + yaml.safe_dump(
            {
                "version": "runtime_2026",
                "status": ESTADO_FILTER_STATUS,
                "confirmed_source": "Fundación Gantz / Jacqueline — Sprint 3.9 fase 2",
                "normalization": "strip_casefold",
                "included_states": list(ESTADO_INCLUDED_STATES),
                "excluded_states": list(ESTADO_EXCLUDED_STATES),
                "audit": {
                    "period": "Julio 2026",
                    "scope_before": 2114,
                    "scope_after": 1364,
                    "reference": "docs/MEDINET_ESTADO_AUDIT.md",
                },
            },
            allow_unicode=True, sort_keys=False,
        ),
        encoding="utf-8",
    )

    pii = _scan_pii([manifest_out, catalog_out, detail_out, zero_out, estado_out])
    if pii:
        for h in pii:
            print(f"PII SOSPECHOSA: {h}", file=sys.stderr)
        raise RemasepError(f"escaneo anti-PII encontró {len(pii)} coincidencia(s); abortado")

    assets = {
        "write_manifest.csv": _sha256(manifest_out),
        "metric_catalog.csv": _sha256(catalog_out),
        "detail_contract.yaml": _sha256(detail_out),
        "zero_policy.yaml": _sha256(zero_out),
        "estado_filter.yaml": _sha256(estado_out),
    }
    bundle = {
        "version": "runtime_2026",
        "generated_by": "scripts/build_runtime_assets.py (Sprint 3.8 + 3.9 fase 2)",
        "template_fingerprint_id": fingerprint,
        "expected_instruction_count": len(manifest_rows),
        "detail_sheet": detail_contract["detail_sheet"],
        "estado_filter_status": ESTADO_FILTER_STATUS,
        "zero_write_policy": resolution,
        "write_manifest": "write_manifest.csv",
        "metric_catalog": "metric_catalog.csv",
        "detail_contract": "detail_contract.yaml",
        "zero_policy": "zero_policy.yaml",
        "estado_filter": "estado_filter.yaml",
        "assets": assets,
        "notes": (
            "Assets versionados con la app. GENERACION DATOS REMASEP.xlsx NO forma "
            "parte del producto; sólo se usó para congelar este contrato."
        ),
    }
    (out_dir / "bundle.yaml").write_text(
        "# Sprint 3.8 — contrato de runtime. Congelado; regenerar con "
        "scripts/build_runtime_assets.py.\n"
        + yaml.safe_dump(bundle, allow_unicode=True, sort_keys=False),
        encoding="utf-8",
    )
    return {
        "fingerprint": fingerprint,
        "instructions": len(manifest_rows),
        "catalog_rows": len(catalog_rows),
        "criterion_refs": len(detail_contract["criterion_refs"]),
        "zero_write_policy": resolution,
        "estado_filter_status": ESTADO_FILTER_STATUS,
        "estado_included_states": list(ESTADO_INCLUDED_STATES),
    }


def _self_check(out_dir: Path, generator: Path, manifest: Path) -> dict:
    """Verifica que los assets reproducen exactamente los PendingWrites del path
    validado. Desde 3.9 fase 2 producción aplica el filtro ESTADO CONFIRMADO, así
    que el lado OLD también se filtra a los estados confirmados."""
    from build_metric_values import (
        _diagnostic_frame,
        build_legacy_reference,
        load_manifest,
        run_producer,
    )

    from remasep.services.common import Period
    from remasep.services.medinet_analysis import processing_scope_frame
    from remasep.services.metric_value_producer import join_metric_values_with_manifest
    from remasep.services.production_pipeline import build_production_pending_writes
    from remasep.services.runtime_assets import load_runtime_bundle

    medinet = "data/local/detalle_citas - 2026-09-07T123630.940.xlsx"
    if not Path(medinet).is_file():
        return {"skipped": "no hay export Medinet local para comparar"}
    period = Period(7, 2026)

    # OLD path (con el filtro ESTADO confirmado aplicado igual que producción)
    reference = build_legacy_reference(generator)
    old_manifest = load_manifest(manifest)
    old_frame = _diagnostic_frame(processing_scope_frame(medinet, period), ESTADO_INCLUDED_STATES)
    old_run = run_producer(
        old_frame, reference, old_manifest, period, mode="LEGACY_EQUIVALENCE_DIAGNOSTIC"
    )
    old_pending = join_metric_values_with_manifest(old_run.metric_values, old_manifest)

    # NEW path
    bundle = load_runtime_bundle(out_dir)
    new = build_production_pending_writes(medinet, period, bundle=bundle)

    old_map = {p.instruction_id: (p.target_sheet, p.target_cell, p.value) for p in old_pending}
    new_map = {p.instruction_id: (p.target_sheet, p.target_cell, p.value) for p in new.pending_writes}
    diffs = [k for k in set(old_map) | set(new_map) if old_map.get(k) != new_map.get(k)]
    return {
        "old_pending": len(old_pending),
        "new_pending": len(new.pending_writes),
        "identical": not diffs,
        "diffs": diffs[:10],
    }


def build_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(prog="build_runtime_assets.py")
    parser.add_argument("--generator", default=DEFAULT_GENERATOR)
    parser.add_argument("--manifest", default=DEFAULT_MANIFEST)
    parser.add_argument("--zero-policy", default=DEFAULT_ZERO_POLICY)
    parser.add_argument("--out", default=DEFAULT_OUT)
    parser.add_argument("--no-self-check", action="store_true")
    return parser


def main(argv: list[str] | None = None) -> int:
    args = build_parser().parse_args(argv)
    generator, manifest = Path(args.generator), Path(args.manifest)
    for p in (generator, manifest, Path(args.zero_policy)):
        if not p.is_file():
            print(f"ERROR: no existe {p}", file=sys.stderr)
            return 2
    try:
        summary = build(generator, manifest, Path(args.zero_policy), Path(args.out))
    except RemasepError as exc:
        print(f"ERROR: {exc}", file=sys.stderr)
        return 2
    print(f"assets congelados en {args.out}:")
    for k, v in summary.items():
        print(f"  {k}: {v}")
    if not args.no_self_check:
        check = _self_check(Path(args.out), generator, manifest)
        print(f"self-check OLD vs NEW: {check}")
        if check.get("identical") is False:
            print("ERROR: los assets NO reproducen los PendingWrites del path validado", file=sys.stderr)
            return 2
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
