# Production runtime decoupling (Sprint 3.8)

`config/runtime_2026/` + `src/remasep/services/runtime_assets.py` +
`src/remasep/services/production_pipeline.py` +
`scripts/build_runtime_assets.py` (dev-only).

Hasta el Sprint 3.7B, generar el REMASEP dependía —en tiempo de ejecución— del
workbook de reverse-engineering `GENERACION DATOS REMASEP.xlsx` y de un CSV bajo
`artifacts/`. El Sprint 3.8 **desacopla el runtime de producción**: el workbook
legacy pasa a ser exclusivamente material de desarrollo.

## Qué necesita cada parte

### PRODUCTION INPUTS (los aporta el usuario)

1. **Export directo de Medinet** — "Detalle de citas" (`.xlsx`).
2. **Plantilla oficial REMASEP** — `.xlsm` (la escritura final la hace Excel COM).

### INTERNAL RUNTIME (versionado con la aplicación, `config/runtime_2026/`)

| asset | contenido | PII |
| --- | --- | --- |
| `bundle.yaml` | contrato: `template_fingerprint_id`, `expected_instruction_count`, `detail_sheet`, `estado_filter_status`, `zero_write_policy`, sha256 de cada asset | no |
| `write_manifest.csv` | 1122 instrucciones de escritura (contrato semántico del Sprint 3.6, sanitizado): `instruction_id`, `source_metric_id`, firmas semánticas, `target_sheet`/`target_cell`, `expected_value_type`, `template_fingerprint_id`, `policy_version`, `match_status` | no |
| `metric_catalog.csv` | 1224 fórmulas legacy por métrica (1122 WRITE_READY + 102 dependencias transitivas): `metric_id`, `sheet`, `cell`, `kind`, `value_ref_coords`, `formula` (`COUNTIF(S)` sobre la hoja de detalle) | no |
| `detail_contract.yaml` | `detail_sheet`, `detail_columns_map` (10 campos → letra de columna) y las 3 `criterion_refs` congeladas (`REMASEP_OD!A8/A9/A10` → su literal de sección) | no |
| `zero_policy.yaml` | `resolution: WRITE_ZERO` (evidence-derived en 3.7A) | no |

**Ausencia de PII confirmada**: los assets contienen sólo reglas, mappings, IDs
y metadatos estructurales. No hay nombres, RUT, fechas de nacimiento
individuales, prestaciones individuales, filas de Medinet ni valores mensuales.
Los literales de `metric_catalog.csv` (`"*5004020 - CONTROL ORTODONCIA
QUIRÚRGICO*"`, `">=25"`, `"...Santiago"`) son **definiciones de categoría
clínica** (código de prestación · sexo · tramo de edad · sucursal), no datos de
pacientes. La palabra `PACIENTE` que aparece en las firmas es un rótulo de
sección del formulario oficial (`"ATENCION AL PACIENTE"`), no un nombre.
`scripts/build_runtime_assets.py` corre un escaneo anti-PII (RUT / fecha
individual / email) que debe salir vacío o aborta.

### DEV-ONLY (NO forma parte del producto)

- `data/local/GENERACION DATOS REMASEP.xlsx` — reverse-engineering e histórico.
- `artifacts/` — salidas intermedias de los Sprints 3.x.
- `scripts/build_metric_values.py`, `scripts/close_legacy_aggregations.py`,
  `scripts/compare_legacy_aggregations.py`, `scripts/build_runtime_assets.py` —
  herramientas de análisis / regeneración de assets.
- `--mode diagnostic` de `scripts/generate_remasep.py` — reproduce el legacy para
  comparación; su import del motor legacy es **perezoso**, dentro de la función.

## Los dos caminos

| | PRODUCCIÓN | DIAGNOSTIC / DEV |
| --- | --- | --- |
| función | `production_pipeline.build_production_pending_writes()` | `build_metric_values.run_producer()` (script) |
| inputs | Medinet + `config/runtime_2026/` | Medinet + `GENERACION DATOS REMASEP.xlsx` + `artifacts/` |
| alcance | `processing_scope_records` (válidos ∩ mes) — **sin filtro ESTADO** | subconjunto de la hipótesis ESTADO (reproduce el snapshot legacy) |
| abre el workbook legacy | **no** | sí |
| CLI | `generate_remasep.py --mode production` (default) | `--mode diagnostic` |

Ambos comparten la **misma** lógica de negocio (`MetricValueProducer`,
`legacy_aggregation`, `legacy_transform`): no hay duplicación. El catálogo de
runtime es el mismo `formula_index` que antes se derivaba abriendo el workbook.

## Filtro ESTADO — sin cambios

`estado_filter_status = PENDING_FUNCTIONAL_CONFIRMATION`. Producción **no**
aplica ningún filtro por ESTADO; el modo diagnóstico usa la hipótesis conocida
sólo para reproducir el legacy y **no** la convierte en regla.

## Validación de assets en runtime

`load_runtime_bundle()` verifica: `bundle.yaml` presente; sha256 de cada asset;
esquema del manifiesto; `instruction_id` únicos; celdas destino únicas; cantidad
== `expected_instruction_count` (1122); un único `template_fingerprint_id`
(`stf:dc624775927d4d4d`); el catálogo cubre cada `source_metric_id` + su cierre
transitivo; `zero_write_policy ∈ {WRITE_ZERO, FORM_SPECIFIC}`.

Un asset ausente/corrupto produce `RuntimeAssetError` con
`code ∈ {RUNTIME_ASSET_MISSING, RUNTIME_ASSET_INCOMPATIBLE, RUNTIME_ASSET_INVALID}`
y un mensaje legible para el usuario ("La instalación no contiene los archivos
internos necesarios para generar el informe."). Nunca un `FileNotFoundError`
crudo. El detalle técnico va aparte.

## Packaging (PyInstaller — a futuro)

`resolve_runtime_root()` localiza `config/runtime_2026/` sin depender del
*current working directory*: primero `sys._MEIPASS` (bundle PyInstaller), luego
subiendo desde el módulo hasta encontrar `config/runtime_2026/bundle.yaml`.
Cuando se empaquete el `.exe` habrá que incluir el directorio:

    pyinstaller ... --add-data "config/runtime_2026:config/runtime_2026"

## Regenerar los assets (dev)

    python scripts/build_runtime_assets.py \
        --generator "data/local/GENERACION DATOS REMASEP.xlsx" \
        --manifest artifacts/writable_target_mapping/write_manifest_ready.csv

Congela `config/runtime_2026/` y auto-verifica que reproduce **exactamente** los
1122 PendingWrites del path validado (OLD vs NEW). La salida se versiona; el
workbook legacy no.
