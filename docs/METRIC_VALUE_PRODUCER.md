# Metric value producer (Sprint 3.7A)

`scripts/build_metric_values.py` +
`src/remasep/services/metric_value_producer.py` +
`config/metric_value_producer_2026/`.

Sprint 3.6 dejó **1122 `WriteInstruction` `WRITE_READY`**: sabe **dónde**
escribir. Sprint 3.7A responde **qué valor** alimenta cada instrucción, y
compara ese valor —celda a celda— contra la referencia legacy y la plantilla
final oficial.

> **No se escribe Excel.** No se usa COM, no se modifica la plantilla, no se
> conecta la UI. Se verifica el SHA256 de todas las fuentes antes y después. Los
> artefactos contienen coordenadas y **un conteo agregado** por métrica: ninguna
> fila de paciente, ningún dato personal.

## Motor reutilizado (una sola implementación)

El productor **no** reimplementa ninguna fórmula. Reutiliza:

| pieza | módulo | qué aporta |
| --- | --- | --- |
| parsing/evaluación `COUNTIF(S)(...) ± celda` | `remasep.services.legacy_aggregation` | `parse_derived_formula(...).evaluate(rows, lookup)` |
| columnas derivadas `AC:AL` | `remasep.services.legacy_transform` | `legacy_derived_record(rec, ruleset)` |
| ruleset de clasificación legacy | `remasep.services.legacy_rules` | `load_legacy_rules()` |
| grafo de dependencias + fórmula por celda | `scripts/close_legacy_aggregations.py` | `ClosureResult.nodes[(sheet, cell)]` con `.formula`, `.kind`, `.value_ref_coords`, `.python_value` |
| alcance del período | `remasep.services.medinet_analysis` | `processing_scope_frame(path, period)` |

`close_legacy_aggregations` ahora expone `ClosureResult.detail_columns_map`
(campo semántico → letra de columna en la hoja de detalle legacy): las fórmulas
`COUNTIF` referencian columnas por letra (`!$AA:$AA`), así que reevaluarlas sobre
otro export exige alinear las filas a **esas** letras.

`medinet_analysis` factoriza `_scope_selection(...)`: `analyze()` y
`processing_scope_frame()` comparten exactamente la definición de
`processing_scope_records` (válidos ∩ mes/año, **sin** filtro por ESTADO).

## `MetricValue`

```
MetricValue(source_metric_id, value, period, producer_version)
```

- `source_metric_id` — `LEGACY::<hoja>::<celda>` (idéntico al `Node.metric_id`
  del cierre y al `source_metric_id` del manifiesto).
- `value` — conteo entero ≥ 0 (`INTEGER_COUNT`). Sin truncamiento: un `float`
  exacto (`5.0`) se coerciona a `int`; un `float` no entero, un `str`, un `bool`
  o un negativo → `VALUE_TYPE_CONFLICT` (la métrica queda sin `MetricValue`).
- `period` — etiqueta del período (`"Julio 2026"`), preservada en cada valor.
- `producer_version` — `medinet_value_producer_2026.v1`.

`produce(source_metric_id)` evalúa una métrica; `produce_run(ids, expected_types)`
evalúa un lote y separa `metric_values` de `value_type_conflicts` y de
`unsupported` (fórmula fuera del subset). Las derivadas (`COUNTIFS ± celda`) se
resuelven por recursión memoizada sobre `value_ref_coords`.

## Dos modos (mutuamente excluyentes)

### `PRODUCTION_PERIOD_SCOPE` (productivo, por defecto)

- registros = `processing_scope_records` del export directo `detalle_citas`
  (válidos ∩ mes). **Julio 2026 = 2114.**
- **no** aplica ningún filtro por ESTADO.
- reporta `input_scope = PERIOD_ONLY`, `estado_filter_applied = false`.
- **no** se compara contra el REMASEP final legacy como si debieran coincidir: el
  filtro por ESTADO no está confirmado (`estado_filter_status =
  PENDING_FUNCTIONAL_CONFIRMATION`).

### `LEGACY_EQUIVALENCE_DIAGNOSTIC` (diagnóstico, opt-in explícito con `--diagnostic`)

- **fuera del flujo productivo.** No cambia ninguna configuración productiva.
- registros = `processing_scope_records` **filtrados por la hipótesis ESTADO
  observada** (`LEGACY_KEPT_STATES`). Julio 2026 → **1364**, exactamente los
  registros activos del detalle legacy.
- sólo sirve para comparar `MetricValue` contra el snapshot legacy de julio.
- **no** convierte la hipótesis en regla.

## `PendingWrite` — unión con el manifiesto

`join_metric_values_with_manifest(metric_values, write_instructions)` →
`PendingWrite(instruction_id, source_metric_id, target_sheet, target_cell,
value, expected_value_type, period)`. **Todavía no se escribe.**

- un `source_metric_id` con más de un `MetricValue` → **error** (no
  *last-one-wins*: la ambigüedad de valor se resuelve, no se elige el último).

`check_write_completeness(...)` → `CompletenessReport(missing, duplicate,
orphan)`: cada instrucción `WRITE_READY` debe tener **exactamente un**
`MetricValue`.

## Contrato de período

`processing_scope_records` = registros estructuralmente válidos ∩ mes/año
seleccionado. Nunca "todos los válidos", nunca con filtro por ESTADO. Idéntico al
contrato del Sprint 3.6.

## Equivalencia de valores de referencia (§7 / §8)

Para las 1122 `WriteInstruction`:

- `source_reference_value` — valor legacy **reevaluado** por el motor de
  agregación sobre el detalle legacy de `GENERACION DATOS REMASEP.xlsx`
  (1364 filas activas).
- `source_reference_value_cached` — lo que ese workbook dejó **en caché**.
- `target_reference_value` — celda final en `REMASEP_V1.4 Julio 2026.xlsm`.

`comparison_status` ∈ `MATCH` · `ZERO_VS_BLANK_EQUIVALENT` · `MISMATCH` ·
`TARGET_UNAVAILABLE` · `SOURCE_UNAVAILABLE`. **`0` y blanco no se tratan como
iguales automáticamente**: `ZERO_VS_BLANK_EQUIVALENT` es un estado explícito y
separado, y la política del cero se deriva de la evidencia.

## Semántica del cero (§9 / §10)

`zero_semantics` clasifica cada caso `source == 0`:
`SOURCE_ZERO_TARGET_ZERO` · `SOURCE_ZERO_TARGET_BLANK` · `SOURCE_ZERO_TARGET_OTHER`.

`config/metric_value_producer_2026/zero_write_policy.yaml` declara `resolution`
∈ `WRITE_ZERO` · `PRESERVE_BLANK_FOR_ZERO` · `FORM_SPECIFIC` · `UNRESOLVED`.
El script **no sobrescribe** ese archivo: recalcula la política desde la
evidencia y avisa si difiere (`zero_write_policy_config.config_matches_evidence`).
`UNRESOLVED` es un resultado válido.

## Investigación de mismatches (§12)

No se autocorrigen. `value_mismatches.csv` + `value_mismatch_clusters.csv`
agrupan por forma / kind / estado de comparación / semántica del cero /
`investigation_axis`:

- `LEGACY_CACHE_MATCHES_TARGET` — la celda final coincide con el valor **cacheado**
  por el workbook legacy pero no con el reevaluado (artefacto de caché legacy,
  típicamente la columna `AF`/edad; ver `docs/TECHNICAL_OVERVIEW.md §8`).
- `LEGACY_CACHE_SUSPECT` — el nodo tiene `CACHE_DIFFERENCE` en el cierre.
- `ZERO_SEMANTICS` — `source == 0` con destino distinto de 0.
- `SOURCE_POPULATION_OR_TARGET_ADJUSTMENT` — no discriminable sin validación
  funcional.

## Filtro por ESTADO — pendiente

`estado_filter_status = PENDING_FUNCTIONAL_CONFIRMATION`. El modo productivo
**no** lo aplica. El modo diagnóstico lo usa sólo para el estudio de equivalencia
y **no** lo promueve a regla. Sin resolver: los 245 `STRONG`, los 34 `AMBIGUOUS`,
los 5 conflictos (siguen igual que en 3.6).

## Artefactos (`artifacts/metric_value_producer/`)

`production_metric_values.csv`, `diagnostic_legacy_metric_values.csv`,
`reference_value_equivalence.csv`, `value_equivalence_summary.csv`,
`zero_semantics_analysis.csv`, `value_mismatches.csv`,
`value_mismatch_clusters.csv`, `pending_writes_reference.csv`,
`pending_writes_production.csv`, `summary.json`, `README.md`.

## Corrida real — Julio 2026

| bloque | resultado |
| --- | --- |
| Producción — alcance | 2114 registros, `input_scope = PERIOD_ONLY`, `estado_filter_applied = false` |
| Producción — `MetricValue` | 1122 / 1122; 0 conflictos de tipo; 0 sin soporte |
| Producción — completitud | missing 0 · duplicate 0 · orphan 0 |
| Equivalencia ref (vs valor reevaluado) | MATCH 980 · ZERO_VS_BLANK 0 · MISMATCH 142 · UNAVAILABLE 0 |
| Equivalencia ref (vs valor **cacheado** legacy) | MATCH 1122 / 1122 |
| Semántica del cero | 935 casos `source == 0`: 917 target 0, **0 target blanco**, 18 target otro |
| `zero_write_policy` | **`WRITE_ZERO`** (evidencia concluyente; la plantilla base tampoco deja blancos) |
| Mismatches | 142, todos `LEGACY_CACHE_MATCHES_TARGET` (5 clusters) — artefacto de caché legacy, no error del productor |
| Diagnóstico | alcance 1364 (= detalle legacy); 1122 / 1122 `MetricValue` **idénticos** al valor legacy reevaluado (`EXACT_EQUIVALENCE`) |

Interpretación: el productor reproduce **exactamente** la lógica legacy
(diagnóstico 1:1). Contra la plantilla final oficial, coincide al 100 % con los
valores que el Excel legacy tenía cacheados; los 142 MISMATCH frente al valor
reevaluado son la caché obsoleta de `AF` ya documentada, no una discrepancia de
producción. La representación del cero en la plantilla es siempre `0` explícito.
