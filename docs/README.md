# Documentación — REMASEP Automation

Índice navegable de toda la documentación del proyecto. Empieza por el
[`README.md`](../README.md) de la raíz para el contexto y el estado general.

---

## Getting Started

- [`../README.md`](../README.md) — objetivo, estado por fases, capacidades
  demostradas, cómo ejecutar.
- [`TECHNICAL_OVERVIEW.md`](TECHNICAL_OVERVIEW.md) → *Stack* y *Fronteras* —
  entorno de desarrollo (venv, `pytest`, `ruff`, UI).

## Architecture

- [`ARCHITECTURE.md`](ARCHITECTURE.md) — componentes Extract → Transform →
  Metrics → Map → Load.
- [`TECHNICAL_OVERVIEW.md`](TECHNICAL_OVERVIEW.md) — stack, fronteras, privacidad,
  semánticas de texto, motor legacy, DAG, limitaciones de cache.

## Data Sources

- [`MEDINET_INPUT_CONTRACT.md`](MEDINET_INPUT_CONTRACT.md) — **input de producción**
  = export directo Medinet "Detalle de citas"; `GENERACION DATOS REMASEP.xlsx` es
  sólo referencia legacy. Alcance de período (`processing_scope_records`) y
  reconciliación privacy-safe 2114 vs 1364 (hipótesis de ESTADO, no regla).
- [`TECHNICAL_OVERVIEW.md`](TECHNICAL_OVERVIEW.md) → *Fuentes de datos* —
  MEDINET · EGRESOS · TABLA QUIRÚRGICA / RESOURCE CALCULATION · OFFICIAL TEMPLATE.
- [`TECHNICAL_OVERVIEW.md`](TECHNICAL_OVERVIEW.md) → *Modelo futuro de métricas* —
  requisito de *provenance* (`source`).

## Workbook Reverse Engineering (Fase 1)

- [`CURRENT_EXCEL_FLOW.md`](CURRENT_EXCEL_FLOW.md) — grafo de dependencias entre
  hojas del workbook generador (observado).
- [`CURRENT_LOGIC.md`](CURRENT_LOGIC.md) — columnas `AC:AL`, tipos de
  transformación, 27 reglas candidatas.

## Legacy Equivalence (Fase 2)

- [`MEDINET_ANALYSIS.md`](MEDINET_ANALYSIS.md) — ingesta real de Medinet,
  validación, filas estructuralmente vacías, buckets legacy.
- [`LEGACY_EQUIVALENCE.md`](LEGACY_EQUIVALENCE.md) — equivalencia `RAW → AC:AL`.
- [`LEGACY_AGGREGATION_EQUIVALENCE.md`](LEGACY_AGGREGATION_EQUIVALENCE.md) —
  agregaciones directas `COUNTIF`/`COUNTIFS`.
- [`LEGACY_AGGREGATION_CLOSURE.md`](LEGACY_AGGREGATION_CLOSURE.md) — cierre por
  dependencias (agregaciones derivadas + DAG).
- [`LEGACY_DOWNSTREAM_EQUIVALENCE.md`](LEGACY_DOWNSTREAM_EQUIVALENCE.md) —
  `SUM` downstream + `IF` validation; cierre completo (2043 / 2043).

## Semantic Layer (Fase 3, en curso)

- [`SEMANTIC_METRIC_INVENTORY.md`](SEMANTIC_METRIC_INVENTORY.md) — Sprint 3.1:
  rótulos de sección / fila / columna por métrica, `semantic_signature`
  candidata, evidencia de fórmula, `source = "MEDINET"`.
- [`SEMANTIC_METRIC_MAPPING.md`](SEMANTIC_METRIC_MAPPING.md) — Sprint 3.2:
  dimensiones explícitas (sexo, edad, alcance de agregación, código de
  procedimiento) con `value` + `status` + `evidence`; vocabulario versionado en
  `config/semantic_mapping_2026/`. *No es un mapping a la plantilla oficial ni
  validación funcional.*
- [`SEMANTIC_MAPPING_VALIDATION.md`](SEMANTIC_MAPPING_VALIDATION.md) — Sprint 3.3:
  capa técnica de *readiness* sobre `SemanticMetric` — `AUTO_READY` /
  `REVIEW_REQUIRED` / `BLOCKED_CONFLICT` / `NOT_APPLICABLE`; policy versionada en
  `config/semantic_validation_2026/`; cola de revisión y clusters; los 5
  conflictos reales quedan `BLOCKED_CONFLICT` con `recommended_action =
  HUMAN_REVIEW`. *`AUTO_READY` ≠ validado MINSAL.*
- [`FINAL_REMASEP_DELTA_PROVENANCE.md`](FINAL_REMASEP_DELTA_PROVENANCE.md) —
  Sprint 3.4: delta sólo-lectura entre el REMASEP julio incompleto y el final
  (0 cambios de fórmula, 113 celdas de valor: 52 directas + 61 propagadas) y
  *provenance* de fuente por celda (`EGRESOS` · `RESOURCE_CALCULATION` ·
  `SURGICAL_TABLE` · `CONTROL_METADATA` · `UNKNOWN_PENDING` · `MIXED_DERIVED`).
  32-vs-31 = `OPEN_FUNCTIONAL_QUESTION`; `candidate_golden_status =
  CANDIDATE_PENDING_SOURCE_COMPLETENESS`. *No automatiza ninguna fuente nueva.*
- [`OFFICIAL_TEMPLATE_SEMANTIC_ALIGNMENT.md`](OFFICIAL_TEMPLATE_SEMANTIC_ALIGNMENT.md)
  — Sprint 3.5: empareja las métricas MEDINET elegibles (`AUTO_READY`) con celdas
  de la plantilla oficial por **evidencia por dimensión**, no por coordenada
  (1122 `EXACT` / 245 `STRONG` / 34 `AMBIGUOUS` de 1401 elegibles). Regiones
  NO-MEDINET (EGRESOS / recursos / tabla quirúrgica) excluidas; `REMASEP B1`
  entero es EGRESOS. Reglas en `config/official_template_alignment_2026/`.
  *No escribe Excel; un match ≠ validado MINSAL.*

## Official Template (Fase 5, en curso)

- [`OFFICIAL_TEMPLATE_SEMANTIC_ALIGNMENT.md`](OFFICIAL_TEMPLATE_SEMANTIC_ALIGNMENT.md)
  — Sprint 3.5: alineación semántica de candidatos generador → plantilla oficial
  (sin escritura Excel).
- [`WRITABLE_TARGET_MAPPING.md`](WRITABLE_TARGET_MAPPING.md) — Sprint 3.6:
  `write_status` (`WRITE_READY` / `WRITE_REVIEW_REQUIRED` / `WRITE_BLOCKED` /
  `NOT_WRITABLE`) y manifiesto de instrucciones de escritura versionable
  (1122 `WRITE_READY`, todos `EXACT`). `instruction_id` semántico, no coordenada;
  `structural_template_fingerprint` ≠ `file_sha256`. **No escribe Excel.**
- [`METRIC_VALUE_PRODUCER.md`](METRIC_VALUE_PRODUCER.md) — Sprint 3.7A: **qué
  valor** alimenta cada `WriteInstruction`. `MetricValue` / `PendingWrite`,
  modos `PRODUCTION_PERIOD_SCOPE` (2114 registros, sin filtro ESTADO) y
  `LEGACY_EQUIVALENCE_DIAGNOSTIC` (1364, opt-in). Equivalencia de valores contra
  la referencia legacy (1122/1122 vs. caché legacy; 980 MATCH / 142 MISMATCH por
  caché `AF` obsoleta vs. valor reevaluado) y `zero_write_policy = WRITE_ZERO`
  (evidence-derived). **No escribe Excel.**
- [`EXCEL_WRITER.md`](EXCEL_WRITER.md) — Sprint 3.7B: escribe los `PendingWrite`
  en una **copia** del REMASEP oficial con Microsoft Excel Desktop (COM,
  Windows). Contrato `WorkbookWriter` + `ExcelComWorkbookWriter` + `FakeWorkbookWriter`;
  `ExcelCapability` (en WSL `can_generate = False`, graceful); copy-first +
  salida atómica; preflight; verificación de integridad de fórmulas/VBA y de las
  1122 celdas; `ControlResult` (independiente de `writer_integrity_status`);
  modos `DIAGNOSTIC_REFERENCE` / `PRODUCTION`; archivo `NOT_FOR_SUBMISSION`. UI
  **no** conectada.
- [`RUNTIME_DECOUPLING.md`](RUNTIME_DECOUPLING.md) — Sprint 3.8: el path de
  producción construye los 1122 `PendingWrite` desde **sólo** el export Medinet
  + assets de runtime versionados (`config/runtime_2026/`: `write_manifest.csv`,
  `metric_catalog.csv`, `detail_contract.yaml`, `zero_policy.yaml`, `bundle.yaml`).
  **No** abre `GENERACION DATOS REMASEP.xlsx` ni lee `artifacts/`. Assets sin
  PII, validados en runtime (`RuntimeAssetError` controlado). El workbook legacy
  queda dev-only (`--mode diagnostic`).
- [`TEMPLATE_ALIGNMENT.md`](TEMPLATE_ALIGNMENT.md) — estructura de
  `REMASEP 2026_V1.4.xlsm` y candidatos de alineación por etiqueta (Sprint 1.4).

## Sprints

- [`sprints/README.md`](sprints/README.md) — tabla resumen Sprint 0 → 2.5.
- Un documento por sprint en [`sprints/`](sprints/).

## Progress / Roadmap

- [`PROGRESS.md`](PROGRESS.md) — vista de estado por fases: completado, siguiente
  (pre-Sprint 3), pendiente.
- [`ROADMAP.md`](ROADMAP.md) — plan por sprints (sin fechas).

## Client / Functional Pending Questions

- [`CLIENT_PENDING.md`](CLIENT_PENDING.md) — preguntas funcionales
  **RESUELTAS** (respuestas del cliente) y **PENDIENTES**.
