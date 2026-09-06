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
  candidata, evidencia de fórmula, `source = "MEDINET"`. *No es un mapping
  validado.*

## Official Template (Fase 5, pendiente)

- [`TEMPLATE_ALIGNMENT.md`](TEMPLATE_ALIGNMENT.md) — estructura de
  `REMASEP 2026_V1.4.xlsm` y candidatos de alineación (mapping NO implementado).

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
