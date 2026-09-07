# Estado del proyecto

Vista de estado post Sprint 3.4 (capa semántica: inventario + mapeo de
dimensiones + readiness técnica; delta del REMASEP final + provenance de fuente).
Para el detalle por sprint ver [`docs/sprints/`](sprints/README.md).

## Fases

| Fase | Contenido | Estado |
| --- | --- | --- |
| **1 — Reverse Engineering** | Inventario del workbook, dependencias, lógica legacy, plantilla oficial, UI shell | **COMPLETE** |
| **2 — Medinet & Legacy Equivalence** | Ingesta Medinet real, `AC:AL`, agregaciones directas, cierre por dependencias, `SUM`/`IF` downstream | **COMPLETE** |
| **3 — Semantic Metrics** | Inventario semántico (3.1 ✅), mapeo de dimensiones sexo/edad/alcance/código (3.2 ✅), readiness técnica + cola de revisión (3.3 ✅), delta del REMASEP final + provenance de fuente (3.4 ✅) | **IN PROGRESS** |
| **4 — Additional Sources / Complete Dataset** | Adaptador de egresos, cálculo de recursos, dataset golden | PENDING |
| **5 — Official Template Mapping** | Mapping semántico generador/Medinet → plantilla MINSAL | PENDING |
| **6 — Excel Generation / CONTROL** | Escritura vía Excel COM (Windows), recálculo, verificación `CONTROL` | PENDING |
| **7 — Pilot / Packaging** | Piloto manual + automático en paralelo, empaquetado `.exe` | PENDING |

## Completado

- Inventario del workbook generador (Sprint 1.1)
- Análisis de dependencias de fórmulas (Sprint 1.2)
- Inventario de la lógica legacy actual (Sprint 1.3)
- Inventario y alineación de la plantilla oficial (Sprint 1.4)
- UI shell navegable con datos mock (Sprint 1 UI)
- Ingesta real de Medinet + clasificación legacy (Sprint 2.1)
- Detección de filas estructuralmente vacías (Sprint 2.1)
- Equivalencia `AC:AL` (Sprint 2.2)
- Equivalencia de agregaciones directas `COUNTIF`/`COUNTIFS` (Sprint 2.3)
- Cierre por dependencias — agregaciones derivadas (Sprint 2.4)
- Equivalencia downstream `SUM` / `IF` (Sprint 2.5)
- **Cierre completo de la lógica de fórmulas legacy derivada de Medinet:
  2043 / 2043 celdas evaluables**
- **Inventario semántico preliminar de las 1771 métricas (Sprint 3.1)**:
  rótulos de sección / fila / columna (raw + normalizado), `semantic_signature`
  candidata, evidencia de fórmula (cotas de edad, criterios de texto),
  consistencia rótulo ↔ fórmula, `source = "MEDINET"`. Ver
  [`docs/SEMANTIC_METRIC_INVENTORY.md`](SEMANTIC_METRIC_INVENTORY.md).

## En curso — Sprint 3

- **3.1 — Semantic Metric Inventory: ✅ COMPLETE.** 1771 métricas candidatas,
  1741 `COMPLETE` / 30 `PARTIAL` / 0 `AMBIGUOUS` / 0 `NO_CONTEXT`, 0 firmas
  duplicadas. Ver
  [`docs/SEMANTIC_METRIC_INVENTORY.md`](SEMANTIC_METRIC_INVENTORY.md).
- **3.2 — Semantic Metric Mapping: ✅ COMPLETE.** 1771 `SemanticMetric` con
  dimensiones explícitas (sexo / edad / alcance de agregación / código de
  procedimiento), cada una con `value` + `status` + `evidence`. Resultado real:
  1414 `mapping_status=CONFIRMED` · 352 `PARTIAL` · 5 `CONFLICT` (reales, del
  workbook: columna `Mujeres` con fórmula `"*Hombre*"` en `REMASEP 01`).
  597 códigos de procedimiento explícitos (19 distintos). Vocabulario versionado
  en `config/semantic_mapping_2026/`
  (`status: preliminary_pending_functional_validation`). Ver
  [`docs/SEMANTIC_METRIC_MAPPING.md`](SEMANTIC_METRIC_MAPPING.md).
- **3.3 — Semantic Mapping Validation & Readiness: ✅ COMPLETE.** Capa técnica de
  *readiness* sobre `SemanticMetric` (no cambia las inferencias de 3.2). Policy
  versionada en `config/semantic_validation_2026/`
  (`status: preliminary_pending_functional_validation`). Resultado real sobre las
  1771 métricas: **1401 `AUTO_READY` · 50 `REVIEW_REQUIRED` · 5
  `BLOCKED_CONFLICT` · 315 `NOT_APPLICABLE`** (roll-ups del workbook).
  Issues: 5 `BLOCKING` / 69 `REVIEW` / 1360 `INFO`, 13 clusters. Los 5 conflictos
  reales (`REMASEP 01!AB84/AB86/AB87/AB89/AB90`) quedan `BLOCKED_CONFLICT` con
  `recommended_action = HUMAN_REVIEW` — **no se corrigen**, sólo se clasifican.
  `overrides.yaml` vacío (arquitectura, sin resolución). `AUTO_READY` ≠ validado
  MINSAL. Ver
  [`docs/SEMANTIC_MAPPING_VALIDATION.md`](SEMANTIC_MAPPING_VALIDATION.md).
- **3.4 — Final REMASEP Delta & Source Provenance: ✅ COMPLETE.** Comparación
  reproducible y sólo-lectura del `2026-7 REMASEP_V1.4.xlsm` (incompleto,
  `INCOMPLETE_PRE_SURGERY`) contra `REMASEP_V1.4 Julio 2026.xlsm` (terminado
  según el cliente, `FINAL_PER_CLIENT` · `approval_status = UNCONFIRMED`).
  Resultado real: **0 cambios de expresión de fórmula, 113 celdas de valor
  cambiadas** (52 input directo + 61 propagación de fórmula + 0 expresión
  reescrita; por hoja REMASEP 01 = 58 · REMASEP B1 = 41 · B2 ANEXO = 10 ·
  CONTROL = 4). `change_kind` se decide con la fórmula **antes y después**
  (`DIRECT_INPUT_CHANGE` / `FORMULA_RESULT_CHANGE` / `FORMULA_EXPRESSION_CHANGE`),
  no sólo con la de la versión final. Provenance por celda
  (`EGRESOS` 51 · `SURGICAL_TABLE` 30 · `UNKNOWN_PENDING` 17 · `RESOURCE_
  CALCULATION` 10 · `CONTROL_METADATA` 4 · `MIXED_DERIVED` 1). `EGRESOS`
  (edad/sexo → B1, códigos Qx → B2 ANEXO) confirmado por el cliente; Sección D de
  REMASEP 01 inferida por encabezado (recursos vs. tabla quirúrgica, revisión
  humana); Sección E (suspensiones) queda `UNKNOWN_PENDING`. `CONTROL = 0` →
  `PASS_INTERNAL_VALIDATION` (≠ aprobación MINSAL). 32-vs-31 registrado como
  `OPEN_FUNCTIONAL_QUESTION`. `candidate_golden_status =
  CANDIDATE_PENDING_SOURCE_COMPLETENESS`. **No** se automatiza ninguna fuente
  nueva. Ver
  [`docs/FINAL_REMASEP_DELTA_PROVENANCE.md`](FINAL_REMASEP_DELTA_PROVENANCE.md).
- **3.5+ — dimensiones de actividad / especialidad** y traducción a la tabla
  larga semántica: **pendiente**.
- **Modelo de *provenance* / `source`**: `MEDINET` se aplica en 3.1–3.3; el delta
  del REMASEP final (3.4) clasifica `EGRESOS` / `RESOURCE_CALCULATION` /
  `SURGICAL_TABLE` / `CONTROL_METADATA` / `UNKNOWN_PENDING` / `MIXED_DERIVED`.
  Todavía **sin** adaptador de egresos ni lector de tabla quirúrgica.

## Pre-Sprint 3 — inventario de fuentes / golden dataset (pendiente en paralelo)

Antes de avanzar en la generación oficial hay que completar el inventario de las
fuentes reales:

- inspeccionar el **REMASEP julio 2026 terminado** que entregará el cliente
  (candidato a golden reference — ver
  [`docs/CLIENT_PENDING.md`](CLIENT_PENDING.md));
- inspeccionar un ejemplo real de **egresos hospitalarios**, si está disponible;
- inspeccionar la **tabla quirúrgica** de julio 2026, si está disponible;
- obtener el **export Medinet original** de julio 2026;
- definir el **criterio de aceptación** del dataset golden.

## Luego

- **Sprint 3.4+ — dimensiones de actividad / especialidad**: traducir
  `semantic_signature` + `row_path` a `formulario · categoría · especialidad ·
  sexo · edad · modalidad · valor`,
  usando la evidencia de fórmula de 3.1.
- Semantic mapping hacia la plantilla oficial.
- Adaptador de egresos + cálculo de recursos.
- Escritura vía Excel COM + verificación `CONTROL`.
- Piloto en paralelo + empaquetado.

## Pendiente (backlog de fase)

- Mapping semántico de métricas.
- **Modelo de *provenance* / `source`** (ver
  [`docs/TECHNICAL_OVERVIEW.md`](TECHNICAL_OVERVIEW.md) → *Modelo futuro de
  métricas*): cada métrica deberá conservar de qué fuente proviene
  (`MEDINET` / `EGRESOS` / `RESOURCE_CALCULATION`). Requisito de diseño, sin
  código todavía.
- Adaptador de egresos hospitalarios.
- Cálculo de recursos / utilización de pabellones.
- Mapping semántico hacia la plantilla oficial MINSAL.
- Salida vía Excel COM (Windows).
- Validación `CONTROL`.
- Validación contra el dataset golden.
- Piloto manual/automático en paralelo.
- Empaquetado `.exe`.

## Lo que el sistema todavía **no** hace

- No genera el REMASEP oficial (no hay mapping ni escritura Excel).
- No integra egresos ni recursos/pabellones.
- No valida las reglas legacy contra reglas oficiales MINSAL.
- La equivalencia demostrada es **numérica contra el workbook generador**, con la
  cache de `AF` no verificable — no es una validación funcional del REMASEP.
