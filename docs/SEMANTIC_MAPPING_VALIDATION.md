# Validación / readiness del mapeo semántico (Sprint 3.3)

`scripts/validate_semantic_mapping.py` +
`src/remasep/services/semantic_validation.py` +
`config/semantic_validation_2026/`.

Sprint 3.2 produjo `SemanticMetric` con dimensiones explícitas (`sex`, `age`,
`aggregation_scope`, `procedure`) y `mapping_status` ∈
`CONFIRMED / PARTIAL / CONFLICT`. Sprint 3.3 **no cambia esas inferencias**:
añade una capa técnica y auditable que responde, por métrica —

> ¿puede usarse automáticamente? · ¿necesita revisión? · ¿debe bloquearse?

> **`AUTO_READY` significa únicamente "suficiente evidencia técnica según la
> policy actual"** (`config/semantic_validation_2026/policy.yaml`,
> `status: preliminary_pending_functional_validation`). **No** significa
> "validado MINSAL" ni "aprobado por Fundación Gantz". La validación funcional
> del REMASEP sigue pendiente.

---

## `mapping_status` (3.2) vs. `readiness_status` (3.3)

Son **ejes distintos**. El primero describe *qué tan resuelta está la semántica*;
el segundo, *qué se puede hacer con la métrica ahora*.

| | Valores | Pregunta que responde |
| --- | --- | --- |
| `mapping_status` (3.2) | `CONFIRMED` · `PARTIAL` · `CONFLICT` | ¿las dimensiones de esta métrica están resueltas por evidencia? |
| `readiness_status` (3.3) | `AUTO_READY` · `REVIEW_REQUIRED` · `BLOCKED_CONFLICT` · `NOT_APPLICABLE` | ¿esta métrica puede usarse automáticamente, necesita ojo humano o debe bloquearse? |

Un `CONFIRMED` **no** implica `AUTO_READY`: la readiness exige que las
dimensiones **requeridas para esa hoja y ese alcance** estén resueltas, no que
*todas* lo estén.

### `readiness_status`

| Estado | Cuándo | Precedencia |
| --- | --- | --- |
| `BLOCKED_CONFLICT` | alguna dimensión tiene `CONFLICT` (rótulo ↔ fórmula se contradicen) | 1 (gana siempre) |
| `NOT_APPLICABLE` | la métrica es un **roll-up** que el propio workbook calcula (`aggregation_scope` ∈ `TOTAL` / `SUBTOTAL`): no es un punto de ingreso desde Medinet | 2 |
| `REVIEW_REQUIRED` | una **dimensión requerida** para esa hoja/alcance no está resuelta (`LABEL_ONLY` / `FORMULA_ONLY` / `UNRESOLVED` / ausente) | 3 |
| `AUTO_READY` | sin conflicto, sin roll-up, y todas las dimensiones requeridas en `CONFIRMED` o `NOT_APPLICABLE` | 4 |

Un `CONFLICT` **nunca** queda absorbido por un `NOT_APPLICABLE` ni por un
`mapping_status = CONFIRMED`: la precedencia de `BLOCKED_CONFLICT` es absoluta.

---

## Policy versionada

`config/semantic_validation_2026/policy.yaml` — reglas **técnicas observables
sobre el workbook de referencia**, no requisitos clínicos.

- **`rollup_scopes: [TOTAL, SUBTOTAL]`** — alcances que son roll-ups del workbook
  ⇒ `NOT_APPLICABLE` (salvo `CONFLICT`).
- **`required_dimensions`** — por formulario y bucket (`DETAIL` = celda de dato,
  `ROLLUP` = total/subtotal). Justificación observable: la grilla de `REMASEP_OD`
  y `REMASEP 01` desagrega por **sexo y edad** (columnas con esos rótulos);
  `B2 ANEXO` es una lista plana de prestaciones **sin esos ejes** pero **con
  código de prestación**.

  | Formulario | `DETAIL` requiere |
  | --- | --- |
  | `REMASEP_OD` | `form`, `row_path`, `aggregation_scope`, `sex`, `age` |
  | `REMASEP_01` | `form`, `row_path`, `aggregation_scope`, `sex`, `age` |
  | `B2_ANEXO` | `form`, `row_path`, `aggregation_scope`, `procedure` |
  | *default* | `form`, `row_path`, `aggregation_scope` |

- **`acceptable_status`** — estados de una dimensión requerida que **no** bloquean
  `AUTO_READY`: `CONFIRMED` y `NOT_APPLICABLE` (para `procedure`, `EXPLICIT` /
  `NOT_APPLICABLE`). Una dimensión requerida con `NOT_APPLICABLE` es aceptable (el
  eje no aplica a esa métrica concreta); con `LABEL_ONLY` / `FORMULA_ONLY` /
  `UNRESOLVED` genera un issue de severidad `REVIEW`.

Cambiar la política = editar el YAML y volver a correr. Sin recompilar reglas.

---

## Severidad de issue

| Severidad | Origen | Efecto en readiness |
| --- | --- | --- |
| `BLOCKING` | un `CONFLICT` de dimensión | `BLOCKED_CONFLICT` |
| `REVIEW` | una **dimensión requerida** sin resolver | `REVIEW_REQUIRED` |
| `INFO` | una dimensión **no requerida** sin resolver (p.ej. código de procedimiento en `REMASEP_OD`, donde el eje real es sexo × edad) | ninguno — sólo se registra |

`review_reasons` (por métrica) = el conjunto ordenado de `issue_type` de
severidad `BLOCKING` o `REVIEW`. Es una **lista explícita**, no un string
ambiguo: p.ej. `("AGE_UNRESOLVED", "SEX_LABEL_ONLY")`.

---

## `SemanticReviewIssue`

Un issue por dimensión problemática de una métrica:

`issue_id` · `metric_id` · `form` · `sheet` · `cell` · `dimension` ·
`issue_type` · `severity` · `label_evidence` · `formula_evidence` ·
`recommended_action` · `status`.

- Los **5 conflictos reales** del workbook
  (`REMASEP 01!AB84, AB86, AB87, AB89, AB90`) generan issues con
  `issue_type = SEX_LABEL_FORMULA_MISMATCH`, `severity = BLOCKING`,
  `label_evidence = FEMALE`, `formula_evidence = MALE`,
  **`recommended_action = HUMAN_REVIEW`**, `status = OPEN`.
- **Nunca** se propone "cambiar a MALE" ni "cambiar a FEMALE". El sistema
  **registra la evidencia de ambos lados y no toma partido.** La discrepancia es
  del workbook legacy y la resuelve una persona.

---

## Cola de revisión y clusters

- **`review_queue.csv`** — una fila por issue accionable (`BLOCKING` + `REVIEW`),
  ordenada con los `BLOCKING` primero. Todas con
  `recommended_action = HUMAN_REVIEW`.
- **`review_reason_summary.csv`** — conteo por `issue_type` × `dimension` ×
  `severity`. Convierte "N métricas con problema" en "N métricas agrupadas por
  **causa**".
- **`review_clusters.csv`** — problemas **equivalentes** agrupados por
  `(form, dimension, issue_type, severity)`, con `metric_count` y unos pocos
  `example_metric_ids`. Un `PARTIAL` que afecta a cientos de celdas **no** son
  cientos de incidencias distintas: es **un** patrón (p.ej. "B2 ANEXO no expone
  código de prestación en 19 filas") que se decide una vez.

---

## Overrides — arquitectura, **sin resolución** en 3.3

`config/semantic_validation_2026/overrides.yaml` está **vacío por diseño**.

En el futuro, una revisión humana podrá resolver de forma **auditable** una
excepción (p.ej. un `BLOCKED_CONFLICT`) **sin tocar** el workbook, la fórmula, el
label ni el motor legacy. Formato previsto de cada entrada:

```yaml
- metric_signature: "REMASEP 01 :: ... :: MUJERES"   # identidad (obligatorio)
  cell: "REMASEP 01!AB84"                             # sólo trazabilidad
  decision: REVIEW_PENDING | ACCEPT_AS_IS | ...
  reason: "..."
  evidence: "..."
  approved_by: "..."
  approved_date: "YYYY-MM-DD"
```

- La **identidad** de un override es `metric_signature` (independiente de
  coordenadas); `cell` se conserva sólo como traza.
- En 3.3 el pipeline **sólo carga el fichero y marca `override_matched`** por
  firma. **No** cambia ningún `readiness_status`.
- Los **5 conflictos reales NO están** en el fichero y **no se inventan
  aprobaciones**: quedan `BLOCKED_CONFLICT` hasta que exista una decisión humana
  explícita.

---

## Rastro de auditoría humana

Todo lo que una persona necesita para revisar, sin abrir el workbook ni ver una
sola fila Medinet:

| Necesidad | Artefacto |
| --- | --- |
| "¿qué está listo / qué no?" | `semantic_metric_readiness.csv`, `readiness_coverage.csv` |
| "¿qué tengo que mirar?" | `review_queue.csv` (accionables), `conflicts.csv` (los 5 bloqueos) |
| "¿son problemas distintos o el mismo repetido?" | `review_reason_summary.csv`, `review_clusters.csv` |
| "revíseme una muestra representativa" | `readiness_manual_review_sample.csv` (determinista, estratificada) |
| "¿qué decidió un humano y quién?" | `overrides.yaml` (hoy vacío) |
| "¿con qué reglas se evaluó?" | `policy.yaml` (versionada) + `summary.json` → `validation_policy_version` |

`summary.json` lleva `validation_status: PRELIMINARY_NOT_VALIDATED` de forma
explícita.

---

## Privacidad

Los artefactos contienen sólo `metric_id`, coordenadas, rótulos del formulario y
estados agregados. **Ninguna fila Medinet individual**, ningún dato de paciente.

---

## Lo que Sprint 3.3 **no** hace

- No corrige los 5 conflictos (fórmula, label, mapping y motor legacy quedan
  intactos — sólo se **clasifican**).
- No es un mapping a la plantilla oficial MINSAL.
- No aplica overrides (sólo prepara la arquitectura).
- No toca EGRESOS, cálculo de recursos, COM ni UI.
- No introduce reglas clínicas nuevas ni codificación de actividad / especialidad.
