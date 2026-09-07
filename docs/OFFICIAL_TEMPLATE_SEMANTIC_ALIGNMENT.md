# Alineación semántica con la plantilla oficial (Sprint 3.5)

`scripts/align_official_template.py` +
`src/remasep/services/template_alignment.py` +
`config/official_template_alignment_2026/`.

Sprints 3.1–3.3 produjeron **1771 `SemanticMetric` MEDINET** que describen *qué
mide* cada celda del workbook **generador** legacy. Sprint 3.5 intenta emparejar
cada métrica **elegible** con la(s) celda(s) de la **plantilla oficial** REMASEP
que representan lo mismo.

> **La plantilla oficial tiene otro layout que el generador** (otras coordenadas,
> otra estructura). El matching se basa en **evidencia por dimensión** —
> `form` · `section` · `row_path` · `column_path` · `sex` · `age` ·
> `aggregation_scope` · `procedure_code` · `semantic_signature` — **nunca** en la
> coordenada. Un match **no** significa "MINSAL validated": reglas versionadas en
> `config/official_template_alignment_2026/` con
> `status: preliminary_pending_functional_validation`.

Sólo lectura: no escribe ni guarda ningún workbook, no usa COM/macros, verifica
el SHA256 de ambos archivos antes y después. **No** resuelve EGRESOS ni
RESOURCE_CALCULATION, **no** toca el mapeo semántico de Sprint 3.2.

---

## Source metric vs. target metric

| | qué es | de dónde |
| --- | --- | --- |
| **source** | `SemanticMetric` MEDINET del generador (Sprint 3.2) con `readiness = AUTO_READY` (Sprint 3.3) | `GENERACION DATOS REMASEP.xlsx` |
| **target** | `TargetMetricContext` — celda de la plantilla con contexto semántico y clasificación | `REMASEP_V1.4 Julio 2026.xlsm` (referencia estructural) |

### `target_kind` · `target_alignment_role` · `target_source_expectation` — tres cosas distintas

| campo | pregunta | valores |
| --- | --- | --- |
| **`target_kind`** | ¿qué **es** físicamente la celda? | `DIRECT_INPUT_TARGET` · `FORMULA_TARGET` · `STRUCTURAL` · `VALIDATION` · `UNKNOWN` |
| **`target_alignment_role`** | ¿**para qué** sirve en la alineación? | `INPUT_TARGET` (puede recibir un mapping) · `DERIVED_TARGET` (fórmula, depende de inputs) · `STRUCTURAL` (rótulo/código, nunca alineable) · `VALIDATION` · `UNKNOWN` |
| **`target_source_expectation`** | ¿de qué **fuente** debería venir? | `MEDINET` · `EGRESOS` · `RESOURCE_CALCULATION` · `SURGICAL_TABLE` · `CONTROL_METADATA` · `UNKNOWN` |

- **`DIRECT_INPUT_TARGET` exige evidencia estructural de ser una celda de
  ingreso.** La plantilla oficial tiene todas las hojas **protegidas**; el autor
  sólo desbloquea (`protection.locked = False`) las celdas donde el usuario debe
  escribir. Ésa es la señal autoritativa. *No tener fórmula* NO basta: los
  rótulos, encabezados y **códigos de prestación** (columna A de B2 ANEXO, p.ej.
  `A1024 = 1103001`) tampoco tienen fórmula y estaban siendo mal clasificados
  como input. Una celda bloqueada, o una hoja sin proteger, → `STRUCTURAL` /
  `UNKNOWN`, nunca `INPUT_TARGET`. Los 11 inputs de B2 conocidos (Sprint 3.4:
  `C1317…C1524`; y `C833/C953/C954/C957/C958`) están todos `locked = False`.
- Una `STRUCTURAL` **nunca** cuenta como "target MEDINET faltante".
- Una `DERIVED_TARGET` (fórmula) **no** es un punto de ingreso: no se confunde con un `INPUT_TARGET`.
- **`target_source_expectation` = `UNKNOWN` por defecto.** Sólo se marca `MEDINET`
  cuando (a) una región con evidencia estructural positiva lo reclama, o (b) un
  source MEDINET se alinea `EXACT`/`STRONG` a la celda (el match sube la
  expectativa `UNKNOWN → MEDINET`). Sólo las regiones explícitamente **NO-MEDINET**
  (`EGRESOS` / recursos / tabla quirúrgica) bloquean el emparejamiento; una
  expectativa `UNKNOWN` **no** bloquea.

### Elegibilidad de la fuente

Sólo `readiness = AUTO_READY` es elegible. Se registran en `source_exclusions.csv`:

| motivo | qué |
| --- | --- |
| `BLOCKED_CONFLICT` | los 5 conflictos `REMASEP 01!AB84/AB86/AB87/AB89/AB90` — **no** entran a AUTO mapping |
| `REVIEW_REQUIRED` | dimensión requerida sin resolver (Sprint 3.3) |
| `ROLLUP_NOT_INPUT` | readiness `NOT_APPLICABLE` — verificado: los 315 son `DOWNSTREAM_TOTAL` con `aggregation_scope` `TOTAL`/`SUBTOTAL` (0 excepciones); un roll-up no es punto de ingreso |

---

## Matching por dimensión

`score_pair` produce un score **determinista y explicable** ∈ [0, 1] = suma
ponderada de dimensiones en `MATCH` sobre las dimensiones aplicables presentes en
ambos lados. `FORM` es una **compuerta** (no puntúa; si no coincide, no hay
match).

| dimensión | peso | evidencia |
| --- | ---: | --- |
| `PROCEDURE_CODE` | 5 | código exacto = evidencia dominante (B2 ANEXO). Se preserva el cero inicial. |
| `ROW_PATH` | 4 | igualdad; misma hoja identifica la fila igual aunque muestre más/menos niveles de ancestro → `MATCH`; si ambos lados llevan el **mismo código** de procedimiento en la fila → `MATCH` |
| `COLUMN_PATH` | 4 | estructura de columna sin los tokens de sexo/edad (esos van aparte); `NOT_APPLICABLE` en B2 ANEXO |
| `SEX` | 3 | `MALE` / `FEMALE` / `BOTH`; distinto en ambos → `MISMATCH` → **CONFLICT** |
| `AGE` | 3 | intervalo `(min, max)`; disjunto → `MISMATCH` → **CONFLICT**; solapado → `PARTIAL` |
| `SECTION` | 2 | rótulo de sección normalizado |
| `AGGREGATION_SCOPE` | 2 | `DETAIL` / `SUBTOTAL` / `TOTAL` |

### `match_status`

| estado | criterio |
| --- | --- |
| `EXACT_SEMANTIC_MATCH` | score ≥ `exact` (1.0), `ROW_PATH` = `MATCH`, `COLUMN_PATH`/`SEX`/`AGE` en `MATCH` o `NOT_APPLICABLE` |
| `STRONG_MATCH` | score ≥ `strong` (0.75) — evidencia suficiente pese a una diferencia estructural (p.ej. profundidad de fila distinta) |
| `AMBIGUOUS` | > 1 target dentro de `tie_margin` (0.05) del mejor score — **no** se elige uno |
| `NO_MATCH` | ningún target con evidencia suficiente (`unmatched_sources.csv`) |
| `CONFLICT` | dimensión incompatible (sexo / edad disjunta / código distinto) |

---

## `target_source_expectation` — regiones declaradas

`config/official_template_alignment_2026/source_regions.yaml` declara regiones de
la plantilla con su fuente esperada. **Orden = prioridad**: las NO-MEDINET (más
específicas, derivadas de [`docs/FINAL_REMASEP_DELTA_PROVENANCE.md`](FINAL_REMASEP_DELTA_PROVENANCE.md))
se evalúan antes que las MEDINET amplias. **Si ninguna región reclama la celda →
`UNKNOWN`** (no se infiere MEDINET por defecto).

| región | ámbito | fuente esperada |
| --- | --- | --- |
| `B1_SURGERY_GRID` / `B1_WHOLE_SHEET` | REMASEP B1 (hoja completa) | `EGRESOS` |
| `R01_SECTION_D_RESOURCES` | REMASEP 01 · filas 124–133 (Sección D) | `RESOURCE_CALCULATION` |
| `R01_SECTION_E_SUSPENSIONS` | REMASEP 01 · filas 134–146 (Sección E) | `UNKNOWN` |
| `B2_SURGICAL_INTERVENTIONS` | B2 ANEXO · filas 1023–2867 (bloque quirúrgico) | `EGRESOS` |
| `OD_MEDINET_SUBGRAPH` | REMASEP_OD (hoja completa) | `MEDINET` (evidencia positiva) |
| `R01_MEDINET_CONSULTAS` | REMASEP 01 · secciones A/B/C (controles y consultas) | `MEDINET` (evidencia positiva) |

Identificación: por uno o más **rótulos de sección** (evidencia semántica; acepta
lista) y/o, para el workbook de referencia actual, por un **rango de filas
documentado y testeado**. Cuando una región declara rango de filas, ése es
autoritativo (la sección queda como documentación).

**No se mezclan fuentes**: si el mejor target de una métrica MEDINET cae en una
región **explícitamente NO-MEDINET**, la métrica queda `unmatched` con motivo
`TARGET_SOURCE_NOT_MEDINET`. Una expectativa `UNKNOWN` **no** bloquea — un match
`EXACT`/`STRONG` sube esa celda a `MEDINET`. `REMASEP B1` entero es EGRESOS → sus
celdas nunca reciben un mapping MEDINET (resultado válido, no fallo — §12).

---

## Resultado real

Generador `GENERACION DATOS REMASEP.xlsx` (sha `fc2e1536…`) vs. plantilla
`REMASEP_V1.4 Julio 2026.xlsm` (sha `6beef1ff…`, `approval_status = UNCONFIRMED`).

```
source metrics        1771
  elegibles           1401   (AUTO_READY)
  excluidos            370   (5 BLOCKED_CONFLICT · 315 ROLLUP_NOT_INPUT · 50 REVIEW_REQUIRED)

match_status          1122 EXACT_SEMANTIC_MATCH · 245 STRONG_MATCH · 34 AMBIGUOUS · 0 NO_MATCH · 0 CONFLICT
  REMASEP_OD          952 EXACT · 170 STRONG · 34 AMBIGUOUS   (de 1156 elegibles)
  REMASEP_01          165 EXACT · 75 STRONG                   (de 240 elegibles)
  B2_ANEXO            5 EXACT                                 (de 5 elegibles, por código de prestación)

procedure-code exact matches   5
1122 targets EXACT distintos · 0 con coordenada igual a la del source
```

- **REMASEP_OD** se alinea muy bien (vocabulario compartido, ambos layouts se
  parsean). Los `AMBIGUOUS` son filas-subtotal cuyo contexto de fila es fino y
  encaja en muchas filas de la plantilla — se marcan, **no** se resuelven.
- **REMASEP_01** se parsea peor que REMASEP_OD (su layout en la plantilla es más
  complejo); aun así 165/240 EXACT.
- **B2 ANEXO** se alinea por **código de prestación** exacto (5/5).
- **REMASEP B1**: 0 sources (no hay subgrafo MEDINET); sus celdas de grilla son
  todas `EGRESOS` en `unmatched_targets.csv`.

### Targets — inventario físico, rol y cobertura (conceptos separados)

```
inventario físico total          16538
  por rol   INPUT_TARGET 8379 · DERIVED_TARGET 1174 · STRUCTURAL 6985
  por fuente esperada   MEDINET 6697 · EGRESOS 5861 · RESOURCE_CALCULATION 280 · UNKNOWN 3700

de los 8379 INPUT_TARGET (= celdas desbloqueadas en hoja protegida):
  MEDINET esperados            5227
    alineados                  1367   (los 1122 EXACT + 245 STRONG, todos caen en INPUT_TARGET)
    genuinamente sin source    3860   ← "faltó encontrar un source" (NO_MEDINET_SOURCE_FOUND)
  NO-MEDINET (EGRESOS/recursos) 1930  ← "nunca debían tener source MEDINET"
  UNKNOWN (fuente no demostrable) 1222

filas técnicas sin match (todas las clases)   15171
  NO_MEDINET_SOURCE_FOUND      3860   (genuino)
  UNKNOWN_SOURCE               1222
  NON_MEDINET_REGION           1930
  DERIVED_TARGET_NOT_INPUT     1174   (fórmulas — no son puntos de ingreso)
  STRUCTURAL_NOT_ALIGNMENT_TARGET 6985  (rótulos / encabezados / códigos)
```

> Revisión de cierre: el criterio de `DIRECT_INPUT_TARGET` pasó de "no tiene
> fórmula" a "**celda desbloqueada en hoja protegida**". **2 906** celdas
> (rótulos, encabezados, códigos de prestación) dejaron de contarse como input
> (`INPUT_TARGET` 11285 → **8379**; `STRUCTURAL` 4079 → **6985**). El matching
> **no cambió** (los 1367 targets alineados ya estaban todos desbloqueados). Los
> targets MEDINET input genuinamente sin alinear bajan de 4237 a **3860**.

---

## Golden reference

El REMASEP final julio se usa como **referencia estructural / semantic target**,
pero se mantiene `approval_status = UNCONFIRMED` y
`candidate_golden_status = CANDIDATE_PENDING_SOURCE_COMPLETENESS`. **No** es un
golden aprobado.

---

## Archivos — `artifacts/official_template_semantic_alignment/`

| archivo | contenido |
| --- | --- |
| `target_metric_inventory.csv` | celdas de la plantilla con `target_kind`, `target_alignment_role`, `target_source_expectation` y contexto |
| `source_metric_inventory.csv` | métricas MEDINET con readiness y `eligible` |
| `semantic_alignment.csv` | una fila por par, con evidencia por dimensión, `target_alignment_role` y `match_score` |
| `alignment_evidence.csv` | evidencia dimensión a dimensión (source vs target) |
| `ambiguous_matches.csv` | sources con varios targets plausibles |
| `unmatched_sources.csv` | sources sin match (`NO_TARGET` / `TARGET_SOURCE_NOT_MEDINET` / `CONTEXT_INSUFFICIENT` / `CONFLICT`) |
| `unmatched_targets.csv` | **todas** las celdas sin match, con `reason` tipado (`NO_MEDINET_SOURCE_FOUND` = genuino · `NON_MEDINET_REGION` · `DERIVED_TARGET_NOT_INPUT` · `STRUCTURAL_NOT_ALIGNMENT_TARGET` · `VALIDATION_NOT_ALIGNMENT_TARGET` · `UNKNOWN_SOURCE`) + `is_genuine_unmatched_medinet_input` |
| `source_exclusions.csv` | sources no elegibles (`BLOCKED_CONFLICT` · `REVIEW_REQUIRED` · `ROLLUP_NOT_INPUT`) |
| `target_source_expectations.csv` | regiones declaradas (NO-MEDINET y MEDINET) |
| `alignment_coverage.csv` | por forma: lado source (match_status) y lado target (inventario por rol + cobertura MEDINET input) |
| `alignment_manual_review_sample.csv` | muestra determinista y estratificada, con `target_kind` / `target_alignment_role` / `target_source_expectation` |
| `summary.json` | cantidades **separadas** (inventario físico · por rol · input targets MEDINET esperados / alineados / genuinamente sin source); `alignment_status: PRELIMINARY_NOT_VALIDATED` |

## Privacidad

Sólo `metric_id`, coordenadas, rótulos del formulario, fórmulas y estados. Sin
datos individuales de paciente.

## Limitaciones

- El layout de la plantilla oficial difiere del generador legacy; `REMASEP 01` se
  parsea peor que `REMASEP_OD`, lo que baja su cobertura.
- `REMASEP B1` no tiene subgrafo MEDINET (todo `EGRESOS`).
- Las regiones de `source_regions.yaml` (NO-MEDINET **y** MEDINET) son
  **preliminares**: derivadas de un solo mes de provenance, sin confirmación
  funcional. La `B2_SURGICAL_INTERVENTIONS` usa un rango de filas aproximado
  porque `SheetLayout` no expone la sección en el catálogo plano de B2.
- `target_source_expectation = UNKNOWN` cuando ninguna región reclama la celda:
  no se infiere `MEDINET` por defecto (revisión de cierre Sprint 3.5).
- Un `EXACT_SEMANTIC_MATCH` es coincidencia estructural de rótulos normalizados,
  **no** validación funcional MINSAL.
- El mapping **no** se usa todavía para escribir el workbook oficial (no hay COM,
  no hay celdas de escritura).
