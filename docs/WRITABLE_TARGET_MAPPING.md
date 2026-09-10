# Writable target mapping (Sprint 3.6)

`scripts/build_write_manifest.py` +
`src/remasep/services/writable_target_mapping.py` +
`config/writable_target_mapping_2026/`.

Sprint 3.5 emparejó cada `SemanticMetric` MEDINET elegible con celdas de la
plantilla oficial (`EXACT` / `STRONG` / `AMBIGUOUS`). Sprint 3.6 convierte los
alignments **técnicamente seguros** en un conjunto explícito, auditable y
**versionable** de instrucciones de escritura.

> **No se escribe Excel.** 3.6 decide **DÓNDE** escribir, no **QUÉ** valor (eso es
> el *value producer* del Sprint 3.7) ni **QUÉ** población clínica entra (el
> filtro por ESTADO sigue `PENDING_FUNCTIONAL_CONFIRMATION` y **no** se aplica).

## Semantic alignment vs. write mapping — capas separadas

| capa | pregunta | estados |
| --- | --- | --- |
| `readiness` (3.3) | ¿la métrica MEDINET está resuelta? | `AUTO_READY` / `REVIEW_REQUIRED` / `BLOCKED_CONFLICT` / `NOT_APPLICABLE` |
| `match_status` (3.5) | ¿hay una celda de la plantilla que representa lo mismo? | `EXACT_SEMANTIC_MATCH` / `STRONG_MATCH` / `AMBIGUOUS` / `NO_MATCH` / `CONFLICT` |
| **`write_status` (3.6)** | ¿puede convertirse en una instrucción de escritura automática? | `WRITE_READY` / `WRITE_REVIEW_REQUIRED` / `WRITE_BLOCKED` / `NOT_WRITABLE` |

Nunca se reutiliza `AUTO_READY` / `EXACT` / `STRONG` para hablar de escritura.

## Regla base — candidato a `WRITE_READY`

Todas estas condiciones (`config/.../policy.yaml`):

- `source.readiness == AUTO_READY`
- `match_status ∈ {EXACT_SEMANTIC_MATCH, STRONG_MATCH}`
- `target.target_kind == DIRECT_INPUT_TARGET`
- `target.target_alignment_role == INPUT_TARGET`
- **`target` unlocked** (revalidado contra el workbook — *defense in depth*)
- `target_source_expectation == MEDINET`
- sin conflicto de dimensiones, **target único**, **source único**, **sin colisión**

### `EXACT_SEMANTIC_MATCH`

Cumple la regla base y no hay condición bloqueante → **`WRITE_READY`**.

### `STRONG_MATCH` — política explícita, sin bajar umbrales

`STRONG` **no** se convierte automáticamente. Sólo es `WRITE_READY` si:

- ninguna dimensión evaluable está en `MISMATCH`;
- **todas** las dimensiones requeridas de la forma están en `MATCH`
  (`REMASEP_OD` / `REMASEP_01`: `SECTION · ROW_PATH · COLUMN_PATH · SEX · AGE`;
  `B2_ANEXO`: `ROW_PATH · PROCEDURE_CODE`);
- el resto (`MISSING` / `NOT_APPLICABLE`) son sólo dimensiones **no** requeridas;
- target único, sin colisión.

Si falta la evidencia de una dimensión requerida (p.ej. `ROW_PATH = PARTIAL`,
`SEX = MISSING`) → **`WRITE_REVIEW_REQUIRED`**, con `missing_evidence` explícito y
agrupado en `write_review_clusters.csv`.

### `AMBIGUOUS`

→ **`WRITE_BLOCKED`**. Nunca se elige candidato; todos los candidatos quedan en
`write_blocked.csv`. (En los datos reales los 34 `AMBIGUOUS` apuntan además a
celdas de **fórmula**, así que estarían bloqueados por partida doble.)

## Exclusiones de source

| readiness | `write_status` / `write_reason` |
| --- | --- |
| `REVIEW_REQUIRED` | `NOT_WRITABLE` / `SOURCE_REVIEW_REQUIRED` |
| `BLOCKED_CONFLICT` (los 5 `REMASEP 01!AB84…AB90`) | `WRITE_BLOCKED` / `SOURCE_BLOCKED_CONFLICT` |
| `NOT_APPLICABLE` (roll-up TOTAL/SUBTOTAL) | `NOT_WRITABLE` / `ROLLUP_COMPUTED_BY_TEMPLATE` |

**No se crean instrucciones de escritura para roll-ups**: el workbook los calcula.

## Non-MEDINET targets

Nunca una `WriteInstruction` MEDINET hacia `EGRESOS` / `RESOURCE_CALCULATION` /
`SURGICAL_TABLE` / `CONTROL_METADATA` → `WRITE_BLOCKED` /
`TARGET_SOURCE_NOT_MEDINET`. Si la expectativa es `UNKNOWN` y el alignment
resuelto **no** la promovió a `MEDINET` (Sprint 3.5) → `WRITE_REVIEW_REQUIRED`.

## Formula targets

Nunca se escribe sobre `FORMULA_TARGET` (`NOT_WRITABLE` / `TARGET_IS_FORMULA`).
Los roll-ups quedan calculados por el workbook; se puede registrar la relación de
validación `source semantic total → target formula`, pero no como escritura.

## Cell protection — *defense in depth* (§17)

Durante la construcción del manifiesto se **re-lee** `protection.locked` de cada
celda-target contra el workbook. Si contradice al alignment (celda locked que 3.5
consideró válida por error) → `WRITE_BLOCKED` / `TARGET_LOCKED`.

## Colisiones — nunca *last write wins*

Se detectan explícitamente:

| tipo | qué |
| --- | --- |
| `MULTIPLE_SOURCES_SAME_TARGET` | varias métricas apuntan a la misma celda |
| `SAME_SOURCE_MULTIPLE_TARGETS` | una métrica resuelta a varias celdas |
| `DUPLICATE_INSTRUCTION_IDENTITY` | mismo `instruction_id` para varios sources |
| `INCOMPATIBLE_TARGET_REUSE` | celda reusada por instrucciones incompatibles |

Cualquier colisión → **`WRITE_BLOCKED` para TODOS los implicados** (sin
instrucción). `write_collisions.csv`.

## Identidad semántica vs. coordenada física

`instruction_id = "wi:" + sha256(source_semantic_signature ⋮
target_semantic_signature ⋮ structural_template_fingerprint_id)[:20]`.

**La coordenada (`H99`) es ubicación física, no identidad.** El mismo par
semántico produce el mismo `instruction_id` aunque cambie la celda.

## Template fingerprint — `file_hash` ≠ `structural_template_fingerprint`

| | qué |
| --- | --- |
| `file_sha256` | hash del `.xlsm`. **Informativo.** Un re-guardado (VBA recompilado, metadatos) lo cambia sin cambiar nada semántico. |
| `structural_template_fingerprint_id` | hash de: nombres/orden de hojas · recuento y coords de fórmulas por hoja · hoja protegida sí/no · recuento y coords de celdas **desbloqueadas** · merges relevantes. **El manifiesto se ata a esto.** |

El manifiesto se **invalida** si cambia la estructura semántica (`policy.yaml →
template_compatibility.invalidate_on`), **no** por un simple re-guardado.

## Contrato de valor (Sprint 3.7)

`MetricValue(source_metric_id, value, period, producer_version)` — **sin PII**.
El *value producer* del Sprint 3.7:

```
source_metric_id  --producer-->  MetricValue
WriteInstruction + MetricValue  --writer-->  (futuro) escritura Excel
```

**Contrato de período**: los `MetricValue` MEDINET se calculan **exclusivamente**
desde `processing_scope_records` (registros válidos ∩ mes/año), **no** desde
todos los válidos, y **sin** aplicar el filtro por ESTADO pendiente de
confirmación (ver [`docs/MEDINET_INPUT_CONTRACT.md`](MEDINET_INPUT_CONTRACT.md)).

`zero_write_policy = UNRESOLVED`: escribir `0` vs. dejar la celda vacía es una
decisión del *writer*, no de la elegibilidad de escritura. Un mapping puede ser
`WRITE_READY` estructuralmente con esto abierto.

## Resultado real

Generador `GENERACION DATOS REMASEP.xlsx` vs. plantilla
`REMASEP_V1.4 Julio 2026.xlsm` (`structural_template_fingerprint_id
= stf:dc624775927d4d4d`).

```
source metrics        1771
  AUTO_READY           1401
  alineados            1401   (1122 EXACT · 245 STRONG · 34 AMBIGUOUS)

write_status
  WRITE_READY          1122   (todos EXACT; 0 STRONG)
  WRITE_REVIEW_REQUIRED 245   (todos STRONG: ROW_PATH PARTIAL 170 · SEX MISSING 65 · AGE MISSING 10)
  WRITE_BLOCKED          39   (34 AMBIGUOUS + 5 SOURCE_BLOCKED_CONFLICT)
  NOT_WRITABLE          365   (315 ROLLUP_COMPUTED_BY_TEMPLATE + 50 SOURCE_REVIEW_REQUIRED)

WRITE_READY por forma  REMASEP_OD 952 · REMASEP_01 165 · B2_ANEXO 5
manifest instructions  1122
colisiones             0

template_formula_targets                  1174   (celdas de fórmula físicas en la plantilla)
source_rows_blocked_by_formula_target       34   (alignments cuyo candidato cae en una
                                                  celda de fórmula; además AMBIGUOUS)
```

> **No confundir**: `template_formula_targets = 1174` son celdas físicas de la
> plantilla (inventario de Sprint 3.5). `source_rows_blocked_by_formula_target =
> 34` son *source rows* cuyo candidato resuelto apunta a una de esas celdas — y
> todos son además `AMBIGUOUS`, así que van a `WRITE_BLOCKED` por partida doble.

`STRONG` mostrando cobertura 0 en `WRITE_READY` es el resultado **correcto**: la
policy no baja umbrales; cada `STRONG` tiene al menos una dimensión requerida sin
demostrar, y va a revisión.

## Archivos — `artifacts/writable_target_mapping/`

| archivo | contenido |
| --- | --- |
| `write_readiness.csv` | una fila por source relevante (los 1771) |
| `write_manifest_ready.csv` | **sólo** las 1122 `WRITE_READY` — sin PII, sin valores |
| `write_review_queue.csv` / `write_review_clusters.csv` | los 245 a revisar, agrupados por causa |
| `write_blocked.csv` | ambiguos, conflictos, locked, colisiones |
| `write_collisions.csv` | colisiones detectadas |
| `write_evidence.csv` | evidencia por dimensión (reexportada de 3.5) |
| `coverage_by_form.csv` / `coverage_by_match_status.csv` | cobertura |
| `template_fingerprint.json` | `file_sha256` vs. `structural_template_fingerprint` |
| `write_mapping_manual_review_sample.csv` | muestra determinista y estratificada (incluye `target_locked` y `expected_value_type`) |
| `summary.json` | totales; `write_status: PRELIMINARY_NOT_VALIDATED` |

## Limitaciones

- El manifiesto dice DÓNDE, no QUÉ valor ni QUÉ población clínica.
- `STRONG` con `ROW_PATH PARTIAL` / `SEX MISSING` va a revisión — no se bajan
  umbrales para cobertura.
- Un `WRITE_READY` es coincidencia estructural + policy, **no** validación
  funcional MINSAL.
- `zero_write_policy` sigue `UNRESOLVED` (precondición del Sprint 3.7).
- No se escribe Excel, no hay COM, no se aplica filtro por ESTADO, no se
  resuelven los 34 `AMBIGUOUS` ni los 5 conflictos.
