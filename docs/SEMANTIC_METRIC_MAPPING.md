# Mapeo semántico de métricas (Sprint 3.2)

`scripts/map_semantic_metrics.py` + `src/remasep/services/semantic_mapping.py` +
`config/semantic_mapping_2026/`.

Sprint 3.1 produjo el **contexto** preliminar de las 1771 métricas MEDINET
(rótulos de fila / columna / sección + evidencia de fórmula). Sprint 3.2 lo
convierte en **dimensiones explícitas y auditables** — sexo, edad, alcance de
agregación, código de procedimiento — cada una con `value`, `status` y
`evidence`.

> **Este mapping NO es todavía un mapping a la plantilla oficial MINSAL**, ni una
> validación funcional por Fundación Gantz. El vocabulario vive en
> `config/semantic_mapping_2026/` con
> `status: preliminary_pending_functional_validation`.

---

## Context vs. dimension vs. evidence

| | Qué es | De dónde |
| --- | --- | --- |
| **context** (Sprint 3.1) | los rótulos que un humano *ve* alrededor de la métrica (`row_path`, `column_path`, `section_path`) | layout del workbook |
| **dimension** (Sprint 3.2) | un hecho semántico explícito (`sex = FEMALE`, `age = 20..24`) | inferido del context + la fórmula |
| **evidence** | de dónde sale cada dimensión, agregada y sin PII | `COLUMN_LABEL` / `ROW_LABEL` / `FORMULA_CRITERION` / `FORMULA_BOUNDS` / `SHEET` |

**Regla fundamental: no inventar semántica.** Sin evidencia suficiente la
dimensión queda `LABEL_ONLY` / `FORMULA_ONLY` / `NOT_APPLICABLE` / `UNRESOLVED`.
Nunca un valor adivinado. Una heurística nunca se convierte en hecho.

---

## Estados por dimensión

| Estado | Significado |
| --- | --- |
| `CONFIRMED` | evidencia de rótulo **y** de fórmula, compatibles |
| `LABEL_ONLY` | sólo el rótulo lo indica |
| `FORMULA_ONLY` | sólo la fórmula lo indica |
| `NOT_APPLICABLE` | la métrica no tiene esa dimensión (p.ej. sexo en un total, o en `B2 ANEXO`) |
| `UNRESOLVED` | podría tener la dimensión pero no se pudo determinar |
| `CONFLICT` | el rótulo y la fórmula se contradicen — **se registra, no se corrige** |

---

## Provenance / `source`

`source = "MEDINET"` en **todas** las métricas de este sprint. El modelo queda
preparado conceptualmente para `EGRESOS` y `RESOURCE_CALCULATION` (ver
[`CLIENT_PENDING.md`](CLIENT_PENDING.md) y
[`TECHNICAL_OVERVIEW.md`](TECHNICAL_OVERVIEW.md)) — **sin código** en este sprint.

---

## Dimensiones

### `form`

Del nombre de hoja (`config/semantic_mapping_2026/form_labels.yaml`):
`REMASEP 01` → `REMASEP_01`, `B2 ANEXO` → `B2_ANEXO`, `REMASEP_OD` → `REMASEP_OD`.
Se conserva `form_raw` (el nombre de hoja tal cual).

### `sex`

Vocabulario explícito y versionado (`sex_labels.yaml`), **no** substring
arbitrario:

- rótulo de columna → `Hombres`/`Hombre` → `MALE`; `Mujeres`/`Mujer` → `FEMALE`;
  `Ambos sexos`/`Ambos sexo` → `BOTH` (matching por rótulo **normalizado**).
- criterio de fórmula → literales entre comillas que contengan `hombre`/`mujer`
  (incluye el patrón `$A8&"Hombre"` sin re-parsear la fórmula).
- `CONFIRMED` si rótulo y fórmula coinciden; `LABEL_ONLY` / `FORMULA_ONLY` si sólo
  hay uno; **`CONFLICT`** si el rótulo dice `Mujeres` y la fórmula filtra
  `"*Hombre*"`.

### `age` (`age_min_years`, `age_max_years`)

Reutiliza el trabajo de Sprint 3.1 (rótulo de edad + cotas `$AF:$AF`). Inclusivo;
`age_max_years` vacío = extremo abierto.

| rótulo | fórmula | resultado |
| --- | --- | --- |
| `Menos de 1 año - 1 año` | `AF < 2` | `min=0, max=1`, `CONFIRMED` |
| `20 A 24 AÑOS` | `AF >=20 <=24` | `min=20, max=24`, `CONFIRMED` |
| `75 Y MÁS` | `AF >=75` | `min=75, max=None`, `CONFIRMED` |

`CONFIRMED` sólo si los intervalos (rótulo y fórmula) se **solapan**; disjuntos →
`CONFLICT`. **No** se acuña todavía `AGE_20_24` como código definitivo.

### `aggregation_scope`

`DETAIL` / `TOTAL` / `SUBTOTAL`, con `status` propio. Sólo por evidencia
estructural segura, **no** por posición:

- rótulo de columna `TOTAL` → `TOTAL`;
- la celda está sobre una fila-encabezado de sección/grupo (`A.4. …`,
  `A. KINESIOLOGÍA`) → `SUBTOTAL`;
- si hay una dimensión de edad/sexo o un código de procedimiento → `DETAIL`
  `CONFIRMED`; en otro caso `DETAIL` `UNRESOLVED`.

**`kind` (`DOWNSTREAM_TOTAL`) y `aggregation_scope` son conceptos distintos**: un
`SUM` no se marca `TOTAL` automáticamente — sólo si su rótulo de columna lo dice.

### `procedure_code_raw` / `procedure_label_raw`

Sólo cuando hay un código **explícito** en un rótulo de fila
(`config/semantic_mapping_2026/procedure_codes.yaml`, 4–8 dígitos):

- `"5010009 - VIDRIO IONÓMERO"` → `code="5010009"`, `label="VIDRIO IONÓMERO"`;
- `"0601105"` + rótulo siguiente → `code="0601105"`, `label="Atención…"` (patrón
  de `B2 ANEXO`).

Números incidentales (`20 A 24`, `…38:51`) **no** se extraen. **No** se afirma a
qué catálogo oficial (FONASA / MINSAL) pertenece el código.

---

## `mapping_status` (≠ `context_status` de Sprint 3.1)

- **`CONFLICT`** — alguna dimensión tiene `CONFLICT`.
- **`CONFIRMED`** — `sex` y `age` están `CONFIRMED` o justificadamente
  `NOT_APPLICABLE`, **y** `aggregation_scope` está `CONFIRMED`.
- **`PARTIAL`** — alguna dimensión queda `LABEL_ONLY` / `FORMULA_ONLY` /
  `UNRESOLVED`. Es **información válida**, no un fallo.

Nunca se usa `COMPLETE` (para no confundir con el `context_status` de Sprint 3.1).

---

## Resultado del workbook de referencia

`GENERACION DATOS REMASEP.xlsx` (sha256 `fc2e1536…d15b79`), julio 2026 —
**1771 SemanticMetric**.

| `mapping_status` | métricas |
| --- | ---: |
| `CONFIRMED` | **1414** |
| `PARTIAL` | 352 |
| `CONFLICT` | **5** |

| Dimensión | `CONFIRMED` | `LABEL_ONLY` | `FORMULA_ONLY` | `NOT_APPLICABLE` | `UNRESOLVED` | `CONFLICT` |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| `sex` | 1331 | 306 | 0 | 129 | 0 | **5** |
| `age` | 1391 | 194 | 2 | 184 | 0 | 0 |

`aggregation_scope`: `DETAIL` 1456 (1436 `CONFIRMED` + 20 `UNRESOLVED`),
`SUBTOTAL` 176, `TOTAL` 139. `procedure`: **597 `EXPLICIT`** (19 códigos
distintos), 267 `NOT_APPLICABLE`, 907 `UNRESOLVED`.

- **REMASEP_OD `BASE` + `DERIVED` (1156): 1156 `CONFIRMED`** — la grilla
  edad × sexo se confirma con rótulo **y** fórmula.
- **Los 352 `PARTIAL`** son sobre todo celdas `DOWNSTREAM_TOTAL` (`SUM`) con
  rótulo de sexo pero sin criterio de fórmula (→ `sex LABEL_ONLY`), y filas de
  `B2 ANEXO` sin código explícito.
- **Los 5 `CONFLICT`** son reales y del **workbook**: `REMASEP 01!AB84/86/87/89/90`
  están en la columna **`Mujeres`** (`50-54 años`) pero su fórmula filtra
  `"*Hombre*"` — un error de copiar/pegar en el generador. Se **registra**
  (`mapping_conflicts.csv`), no se corrige.

---

## Muestra de revisión manual

`manual_review_sample.csv` — muestra **determinista y estratificada** (**38**
métricas en el workbook de referencia: 25 `CONFIRMED` · 8 `PARTIAL` · 5
`CONFLICT`; cada hoja, cada `kind`, `MALE`/`FEMALE`/`BOTH`, con y sin código de
procedimiento, distintas bandas de edad, `PARTIAL`, y **todos** los `CONFLICT`).
Columnas: `metric_id`, `source`, `form`, `kind`, `cell`, `row_path_raw`,
`column_path_raw`, dimensiones y `mapping_status` — autocontenida para verificar
la estratificación `BASE`/`DERIVED`/`DOWNSTREAM_TOTAL`. El tamaño exacto lo
publica `summary.json` → `manual_review_sample_size` (`_by_status` incluido), así
que artefacto y summary siempre coinciden. Para revisión humana; los mappings
**no** se ajustan automáticamente en función de ella.

> El sample lo componen ≈ 17 *buckets* de hasta 3 métricas cada uno (ordenadas
> por `metric_id`) más **todos** los `CONFLICT`, deduplicados por `metric_id`:
> el solapamiento entre buckets hace que el total (38) sea menor que 17 × 3.

---

## Limitaciones

- El vocabulario (`config/semantic_mapping_2026/`) cubre lo **observado** en este
  workbook; `status: preliminary_pending_functional_validation`.
- La evidencia de fórmula reutiliza la de Sprint 3.1 (comodines `"*x*"`,
  cotas `$AF:$AF`) más los literales `"Hombre"`/`"Mujer"` de concatenaciones
  `$Ax&"…"`. Un criterio de sexo por otro mecanismo quedaría `LABEL_ONLY`.
- `PARTIAL` / `UNRESOLVED` es un resultado válido, no un fallo del pipeline.
- **No** hay dimensiones de actividad / especialidad codificadas (`row_path`
  sigue siendo la dimensión de fila). **No** hay mapping a la plantilla oficial.
  **No** hay validación funcional del cliente.

---

## Artefactos

`scripts/map_semantic_metrics.py "<workbook>"` →
`artifacts/semantic_metric_mapping/`:

| Archivo | Contenido |
| --- | --- |
| `semantic_metrics.csv` | una fila por métrica con todas las dimensiones + `mapping_status` |
| `dimension_evidence.csv` | una fila por evidencia (raw + normalizado, `evidence_status`) |
| `mapping_coverage.csv` | cobertura por `form` × `kind` |
| `mapping_conflicts.csv` | conflictos rótulo ↔ fórmula (sin resolver) |
| `procedure_codes.csv` | códigos de procedimiento explícitos encontrados |
| `manual_review_sample.csv` | muestra determinista estratificada (**38** en el workbook de referencia; con `source` y `kind`) |
| `summary.json` | totales; `mapping_status: PRELIMINARY_NOT_VALIDATED` |
| `README.md` | resumen de la corrida |

### Privacidad

Sólo rótulos del formulario, coordenadas y valores agregados. Ninguna fila
Medinet individual.
