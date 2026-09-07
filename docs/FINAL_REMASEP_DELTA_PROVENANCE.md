# Delta del REMASEP final & provenance de fuente (Sprint 3.4)

`scripts/analyze_final_remasep_delta.py` +
`src/remasep/services/workbook_delta.py` +
`src/remasep/services/source_provenance.py`.

## Por qué se comparan las dos versiones

Tenemos dos archivos del REMASEP oficial de julio 2026:

| | archivo | estado |
| --- | --- | --- |
| **before** | `data/local/2026-7 REMASEP_V1.4.xlsm` | `INCOMPLETE_PRE_SURGERY` — el cliente confirmó que **no** estaba terminado: faltaban las cirugías |
| **after** | `data/local/REMASEP_V1.4 Julio 2026.xlsm` | `FINAL_PER_CLIENT` · `approval_status = UNCONFIRMED` — "REMASEP julio terminado" según el cliente; **no** consta que se enviara ni que MINSAL lo aceptara |

Comparar ambas versiones responde dos preguntas:

1. **¿qué cambió** entre el REMASEP incompleto y el final?
2. **¿de qué fuente / proceso proviene** cada cambio?

Sólo lectura: `openpyxl` con `data_only` en `False` (fórmulas) y `True` (valores
cacheados). **No** se modifica ni se guarda ningún workbook, **no** se usa COM,
LibreOffice ni macros, **no** se recalcula. El SHA256 de ambos archivos se
verifica antes y después del análisis.

> El PDF `EGRESO HOSPITALARIO 2025.pdf` es el **formulario/instructivo** de egreso
> hospitalario, no una base mensual. **No** se parsea, **no** hay OCR, **no** se
> trata como los egresos de julio. Queda documentado como *especificación de
> fuente*; el cliente aún debe entregar un reporte real de egresos.

## Resultado (archivos reales)

```
formula_expression_changes = 0
total_changed_cells        = 113   (52 input directo + 61 propagación de fórmula)

por hoja:  REMASEP 01 = 58 · REMASEP B1 = 41 · B2 ANEXO = 10 · CONTROL = 4
```

Único cambio estructural detectado: el **payload VBA** tiene distinto SHA256 —
con 0 cambios de fórmula ni de dimensiones, es compatible con un simple
re-guardado en Excel (que recompila el proyecto). Se reporta, no se interpreta
como cambio de lógica.

## `change_kind` — input directo vs. resultado vs. expresión

Se decide comparando la fórmula **antes y después** (no sólo la de la versión
final):

| valor | criterio |
| --- | --- |
| `DIRECT_INPUT_CHANGE` | sin fórmula ni en *before* ni en *after*: es un dato ingresado |
| `FORMULA_RESULT_CHANGE` | la **expresión** de la fórmula es idéntica en ambas versiones; cambió sólo su valor cacheado por propagación de un cambio aguas arriba |
| `FORMULA_EXPRESSION_CHANGE` | la **expresión** de la fórmula cambió — añadida, eliminada o modificada (`formula_change_type` ∈ `FORMULA_ADDED` / `FORMULA_REMOVED` / `FORMULA_MODIFIED`) |

Una *formula cache difference* **no** es un input manual. Un
`FORMULA_EXPRESSION_CHANGE` **no** se trata como input directo ni como
propagación de resultado: su `source` queda `UNKNOWN_PENDING` /
`needs_human_review` salvo evidencia explícita, y se reportan `before_formula` y
`formula` (después). *(En los archivos reales de julio 2026 no hay ninguno:
`formula_expression_changes = 0`.)*

## `source` — provenance

`change_origin` (¿directo o propagado?) es **distinto** de `source` (¿de qué
fuente?).

| source | de dónde | `status` observado |
| --- | --- | --- |
| `EGRESOS` | edad/sexo de cirugía → REMASEP B1 Sección A; códigos de intervención quirúrgica → B2 ANEXO | `CLIENT_CONFIRMED` (directo) / `DERIVED_FROM_DEPENDENCIES` (propagado) |
| `RESOURCE_CALCULATION` | capacidad instalada de quirófanos (días hábiles × ~7–8 h) — REMASEP 01 Sección D, columnas de dotación / habilitados / horas hábiles | `STRUCTURALLY_INFERRED` |
| `SURGICAL_TABLE` | horas programadas / ocupadas de tabla quirúrgica — REMASEP 01 Sección D, columnas de tabla quirúrgica / horas ocupadas | `STRUCTURALLY_INFERRED` |
| `CONTROL_METADATA` | causales de hojas sin datos en la hoja `CONTROL` | `STRUCTURALLY_INFERRED` |
| `MEDINET` | subgrafo ambulatorio (Sprints 2–3) — no aparece en este delta | — |
| `UNKNOWN_PENDING` | no asignable con la evidencia actual: Sección E (causas de suspensión), columnas ambiguas de Sección D | `UNKNOWN` |
| `MIXED_DERIVED` | una fórmula alimentada por celdas cambiadas de **varias** fuentes distintas | `DERIVED_FROM_DEPENDENCIES` |

### Regla fundamental: no inferir más de lo que la evidencia permite

Sólo se marca `CLIENT_CONFIRMED` donde el cliente lo confirmó explícitamente
(edad/sexo → B1, códigos Qx → B2 ANEXO). La asignación por columna de la
Sección D usa el **texto de los encabezados** del formulario y queda
`STRUCTURALLY_INFERRED` con `needs_human_review = yes`; lo que no calza queda
`UNKNOWN_PENDING`. Nunca se "fuerza" una fuente.

### Provenance de los cambios propagados

Se construye un grafo de dependencia estructural entre las celdas cambiadas
(referencias A1 de la fórmula *después*) y se propaga:

```
EGRESOS (input directo B2 ANEXO)
   ↓  C1317 / C1377 / …
B2 subtotal  =SUM(...)                 source = EGRESOS   status = DERIVED_FROM_DEPENDENCIES
   ↓  'B2 ANEXO'!C1308
REMASEP B1 Sección B  ='B2 ANEXO'!C1308  source = EGRESOS   status = DERIVED_FROM_DEPENDENCIES
```

Si una fórmula reúne varias fuentes (p.ej. el gran total `REMASEP 01!A307`, que
suma Sección D — recursos + tabla quirúrgica — **y** Sección E — desconocida) →
`MIXED_DERIVED`. Si no hay ninguna celda cambiada trazable aguas arriba →
`UNKNOWN_PENDING`.

## Resultado de provenance (archivos reales)

| source | directo | derivado | total | revisión humana |
| --- | ---: | ---: | ---: | ---: |
| `EGRESOS` | 14 | 37 | 51 | 0 |
| `SURGICAL_TABLE` | 19 | 11 | 30 | 30 |
| `UNKNOWN_PENDING` | 10 | 7 | 17 | 17 |
| `RESOURCE_CALCULATION` | 5 | 5 | 10 | 10 |
| `CONTROL_METADATA` | 4 | 0 | 4 | 0 |
| `MIXED_DERIVED` | 0 | 1 | 1 | 1 |

`provenance_coverage`: 96 clasificadas / 17 `UNKNOWN_PENDING`.

## Egresos

### REMASEP B1 — Sección A (cirugías por grupo de edad / sexo)

8 inputs directos nuevos en la fila *Electivas · Mayor Ambulatorias*
(`D13:M13`) → `source = EGRESOS`, `status = CLIENT_CONFIRMED`.

Consistency checks generados (`source_consistency_checks.csv`):

- **`B1_SECCION_A_AGE_TOTAL_VS_SEX_TOTAL`**: Σ columnas por edad (32) = Σ columnas
  por sexo (32) → `OK`. *(no se hardcodea el 32: se descubre y se compara)*.

### B2 ANEXO — códigos de intervención quirúrgica

6 inputs directos nuevos → `source = EGRESOS`, `status = CLIENT_CONFIRMED`.
Códigos encontrados:

| celda | código | prestación | count |
| --- | --- | --- | ---: |
| C1317 | 1302008 | Tratamiento quirúrgico de Mucositis timpánica… | 6 |
| C1377 | 1302052 | Rinoplastía y/o septoplastía, cualquier técnica | 9 |
| C1470 | 1402052 | Osteotomías segmentarias del maxilar o mandíbula | 2 |
| C1513 | 1502028 | Corrección nasal parcial (alares, alargamiento columela…) | 3 |
| C1518 | 1502033 | Cierre de paladar duro y/o cierre de comunicación oro-nasal | 3 |
| C1524 | 1502039 | Reconstrucción osteoplástica reborde alveolar unilateral | 8 |

Los subtotales de B2 (`C1308 = 15`, `C1419 = 2`, `C1480 = 14`) y las filas de
Sección B de B1 que hacen *pull* de ellos son `EGRESOS` /
`DERIVED_FROM_DEPENDENCIES` — **no** se marcan como input de egresos
independiente.

## Recursos / tabla quirúrgica — REMASEP 01 Sección D

40 celdas cambiadas en *SECCIÓN D: CAPACIDAD INSTALADA Y UTILIZACIÓN DE
QUIRÓFANOS* (fila *De Cirugía electiva* y sus roll-ups):

- **10** → `RESOURCE_CALCULATION` (columnas *NÚMERO DE QUIRÓFANOS EN DOTACIÓN*,
  *PROMEDIO … HABILITADOS / EN TRABAJO*, *TOTAL DE HORAS MENSUALES … HÁBIL*).
- **30** → `SURGICAL_TABLE` (columnas *Horas Mensuales Programadas de Tabla
  Quirúrgica…*, *Horas Mensuales Ocupadas…*).

Todas con `status = STRUCTURALLY_INFERRED` y `needs_human_review = yes`: el
cliente confirmó el **proceso** (capacidad = días hábiles × 7–8 h; ocupación se
revisa desde la tabla quirúrgica) pero **no** la regla columna por columna.

## CONTROL metadata

4 cambios directos en la hoja `CONTROL`, columna *Causal (Indicar)*:

| hoja sin datos | causal |
| --- | --- |
| URGENCIAS | No tenemos servicio de Urgencia |
| EXÁM Y PESQ ENF. TRANSMISIBLES | No tenemos laboratorio |
| TV MI | No tenemos atención materno infantil |
| SERVICIOS DE SANGRE | No tenemos laboratorio |

→ `source = CONTROL_METADATA`. No son producción Medinet/Egresos.

### CONTROL final

`TOTAL ERRORES = 0` → `control_status = PASS_INTERNAL_VALIDATION`. Hojas en
estado `SIN DATOS` (con causal): URGENCIAS, EyP_ET, TV_MI, SERV_SANGRE; el resto
`OK`.

> `CONTROL = 0` **no** significa "MINSAL approved". `approval_status` se mantiene
> `UNCONFIRMED`.

## Sección E — causas de suspensión: fuente desconocida

10 inputs directos nuevos en *SECCIÓN E: CAUSAS DE SUSPENSIÓN DE CIRUGÍAS
ELECTIVAS* (más 7 roll-ups) → `source = UNKNOWN_PENDING`, `status = UNKNOWN`. La
respuesta del cliente **no** indicó de dónde sale esta sección. **No** se asume
tabla quirúrgica. Pregunta abierta registrada en
[`docs/CLIENT_PENDING.md`](CLIENT_PENDING.md).

## 32 vs 31

`B1_GRID_TOTAL_VS_CODED_INTERVENTIONS`: la grilla por edad/sexo de la Sección A
suma **32** intervenciones; la suma de códigos de B2 ANEXO / total de la Sección
B suma **31**. Diferencia = **1**.

`status = OPEN_FUNCTIONAL_QUESTION` — **nunca `FAIL`**: no existe una regla
funcional confirmada que obligue a la igualdad. Hallazgo pendiente de
reconciliación con el cliente (¿una intervención sin código? ¿múltiples
procedimientos en un mismo egreso?).

`B1_SECCION_B_PULLS_VS_TOTAL` (31 = 31, `OK`) confirma que la cadena
EGRESOS → B2 subtotales → B1 Sección B propaga bien.

## Candidate golden

`candidate_golden_status = CANDIDATE_PENDING_SOURCE_COMPLETENESS`.

El REMASEP final es candidato a *golden reference*, pero todavía falta:

- el **Medinet original** de julio 2026,
- los **egresos hospitalarios reales** de julio 2026,
- la **tabla quirúrgica** de julio 2026,
- la **confirmación de envío / aprobación** MINSAL,
- un **criterio formal** de aceptación del golden dataset.

## Artefactos — `artifacts/final_remasep_delta_provenance/`

| archivo | contenido |
| --- | --- |
| `workbook_structure_comparison.json` | hojas, orden, dimensiones, recuento de fórmulas, merges, VBA, SHA256; `significant_changes` |
| `formula_expression_changes.csv` | cambios de expresión de fórmula (0; descubierto, no *hardcodeado*) |
| `delta_cells.csv` | todas las celdas cambiadas: `change_kind`, `before_formula` / `formula`, `formula_change_type`, contexto semántico y provenance |
| `direct_input_changes.csv` / `formula_result_changes.csv` | vistas filtradas |
| `delta_dependency_edges.csv` | dependencia estructural entre celdas cambiadas |
| `provenance_map.csv` | una fila por delta: `change_origin`, `source`, `status`, directo/derivado, upstream, `needs_human_review` |
| `source_summary.csv` | recuento por fuente (directo / derivado / expresión reescrita / total / revisión) |
| `source_consistency_checks.csv` | chequeos cruzados; 32-vs-31 = `OPEN_FUNCTIONAL_QUESTION` |
| `control_summary.csv` | estado por hoja de `CONTROL` + causales |
| `summary.json` | totales, `candidate_golden_status` |

## Privacidad

Estos workbooks son outputs REMASEP **agregados**. Aun así los artefactos sólo
llevan coordenadas, etiquetas del formulario, fórmulas y valores agregados —
nunca datos individuales.

## Limitaciones

- La asignación `RESOURCE_CALCULATION` vs. `SURGICAL_TABLE` de la Sección D es una
  **inferencia estructural** por encabezado, no una regla confirmada por columna;
  toda la sección queda marcada para revisión humana.
- La fuente de la Sección E (suspensiones) es **desconocida**.
- El delta es contra **valores cacheados** por Excel, no contra un recálculo.
- No se automatiza ninguna fuente nueva (sin adaptador de egresos, sin lector de
  tabla quirúrgica, sin calculadora de recursos). Sprint 3.4 es **inventario y
  clasificación**, no generación.
