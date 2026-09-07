# Inventario semántico de métricas (Sprint 3.1)

`scripts/inventory_semantic_metrics.py` +
`src/remasep/services/semantic_layout.py`.

Sprint 0–2.5 cerraron la **reproducción matemática** del subgrafo Medinet del
workbook legacy: **2043 / 2043** celdas evaluables (1325 base + 102 derivadas +
344 totales `SUM` + 272 validaciones `IF`), DAG de profundidad 4, 0 ciclos, 0
dependencias faltantes.

Sprint 3.1 empieza la **capa semántica**: sin ampliar el parser de Excel,
descubre *qué significa* cada métrica a partir del **layout visual** del
workbook.

> **Esto NO es un semantic mapping validado.** Son *candidatos* extraídos de los
> rótulos del formulario. No hay códigos de sexo, grupo de edad, actividad ni
> especialidad — eso es Sprint 3.2. Tampoco hay mapping a la plantilla oficial.

---

## Por qué abandonamos progresivamente las coordenadas

La lógica de negocio **no debe conocer `G54`** (principio de
[`ARCHITECTURE.md`](ARCHITECTURE.md)). Hasta ahora cada métrica era un
`metric_id` técnico (`LEGACY::<hoja>::<celda>`) — útil para demostrar
equivalencia, inútil para razonar. Para mapear a la plantilla oficial y para
integrar egresos/recursos hace falta describir la métrica por **lo que un humano
ve**: en qué sección está, qué fila la nombra, qué columna la clasifica.

`metric_id` técnico **sigue siendo el identificador**. `semantic_signature` es
una **firma candidata** que se irá refinando (Sprint 3.2 la traducirá a
dimensiones: sexo, grupo de edad, actividad, especialidad).

---

## `MetricContext`

Por cada métrica (`BASE_AGGREGATION` / `DERIVED_AGGREGATION` /
`DOWNSTREAM_TOTAL`):

| Campo | Contenido |
| --- | --- |
| `metric_id`, `source`, `sheet`, `cell`, `kind`, `dependency_depth`, `value`, `depends_on_AF` | del cierre (Sprint 2.4/2.5). `source = "MEDINET"` |
| `section_label_1..2` | encabezados de sección que abarcan varias filas (raw + normalizado) |
| `row_label_1..3` | jerarquía de fila: grupo (`A.4. …`) + rótulos de la propia fila (código, descripción) |
| `column_label_1..3` | jerarquía de columna: banda + subnivel + hoja/sexo |
| `semantic_signature` | `hoja :: sección… :: fila… :: columna…` (normalizado) — **candidata** |
| `context_status` | `COMPLETE` / `PARTIAL` / `AMBIGUOUS` / `NO_CONTEXT` |

Cada nivel se guarda **raw** (texto original, con tildes y espacios) y
**normalized** (mayúsculas, sin diacríticos, whitespace colapsado). El raw
**nunca** se pierde.

---

## Raw vs. normalized labels

`normalize_semantic_label` (en `semantic_layout.py`) es una normalización
**propia de la capa semántica**, independiente de `core.text` y de
`normalize_legacy_text`:

```
"20 A 24 AÑOS"                 -> "20 A 24 ANOS"
"  Menor de \n10 años "        -> "MENOR DE 10 ANOS"
"5010009 - VIDRIO IONÓMERO"    -> "5010009 - VIDRIO IONOMERO"
```

Sólo facilita el *matching* posterior. El raw se conserva en todas las salidas.

---

## Jerarquía de filas

Se extrae hacia la izquierda de la métrica, preservando el orden:

1. **Encabezado de grupo** más cercano por encima — la fila-rótulo en **negrita**
   entre la sección y la métrica (`A.1. CONSULTAS Y ALTAS`,
   `A.4. ACTIVIDADES DE ODONTOLOGÍA GENERAL`). Sólo el más cercano; no se cruza
   otro grupo ni la sección. Si el workbook tuviera anidamiento más profundo, se
   sub-reporta antes que atribuir mal — «no inventar».
2. **Rótulos de la propia fila** — las celdas de texto a la izquierda de la
   métrica, de izquierda a derecha (código de prestación, descripción,
   especialidad).

Los niveles se llaman `row_label_1 … row_label_3` — **sin nombrarlos
funcionalmente**. En el workbook de referencia la profundidad máxima es **3**.

## Jerarquía de columnas

Se extrae hacia arriba, resolviendo celdas combinadas:

- banda superior (`SEGÚN GRUPOS DE EDAD O DE RIESGO`, `POR SEXO`);
- subnivel (`20-24 años`, `Menos de 1 año - 1 año`);
- hoja final (`Hombres` / `Mujeres` / `EMBARAZADAS` / `MIGRANTES`).

`column_label_1 … column_label_3`, **sin asumir** que significan sexo/edad. En el
workbook de referencia la profundidad máxima es **3**.

## Secciones

Una fila de sección es una fila-rótulo en **negrita**, no repartida por la
grilla, con ≤ 2 valores de texto distintos y su primera celda de texto en las
columnas B–D. Si la hoja usa la palabra **«SECCIÓN»**, ésas son las secciones y
el resto de encabezados en negrita (`A.4. …`) son niveles de fila. Si no
(p.ej. `B2 ANEXO`), los encabezados en negrita (`A. KINESIOLOGÍA`) actúan como
sección. Se detecta anidamiento cuando dos filas de sección están contiguas
(`OTORRINOLARINGOLOGÍA` / `PROCEDIMIENTOS DIAGNÓSTICOS`). Profundidad máxima
observada: **2**. **No se inventa** sección cuando no hay evidencia estructural.

---

## Merged cells

Todos los encabezados se resuelven contra el **ancla** de su celda combinada; no
se hace *fill-forward* destructivo ni se modifica el workbook. Ejemplo real:
`F5:G5 = "Menos de 1 año - 1 año"`, `F6 = "Hombres"` → la métrica `F30` obtiene
`column_labels = ["SEGÚN GRUPOS DE EDAD O DE RIESGO", "Menos de 1 año - 1 año",
"Hombres"]`.

---

## `context_status` — criterios objetivos

| Estado | Criterio |
| --- | --- |
| `COMPLETE` | ≥ 1 rótulo de fila **y** ≥ 1 rótulo de columna, sin ambigüedad |
| `PARTIAL` | contexto de un solo eje (fila **o** columna; puede haber sección) |
| `AMBIGUOUS` | la extracción detectó una ambigüedad — p.ej. la banda de encabezados continúa en una columna vecina pero la de la métrica está vacía |
| `NO_CONTEXT` | no se recuperó ningún rótulo visible |

Las ambigüedades **no se resuelven en silencio**: se anotan en `context_notes` y
la métrica queda `AMBIGUOUS`.

---

## `source` / provenance

`source = "MEDINET"` para **todo** este inventario. El modelo semántico futuro
tendrá también `EGRESOS` y `RESOURCE_CALCULATION` (ver
[`TECHNICAL_OVERVIEW.md`](TECHNICAL_OVERVIEW.md) → *Modelo futuro de métricas*).
En este sprint **no** se implementan esas fuentes ni se inventan métricas suyas.

---

## Evidencia de fórmula

`formula_evidence.csv` extrae, como **evidencia técnica** (no mapping):

- `referenced_detail_columns` — columnas `AC:AL` de la hoja de detalle usadas;
- `text_criteria` — literales con comodín (`"*VIDRIO IONÓMERO*"`, `"*Hombre*"`);
- `age_lower_bound` / `age_upper_bound` — de los criterios `$AF:$AF`
  (`">=20"` + `"<=24"` → `20 / 24`; `"<2"` → `None / 1`);
- `formula_evidence_status` — `AGE_AND_TEXT` / `AGE_BOUNDS` / `TEXT_ONLY` /
  `AGGREGATE_ONLY` (para `SUM`) / `NONE`.

### Consistencia rótulo ↔ fórmula

`label_formula_consistency.csv` compara — **objetivamente**, por solapamiento de
intervalos — el rótulo de columna de edad contra las cotas de la fórmula:

| Estado | Criterio |
| --- | --- |
| `CONSISTENT` | los intervalos de edad (rótulo y fórmula) se solapan |
| `CONFLICT` | los intervalos son disjuntos (p.ej. rótulo «20 A 24» pero `AF >= 25, <= 29`) |
| `NO_COMPARABLE` | uno de los dos lados no aporta edad |

Un `CONFLICT` **se registra, no se corrige**. Sin interpretación clínica.

---

## Validaciones (las 272 `IF`)

Van en `validation_context.csv`, **separadas** del inventario de métricas. Se
guarda su contexto visible, qué campo manual referencian (`EMBARAZADAS` /
`MIGRANTES`, de las columnas `AN` / `AO`), la fórmula y su valor Python. **No se
convierten en métricas** ni se interpreta su significado funcional más allá del
rótulo existente.

---

## Duplicados

`duplicate_semantic_signatures.csv` lista los `metric_id` distintos con la misma
`semantic_signature` (`status = UNRESOLVED`). **No se resuelven
automáticamente** — si aparecieran, indicarían que la firma aún no discrimina lo
suficiente (Sprint 3.2 añadirá dimensiones).

---

## Resultado del workbook de referencia

`GENERACION DATOS REMASEP.xlsx` (sha256 `fc2e1536…d15b79`), julio 2026.

| | valor |
| --- | ---: |
| Métricas candidatas | **1771** (B2 ANEXO 30 · REMASEP 01 335 · REMASEP_OD 1406) |
| por tipo | BASE 1325 · DERIVED 102 · DOWNSTREAM_TOTAL 344 |
| `context_status` | **COMPLETE 1741** · PARTIAL 30 · AMBIGUOUS 0 · NO_CONTEXT 0 |
| firmas semánticas duplicadas | **0** (1771 firmas distintas / 1771) |
| profundidad máx. jerarquía | filas 3 · columnas 3 · secciones 2 |
| métricas con cotas de edad detectables | 1393 |
| métricas con criterios de texto | 1191 |
| rótulo ↔ fórmula | **CONSISTENT 1391** · CONFLICT **0** · NO_COMPARABLE 196 (184 sin edad en ningún lado) |
| validaciones | 272 (todas `PARTIAL`; 136 referencian `EMBARAZADAS`, 136 `MIGRANTES`) |

- Las 30 `PARTIAL` son **todo `B2 ANEXO`**: es una lista plana con una única
  columna de total, sin desglose por edad/sexo.
- **0 `AMBIGUOUS` / 0 `NO_CONTEXT`**: toda métrica de `REMASEP 01` y `REMASEP_OD`
  tiene fila + columna recuperables.
- **0 conflictos rótulo/fórmula**: los rótulos de edad del workbook coinciden con
  sus propios criterios `AF` (se construyeron juntos). El detector de conflictos
  está cubierto por tests sintéticos.

---

## Limitaciones

- Las heurísticas de layout (negrita = encabezado; «SECCIÓN» = sección;
  encabezado de grupo = fila-rótulo en negrita más cercana) están calibradas
  contra **este** workbook. `context_status` refleja la confianza, no una
  garantía.
- Sólo se recuperan hasta 3 niveles de fila, 3 de columna y 2 de sección (el
  máximo observado). Un layout más profundo se truncaría a los niveles más
  externos.
- La normalización semántica es preliminar; conserva puntuación y dígitos.
- **No es un semantic mapping validado.** No hay dimensiones clínicas, ni
  mapping a la plantilla oficial, ni fuentes `EGRESOS`/`RESOURCE_CALCULATION`.

---

## Artefactos

`scripts/inventory_semantic_metrics.py "<workbook>"` →
`artifacts/semantic_metric_inventory/`:

| Archivo | Contenido |
| --- | --- |
| `semantic_metric_candidates.csv` | una fila por métrica: contexto (raw + normalizado), `semantic_signature`, `context_status` |
| `validation_context.csv` | las 272 `IF`, su contexto y qué campo manual referencian |
| `formula_evidence.csv` | evidencia técnica por métrica (columnas de detalle, criterios, cotas de edad) |
| `label_formula_consistency.csv` | comparación objetiva rótulo de edad ↔ cotas de fórmula |
| `duplicate_semantic_signatures.csv` | `metric_id` distintos con la misma firma (sin resolver) |
| `context_coverage.csv` | cobertura de contexto por hoja y tipo |
| `layout_<HOJA>.csv` | volcado del texto resuelto de cada hoja (auditar heurísticas) |
| `summary.json` | totales; `semantic_layer_status: PRELIMINARY_NOT_VALIDATED` |
| `README.md` | resumen de la corrida |

### Privacidad

Sólo rótulos del formulario, coordenadas y valores agregados. Ninguna fila
Medinet individual (RUN, nombre, fecha de nacimiento, sexo/prestación/tipo de
cita por atención).
