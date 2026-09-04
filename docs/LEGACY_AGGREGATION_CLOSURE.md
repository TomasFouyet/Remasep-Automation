# Cierre por dependencias de las agregaciones legacy (Sprint 2.4)

`scripts/close_legacy_aggregations.py` +
`src/remasep/services/legacy_aggregation.py` (`parse_derived_formula`).

Sprint 2.3 reprodujo las **agregaciones base** (`COUNTIF`/`COUNTIFS` sobre
`Atenciones - Detalles de citas`) y dejó **102 fórmulas `UNSUPPORTED`** de la
forma `COUNTIFS(...) - F30`: una agregación base **menos** otra celda agregada de
la misma hoja. Sprint 2.4 las reproduce mediante un **grafo de dependencias**
evaluado en orden topológico.

> **Aviso.** Equivalencia con `GENERACION DATOS REMASEP.xlsx` **no** es validación
> oficial MINSAL. Se demuestra únicamente que, dado el mismo conjunto de
> atenciones, la aritmética entre celdas agregadas del workbook y su
> reimplementación en Python coinciden.

## Base vs. derived

| | Definición | Evaluación |
| --- | --- | --- |
| **BASE_AGGREGATION** | `COUNTIF`/`COUNTIFS` (o suma de ellos con `+`) sobre la hoja de detalle, sin más | directa sobre las filas activas (Sprint 2.3) |
| **DERIVED_AGGREGATION** | ≥ 1 `COUNTIF(S)` **y** ≥ 1 referencia de celda usada como **valor** (`… - F30`, `COUNTIF(…) + X7`) o literal | tras resolver la(s) celda(s) referenciada(s) por el grafo |
| **DOWNSTREAM_TOTAL** | sólo `SUM(...)` y/o `+`/`-` entre celdas/rangos, sin `COUNTIF(S)` propio | fuera de alcance de este sprint (inventariada) |
| **VALIDATION** | contiene `IF`/`IFS`/`AND`/`OR`/`IFERROR`/… (comprobaciones, mensajes) | fuera de alcance (inventariada) |
| **UNSUPPORTED_OTHER** | `*`, `/`, `&`, paréntesis de agrupación aritmética, u otra función | fuera de alcance (inventariada) |

El subset aritmético implementado en `legacy_aggregation.py` es **sólo** el
observado: operadores `+` y `-`, referencias de celda como valor, literales
numéricos. `parse_derived_formula` lanza `UnsupportedFormulaError` ante `* / &
SUM IF (...)` — no se implementa nada por inferencia. `arithmetic_constructs()`
inventaría qué construct concreto dejó una fórmula fuera del subset.

### Referencia como criterio ≠ referencia como valor (Parte E)

- `COUNTIFS(rango, $A26)` → `$A26` es un **criterio**: se resuelve a texto contra
  el valor cacheado de esa celda-rótulo (comportamiento de Sprint 2.3, intacto).
  **No** genera arista en el grafo.
- `COUNTIFS(...) - F30` → `F30` es un **valor**: genera la arista `F30 → celda` y
  se resuelve con el valor evaluado de `F30`.

`parse_derived_formula` devuelve ambas listas por separado
(`criterion_refs` vs. `value_refs`).

## Dependency DAG

- **Nodo:** `(hoja, celda)` para `REMASEP 01` / `B2 ANEXO` / `REMASEP_OD`.
- **Arista `A → B`:** la fórmula de `B` necesita el valor de `A`
  (`aggregation_dependency_edges.csv`, `dependency_type`: `SAME_SHEET_VALUE`,
  `SAME_SHEET_RANGE`, `CROSS_SHEET_VALUE`, `CROSS_SHEET_RANGE`).
- **Universo Medinet transitivo:** una celda es Medinet-dependiente si referencia
  directamente `Atenciones - Detalles de citas` **o** depende (por celda o rango)
  de otra celda Medinet-dependiente. Se itera hasta punto fijo.
- **Ciclos:** se detectan (Kahn deja nodos sin ordenar) y se marcan `CYCLE`;
  **nunca** se resuelven en silencio ni se evalúan.
- **Dependencias faltantes:** una referencia de valor a una celda que no es
  `BASE`/`DERIVED` evaluable → el nodo queda `MISSING_DEPENDENCY`, sin evaluar.

## Evaluación topológica

Sin `eval()`. `DerivedFormula` es una estructura explícita
`Σ (signo · COUNTIF(S)) ± (signo · celda) ± (signo · literal)`. El orquestador:

1. evalúa las **métricas base** sobre las filas activas;
2. ordena topológicamente las **derivadas** (Kahn);
3. evalúa cada derivada sustituyendo sus `value_refs` por el valor ya calculado.

`metric_nodes.csv` es la representación técnica resultante
(`metric_id = LEGACY::<hoja>::<celda>`, `kind`, `dependencies`, `value`,
`depends_on_age`) — el objetivo es depender del **DAG de métricas**, no del texto
de la fórmula, en sprints posteriores. Todavía **sin** semántica clínica.

## Evaluator vs. cache · `depends_on_AF` transitivo

Igual que Sprint 2.3, se separan dos estados por celda:

- `formula_evaluation` — ¿se pudo evaluar? (`BASE`/`DERIVED` soportadas).
- `cache_comparison_status` — `MATCH` / `CACHE_DIFFERENCE` / `CACHE_UNAVAILABLE`
  contra el valor cacheado por Excel (`data_only=True`).

La señal `depends_on_age` (referencia a la columna de edad `AF`) se **propaga
transitivamente**: si `F30` depende de `AF`, entonces `F46 = COUNTIFS(...) - F30`
también, aunque su propio `COUNTIFS` no mencione `AF`. Una `CACHE_DIFFERENCE` en
una fórmula `depends_on_age` se anota

> `Formula depends on AF; workbook reference contains known stale AF cache.`

y **no** se convierte en `MATCH` (Sprint 2.2 confirmó cache de `AF` obsoleta).

## Resultado del workbook de referencia

`GENERACION DATOS REMASEP.xlsx` (sha256 `fc2e1536…d15b79`),
2006 filas físicas = 642 estructuralmente vacías + **1364 activas**.

### Universo Medinet-dependiente

| | celdas |
| --- | ---: |
| directas (referencian `Atenciones`) | 1427 |
| transitivas (vía otras celdas) | 616 |
| **total** | **2043** |

Las 616 transitivas son **totales `SUM(...)` (344)** y **celdas de validación
`IF(...)` (272)**; ninguna es una agregación derivada nueva. **El universo no
termina en las 1427 directas** — hay 616 celdas downstream que también dependen
de Medinet, inventariadas pero fuera del alcance de evaluación de este sprint.

### Clasificación y soporte

| kind | celdas | evaluadas |
| --- | ---: | ---: |
| `BASE_AGGREGATION` | 1325 | **1325 / 1325** |
| `DERIVED_AGGREGATION` | 102 | **102 / 102** |
| `DOWNSTREAM_TOTAL` | 344 | 0 (inventariadas) |
| `VALIDATION` | 272 | 0 (inventariadas) |
| `UNSUPPORTED_OTHER` | 0 | — |

- **Las 102 forman 1 familia conceptual** (`COUNTIFS(detalle) − celda agregada de
  la misma hoja`), con **2 variantes estructurales** (un corte de edad `AF,"<2"`
  vs. una banda `AF,">=8" … AF,"<=N"`). Todas en `REMASEP_OD`; cada una resta una
  celda distinta (`F30`, `G30`, …) que es a su vez una agregación base.
- **`formula_support_status: PASS`** — todas las base y las 102 derivadas
  evaluables.
- **Profundidad del DAG:** máx. **1** en el subgrafo evaluable (base → derivada);
  máx. **4** en el grafo Medinet completo (cadenas de `SUM` de `SUM`).
- **Ciclos: 0. Dependencias faltantes: 0.**

### Cache

| | evaluadas (base + derived) |
| --- | ---: |
| `MATCH` | 1263 |
| `CACHE_DIFFERENCE` | 164 |
| `CACHE_UNAVAILABLE` | 0 |

- **Las 34 celdas base que NO dependen de `AF` coinciden 100 % con la cache**
  (idéntico a Sprint 2.3).
- **Las 164 `CACHE_DIFFERENCE` dependen todas de `AF`** y llevan la nota de cache
  obsoleta. De ellas, 142 son base (= Sprint 2.3) y 22 son derivadas: la resta
  `COUNTIFS(...) - F30` propaga la edad obsoleta de ambos términos.
- `cache_consistency_status: DIFFERENCES`.
- 1999 de las 2043 celdas Medinet-dependientes son `depends_on_AF`.

## Qué sigue sin estar soportado

- **`DOWNSTREAM_TOTAL` (344):** `SUM(rango)` y sumas de celdas agregadas.
  Evaluables en un sprint posterior una vez fijado el modelo de métricas
  semántico (sus dependencias ya están en el DAG).
- **`VALIDATION` (272):** `IF(...)` de comprobación / mensajes de aviso. No son
  métricas; quedan como inventario.
- Operadores `* / &`, paréntesis de agrupación aritmética y otras funciones:
  se inventarían (`arithmetic_constructs`), no se implementan.

## Limitaciones

- Comparación contra el **valor cacheado** del workbook (no se recalcula Excel);
  confirmado obsoleto para `AF`.
- El grafo de dependencias se deriva del **texto de la fórmula** (analizador A1
  de `scripts/formula_refs.py`): sin `INDIRECT`/`OFFSET`/named ranges (el
  workbook de referencia no los usa).
- Sólo se reproduce el subset aritmético **observado** (`+`, `-`, celda como
  valor, literal). Cualquier otra cosa → `UNSUPPORTED_OTHER`, nunca inferencia.
- Equivalencia numérica con este workbook **≠** validación funcional MINSAL.

## Artefactos

`scripts/close_legacy_aggregations.py "<workbook>"` →
`artifacts/legacy_aggregation_closure/`:

| Archivo | Contenido |
| --- | --- |
| `unsupported_family_inventory.csv` | las 102 derivadas: nº de `COUNTIF(S)`, refs de misma hoja, operadores, patrón, profundidad candidata |
| `aggregation_dependency_edges.csv` | aristas `origen → destino` + `dependency_type` |
| `aggregation_dependency_summary.csv` | por nodo: grado entrada/salida, profundidad, `status` |
| `transitive_medinet_cells.csv` | toda celda Medinet-dependiente (directa/transitiva), profundidad, columnas raíz, `evaluation_supported` |
| `metric_nodes.csv` | DAG de métricas: `metric_id`, `kind`, `dependencies`, `value`, `depends_on_age` |
| `equivalence.csv` | valor Python vs. cache para `BASE` + `DERIVED`, `cache_comparison_status`, `depends_on_AF`, notas |
| `sheet_summary.csv` | conteos por hoja |
| `summary.json` | totales y estados (`formula_support_status` y `cache_consistency_status` **separados**) |
| `README.md` | resumen de la corrida |

### Privacidad

La unidad de comparación es una **celda agregada**. Los artefactos contienen sólo
fórmulas, coordenadas, nombres de columna y valores agregados. **Ningún** dato
individual de paciente (RUN, nombre, fecha de nacimiento, sexo, prestación ni
tipo de cita por fila).
