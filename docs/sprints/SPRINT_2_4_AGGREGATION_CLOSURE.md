# Sprint 2.4 — Legacy Aggregation Dependency Closure

## Objetivo

Cerrar las 102 fórmulas que Sprint 2.3 dejó `UNSUPPORTED` (`COUNTIFS(...) - F30`)
mediante un **grafo de dependencias** entre celdas agregadas, y medir el universo
completo de celdas Medinet-dependientes.

## Problema que resolvía

Las 102 restaban una celda agregada de la misma hoja: no era un problema de
parser sino la ausencia de un modelo de dependencias. Hacía falta resolver esas
referencias por valor **sin hardcodear coordenadas** y descubrir si hay más
fórmulas downstream.

## Implementación

- `legacy_aggregation.py` gana `parse_derived_formula` — AST explícito
  `Σ (signo · COUNTIF(S)) ± (signo · celda) ± (signo · literal)`. Distingue
  **referencia como criterio** (`COUNTIFS(rango, $A26)`, se resuelve a texto, no
  es arista) de **referencia como valor** (`… - F30`, sí es arista). Sin `eval()`.
- `scripts/close_legacy_aggregations.py` — construye el grafo (nodo = `(hoja,
  celda)`), calcula el **cierre Medinet transitivo** hasta punto fijo, detecta
  ciclos (Kahn), profundidad y propaga `depends_on_age` transitivamente. Evalúa
  en orden topológico: primero las base, luego las derivadas sustituyendo sus
  `value_refs`.
- Salida a `artifacts/legacy_aggregation_closure/` (`metric_nodes.csv` = DAG de
  métricas: `metric_id`, `kind`, `dependencies`, `value`, `depends_on_age`).

## Archivos principales

- [`src/remasep/services/legacy_aggregation.py`](../../src/remasep/services/legacy_aggregation.py)
  (`parse_derived_formula`, `arithmetic_constructs`)
- [`scripts/close_legacy_aggregations.py`](../../scripts/close_legacy_aggregations.py)
- [`docs/LEGACY_AGGREGATION_CLOSURE.md`](../LEGACY_AGGREGATION_CLOSURE.md)

## Decisiones técnicas

- **Clasificación de nodo**: `BASE_AGGREGATION`, `DERIVED_AGGREGATION`,
  `DOWNSTREAM_TOTAL` (`SUM`), `VALIDATION` (`IF`), `UNSUPPORTED_OTHER`.
- Los ciclos se marcan `CYCLE` y **nunca** se resuelven en silencio; una
  referencia de valor a algo no evaluable → `MISSING_DEPENDENCY`.
- `DOWNSTREAM_TOTAL` y `VALIDATION` se **inventarían pero no se evalúan** en este
  sprint (llega en 2.5).

## Tests

`parse_derived_formula` (criterio vs. valor, sumas, literales, rechazo de
`* / & SUM IF`), orden topológico, detección de ciclos, dependencia faltante,
propagación de `depends_on_AF`, privacidad, `metric_id` estable.

## Resultado sobre workbook de referencia

**Universo Medinet-dependiente: 2043 celdas** = **1427 directas** (referencian
`Atenciones`) + **616 transitivas** (344 `SUM` + 272 `IF`; ninguna es una nueva
agregación derivada).

| kind | celdas | evaluadas |
| --- | ---: | ---: |
| `BASE_AGGREGATION` | 1325 | 1325 / 1325 |
| `DERIVED_AGGREGATION` | 102 | **102 / 102** |
| `DOWNSTREAM_TOTAL` | 344 | 0 (inventariadas) |
| `VALIDATION` | 272 | 0 (inventariadas) |
| `UNSUPPORTED_OTHER` | 0 | — |

- **Las 102 son 1 familia conceptual**: `COUNTIFS(detalle) − celda agregada de la
  misma hoja` (2 variantes estructurales: corte de edad único vs. banda de edad).
  Todas en `REMASEP_OD`; cada una resta una celda base distinta (`F30`, `G30`…).
- **DAG**: profundidad máx. **4** en el grafo completo (cadenas `SUM` de `SUM`),
  **1** en el subgrafo base → derivada. **Ciclos: 0. Dependencias faltantes: 0.**
- **Cache (base + derived)**: 1263 `MATCH`, 164 `CACHE_DIFFERENCE`, 0
  `CACHE_UNAVAILABLE`. Las 164 diferencias dependen **todas** de `AF` (142 base +
  22 derivadas). 1999 / 2043 celdas son `depends_on_AF`.
- `formula_support_status: PASS` (base + las 102).

## Hallazgos

- **El universo no termina en las 1427 directas**: hay 616 celdas downstream
  (`SUM`/`IF`) que también dependen de Medinet — se midió, no se asumió.
- La resta `COUNTIFS(...) - F30` propaga la obsolescencia de la cache de `AF` de
  ambos términos: 22 derivadas nuevas con `CACHE_DIFFERENCE`.

## Limitaciones

- Grafo derivado del **texto de la fórmula** (analizador A1); sin
  `INDIRECT`/`OFFSET`/named ranges (no presentes).
- Sólo el subset aritmético **observado** (`+`, `-`, celda, literal).
- Equivalencia con el workbook **≠** validación funcional MINSAL.

## Estado final

**COMPLETE.** BASE + DERIVED cerrados (1427 / 1427); `DOWNSTREAM_TOTAL` y
`VALIDATION` inventariados para Sprint 2.5.
