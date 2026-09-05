# Sprint 2.3 — Direct Aggregation Equivalence

## Objetivo

Reproducir en Python las fórmulas de conteo de `REMASEP 01`, `B2 ANEXO` y
`REMASEP_OD` que **referencian directamente** la hoja de detalle
(`COUNTIF`/`COUNTIFS` y sumas de ellas), y compararlas con el valor cacheado por
Excel.

## Problema que resolvía

`AC:AL` (Sprint 2.2) es la capa base por atención; el siguiente paso es la primera
capa de agregación: los conteos directos que producen los números del REMASEP.

## Implementación

- `src/remasep/services/legacy_aggregation.py` — evaluador del **subset
  observado**: `COUNTIF(rango, criterio)`, `COUNTIFS(...)`, suma con `+` de esas
  llamadas; rango = columna entera de la hoja de detalle; criterio = literales
  `"..."` y/o referencias de celda de rótulo concatenadas con `&`; criterios
  numéricos `= <> < <= > >=`; comodines `*` `?` con escape `~`. Todo lo demás →
  `UNSUPPORTED` (no se infiere).
- El dataset se arma como en Sprint 2.2: `Atenciones` + `AC:AL` calculado por
  `legacy_transform`, excluyendo `structural_empty_rows`.
- Estados **separados**: `formula_evaluation_status` (¿se pudo evaluar?) vs.
  `cache_comparison_status` (`MATCH` / `CACHE_DIFFERENCE` / `CACHE_UNAVAILABLE`).
- `scripts/compare_legacy_aggregations.py` → `artifacts/legacy_aggregation_equivalence/`.

## Archivos principales

- [`src/remasep/services/legacy_aggregation.py`](../../src/remasep/services/legacy_aggregation.py)
- [`scripts/compare_legacy_aggregations.py`](../../scripts/compare_legacy_aggregations.py)
- [`docs/LEGACY_AGGREGATION_EQUIVALENCE.md`](../LEGACY_AGGREGATION_EQUIVALENCE.md)

## Decisiones técnicas

- Semántica de criterios de texto = `normalize_legacy_text` (case-insensitive,
  accent-sensitive, whitespace-exact), consistente con Sprint 2.2.
- **Sensibilidad al padding**: para cada fórmula se comprueba si una fila
  estructuralmente vacía *podría* contar (valores `AC:AE=""`, `AF=0`,
  `AG:AI=#N/A`, `AJ:AL="0"`).
- Una `CACHE_DIFFERENCE` sólo se llama "obsoleta" si hay evidencia; el único caso
  es `AF`.

## Tests

Criterios (exacto / `*` / `?` / accent / whitespace / `~` / `<>`), numéricos
parametrizados, parsing (single/multi/suma/`&`/cell-ref), unsupported (otra
función, contenido tras el `COUNTIF`, rango no columna entera, otra hoja),
evaluación, padding (caso no sensible + caso sintético sensible), privacidad.

## Resultado sobre workbook de referencia

`GENERACION DATOS REMASEP.xlsx`:

| Hoja | Directas | Soportadas | Unsupported | Cache MATCH | Cache DIFF |
| --- | ---: | ---: | ---: | ---: | ---: |
| REMASEP 01 | 247 | 247 | 0 | 186 | 61 |
| B2 ANEXO | 24 | 24 | 0 | 24 | 0 |
| REMASEP_OD | 1156 | 1054 | 102 | 973 | 81 |
| **Total** | **1427** | **1325** | **102** | **1183** | **142** |

- **Constructs observados**: `COUNTIFS` (1393), `COUNTIF` (34), sumas con `+`,
  rangos de columna entera, criterios de texto con `*`, `&`-concatenación,
  criterios numéricos, criterios por referencia de celda de rótulo. 32 patrones
  distintos.
- **102 unsupported**, todas en `REMASEP_OD`, de la forma `=COUNTIFS(…) - F30`
  (restan una celda de la misma hoja → agregación *derivada*, fuera de alcance
  por diseño en este sprint).
- **34 celdas que NO dependen de `AF` → 34/34 `MATCH`** (la evidencia más limpia
  de equivalencia).
- **Las 142 `CACHE_DIFFERENCE` dependen todas de `AF`** (cache de edad obsoleta,
  ver Sprint 2.2). No se convierten en `MATCH`.
- `formula_support_status: PARTIAL` (por las 102); `cache_consistency_status:
  DIFFERENCES` (por las 142).

## Hallazgos

- La primera capa de agregación es reproducible salvo por la cache de `AF`.
- Las 102 `…-F30` no son un problema de parser: son una capa **derivada** que
  necesita un grafo de dependencias → Sprint 2.4.

## Limitaciones

- Comparación contra el valor **cacheado**, no recalculado.
- Sólo el subset **observado**; otras fórmulas de conteo darían `UNSUPPORTED`,
  no un resultado inventado.
- Equivalencia numérica con el workbook **≠** validación funcional MINSAL.

## Estado final

**COMPLETE** para su alcance (agregaciones directas). Las 102 derivadas se cierran
en Sprint 2.4.
