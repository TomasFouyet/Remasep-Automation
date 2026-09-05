# Sprint 2.5 — Legacy Downstream Equivalence

## Objetivo

Evaluar las 616 celdas transitivas que Sprint 2.4 dejó inventariadas — **344
`SUM` (`DOWNSTREAM_TOTAL`) + 272 `IF` (`VALIDATION`)** — sobre el mismo grafo de
dependencias, cerrando las **2043** celdas Medinet-dependientes conocidas.

## Problema que resolvía

Sprint 2.4 cerró la agregación base y derivada pero dejó fuera los totales `SUM`
y las validaciones `IF`. Faltaba probar que también son reproducibles a partir de
las mismas celdas ya evaluadas, sin `eval()` ni releer la cache como atajo.

## Implementación

- `legacy_aggregation.py` gana `parse_sum_formula` y `parse_if_formula`
  (AST explícito, sin `eval()`).
  - **SUM**: rango rectangular de la misma hoja (`SUM(Q86:Q92)`) o cadena de
    celdas con `+` (`SUM(F40+H40+…)`), más `± celda ± literal`. Nada más.
  - **IF**: `IF(izq OP der, then, else)` con `OP ∈ {= <> < <= > >=}`; ramas =
    literal numérico/texto, referencia de celda o `IF` anidado (1 nivel, la única
    composición observada).
- `scripts/evaluate_legacy_downstream.py` — llama a
  `close_legacy_aggregations()` (reusa grafo, orden topológico, ciclos,
  `depends_on_age`) y evalúa `DOWNSTREAM_TOTAL`/`VALIDATION` en **orden
  topológico real** (no por fases). `Node` gana `evaluation_status`; salida a
  `artifacts/legacy_downstream_equivalence/`.

## Archivos principales

- [`src/remasep/services/legacy_aggregation.py`](../../src/remasep/services/legacy_aggregation.py)
  (`parse_sum_formula`, `parse_if_formula`)
- [`scripts/evaluate_legacy_downstream.py`](../../scripts/evaluate_legacy_downstream.py)
- [`docs/LEGACY_DOWNSTREAM_EQUIVALENCE.md`](../LEGACY_DOWNSTREAM_EQUIVALENCE.md)

## Decisiones técnicas

- `evaluation_status` ∈ `SUPPORTED` / `DEPENDENCY_UNAVAILABLE` / `UNSUPPORTED_SUM`
  / `UNSUPPORTED_IF` / `UNSUPPORTED_FORMULA` / `CYCLE` — **separado** de
  `cache_comparison_status`.
- Los valores se resuelven de los **nodos ya evaluados del DAG**; una celda del
  rango sin fórmula cuenta como `0` (como Excel).
- Algunas condiciones `IF` comparan contra un **campo de ingreso manual** del
  formulario (no una fórmula, p.ej. `AO9` = "digite EMBARAZADAS"): se lee su
  valor **crudo** (es una entrada, no una métrica), nunca su cache como
  sustituto de un cálculo.
- `celda = ""` usa `normalize_legacy_text` (celda vacía cuenta como `""`; `0` no).
- Las 272 `VALIDATION` **siguen clasificadas `VALIDATION`** — se evalúan y
  reportan, no se interpretan como reglas clínicas.

## Tests

`parse_sum_formula` (rango, cadena `+`, `± celda`, blancos = 0, variantes no
soportadas), `parse_if_formula` (los 6 operadores, ref-vs-literal, resultado
literal/celda, `celda=""`, `IF` anidado, variantes no soportadas), y end-to-end:
`SUM` de `SUM`, `IF` que depende de `SUM`/`DERIVED`, `DEPENDENCY_UNAVAILABLE`,
ciclo, propagación de `depends_on_AF`, cache MATCH/DIFFERENCE/UNAVAILABLE,
privacidad, `metric_id` estable.

## Resultado sobre workbook de referencia

| kind | total | evaluadas | cache MATCH | cache DIFF | cache UNAVAIL |
| --- | ---: | ---: | ---: | ---: | ---: |
| `BASE_AGGREGATION` | 1325 | 1325 | 1183 | 142 | 0 |
| `DERIVED_AGGREGATION` | 102 | 102 | 80 | 22 | 0 |
| `DOWNSTREAM_TOTAL` | 344 | 344 | 260 | 84 | 0 |
| `VALIDATION` | 272 | 272 | 136 | 0 | 136 |
| **Total** | **2043** | **2043** | **1659** | **248** | **136** |

- **`formula_support_status: PASS`** — las 2043 celdas conocidas se evalúan.
  0 `UNSUPPORTED`, 0 `DEPENDENCY_UNAVAILABLE`.
- **SUM: 344 / 344** (5 patrones `formula_pattern`). **IF: 272 / 272** (4 patrones).
- **DAG**: profundidad máx. **4**, 0 ciclos, 0 dependencias faltantes.
- **Las 248 `CACHE_DIFFERENCE` dependen todas de `AF`** (142 base + 22 derivadas +
  84 totales que suman celdas con edad).
- **136 `CACHE_UNAVAILABLE`**, todas `VALIDATION` de resultado **texto**: se
  verificó celda a celda que Excel no guardó valor (`<v>` ausente). Las 136 de
  resultado **numérico** sí están cacheadas y coinciden al 100 %.
- **0 validaciones "activas"** (valor Python ≠ `0`/`""`) para este dataset:
  ninguna de las 272 comprobaciones dispararía su aviso.
- Constructs secundarios en las 616 transitivas: sólo `SUM` e `IF`.

## Hallazgos

- El subset `SUM`/`IF` **observado** cubre el 100 % de las 616 celdas; nada queda
  `UNSUPPORTED` en el workbook de referencia.
- Las validaciones de mensaje son la única fuente de `CACHE_UNAVAILABLE`: Excel
  no cacheó sus resultados de texto.

## Limitaciones

- Comparación contra la cache de Excel (no se recalcula); obsoleta para `AF`,
  ausente para las 136 celdas de mensaje.
- Universo = el **conocido** por el grafo de referencias A1.
- Equivalencia numérica con el workbook **≠** validación funcional MINSAL.

## Estado final

**COMPLETE.** Cierre completo de la lógica de fórmulas legacy derivada de
Medinet: **2043 / 2043 celdas evaluables**.
