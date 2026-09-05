# Sprint 1.2 — Formula Dependency Analysis

## Objetivo

Descubrir el **grafo de dependencias técnicas** entre celdas y hojas del workbook
generador: qué columna de la hoja de detalle alimenta qué hoja de salida y por
qué fórmulas, usando siempre el **texto original** de la fórmula.

## Problema que resolvía

Para reproducir la lógica hay que saber qué columnas de `Atenciones` son
"raíces" del cálculo y cuáles son derivadas intermedias, y cómo se conectan las
hojas de salida entre sí.

## Implementación

- `scripts/formula_refs.py` — analizador ligero de referencias A1
  (`A1`, `$A$1`, `A1:B10`, `$A:$A`, `'nombre con espacios'!$AF:$AF`), que ignora
  lo que hay dentro de literales de texto y reporta constructs no soportados
  (`INDIRECT`, `OFFSET`, referencias externas, 3D, `#REF!`) en vez de inferir.
- `scripts/analyze_dependencies.py` — recorre las fórmulas, construye
  `derived_columns.csv`, `formula_dependencies.csv`, `sheet_dependencies.csv`,
  `output_column_dependencies.csv` y genera
  [`docs/CURRENT_EXCEL_FLOW.md`](../CURRENT_EXCEL_FLOW.md) con un diagrama Mermaid.

## Archivos principales

- [`scripts/formula_refs.py`](../../scripts/formula_refs.py),
  [`scripts/analyze_dependencies.py`](../../scripts/analyze_dependencies.py)
- [`docs/CURRENT_EXCEL_FLOW.md`](../CURRENT_EXCEL_FLOW.md)

## Decisiones técnicas

- Trabajar sobre la fórmula **original**, nunca sobre la versión normalizada.
- Rangos simbólicos: `A:A` no se expande a celdas.
- Los constructs no soportados se listan, no se resuelven.

## Tests

Workbooks sintéticos: extracción de referencias A1/absolutas/rangos/cross-sheet,
ignorar referencias dentro de strings, detección de constructs no soportados.

## Resultado sobre workbook de referencia

**10 columnas derivadas** en `Atenciones - Detalles de citas`
(`AC AD AE AF AG AH AI AJ AK AL`), cada una con 2006 fórmulas y **1 patrón por
columna** — todas las filas replican la misma fórmula. Todas se calculan a
partir de columnas *locales* de la propia hoja de detalle.

**Grafo entre hojas** (referencias cross-sheet):

| origen | destino | referencias |
| --- | --- | ---: |
| Atenciones - Detalles de citas | REMASEP_OD | 2312 |
| Atenciones - Detalles de citas | REMASEP 01 | 484 |
| Atenciones - Detalles de citas | B2 ANEXO | 24 |

No hay otras aristas cross-sheet. Las tres hojas de salida además se
auto-referencian (roll-ups locales con `SUM`).

## Hallazgos

- Sin `INDIRECT`/`OFFSET`/referencias externas/3D/`#REF!` en el workbook.
- Todo el cálculo baja de `Atenciones` a las 3 hojas de salida; no hay
  dependencias laterales entre `REMASEP 01`, `B2 ANEXO` y `REMASEP_OD`.

## Limitaciones

- Grafo **técnico**, no semántico: describe qué celda lee cuál, no qué mide.
- Analizador de referencias por regex, no un parser completo de Excel (el
  workbook de referencia no define named ranges, así que no hay ambigüedad).

## Estado final

**COMPLETE.** El grafo de dependencias es la base de los cierres de Sprint 2.4 y
2.5.
