# Sprint 2.2 — Legacy Derived Equivalence (AC:AL)

## Objetivo

Demostrar que la implementación Python reproduce **exactamente** las 10
transformaciones derivadas `AC:AL` de la hoja de detalle del workbook legacy,
celda a celda, para cada atención real.

## Problema que resolvía

Antes de reproducir agregaciones o mapping oficial hay que probar la capa base:
que `RAW → AC:AL` en Python coincide con lo que Excel dejó cacheado.

## Implementación

- `src/remasep/services/legacy_transform.py` — módulo de producción reutilizable:
  `legacy_derived_record(rec, ruleset)`, `legacy_age_years`, `excel_str`.
  Reproduce Excel **exactamente**: concatenación literal sin separador (`""` para
  blancos), propagación de `#N/A` del `IFS`, conteo de `COUNTIF` (un registro
  puede sumar 2+), `DATEDIF "Y"` con `nacimiento > atención → 0`, parseo
  `dd/mm/aaaa`.
- `src/remasep/core/text.py` — **`normalize_legacy_text`** = `str(v).casefold()`
  (`None → ""`): case-insensitive, **preserva tildes y whitespace exactamente**.
  Es la única semántica legacy; la comparten `LegacyRuleSet` y `legacy_transform`.
  `normalize_text` global (tolerante: quita tildes, colapsa espacios) **no** se
  usa para lógica legacy y no se modificó.
- `read_medinet` deja de recortar (`.str.strip()`) el texto de las celdas para
  que adaptador → servicio → reglas usen el mismo texto que el workbook.
- `scripts/compare_legacy_derived.py` — abre el workbook con `data_only=False` y
  `data_only=True`, compara `AC:AL` cacheado vs. Python por fila activa. Salida a
  `artifacts/legacy_equivalence/` (`mismatches.csv` = sólo
  `row/column/status/type`, sin valores).

## Archivos principales

- [`src/remasep/services/legacy_transform.py`](../../src/remasep/services/legacy_transform.py)
- [`src/remasep/core/text.py`](../../src/remasep/core/text.py)
- [`scripts/compare_legacy_derived.py`](../../scripts/compare_legacy_derived.py)
- [`docs/LEGACY_EQUIVALENCE.md`](../LEGACY_EQUIVALENCE.md)

## Decisiones técnicas

- **Semántica legacy:** case-insensitive, **accent-sensitive**, **whitespace-exact**.
  `"CONTROL EVOLUCIÓN DENTARIA"` = `"control evolución dentaria"`, pero ≠
  `"CONTROL EVOLUCION DENTARIA"` y ≠ `"CONTROL  EVOLUCIÓN DENTARIA"`.
- `exact_equivalence` (strings idénticos) vs. `semantic_equivalence` (iguales
  sólo tras normalizar) — este último se reporta como `SEMANTIC_MISMATCH`, no
  como equivalencia.
- Los literales de salida (`"consulta médico"`, …) y los patrones `AJ:AL` viven
  en `rules.yaml`.

## Tests

Reproducción exacta de cada transformación, semántica de matching (tildes,
whitespace, `~` escape), la excepción de edad, el `#N/A` del `IFS`, privacidad
del `mismatches.csv`.

## Resultado sobre workbook de referencia

`GENERACION DATOS REMASEP.xlsx`, hoja `Atenciones` (2006 físicas = 642 vacías +
**1364 activas**), 13 640 comparaciones (1364 × 10):

| Columnas | Exactas / activas | Estado |
| --- | --- | --- |
| `AC AD AE` | 1364 / 1364 (100 %) | **PASS** |
| `AG AH AI` | 1364 / 1364 (100 %) | **PASS** |
| `AJ AK AL` | 1364 / 1364 (100 %) | **PASS** |
| `AF` (edad) | 793 / 1364 (58 %) | **CACHE_UNAVAILABLE** |

**0 mismatches reales.** `overall_status = CACHE_UNAVAILABLE`.

`legacy_category_hits`: `AG` 111 · `AH` 45 · `AI` 44 · `AJ` 256 · `AK` 11 ·
`AL` 18; `records_with_any_legacy_match` = **485**.

## Hallazgos

- **Cache de `AF` obsoleta**: en 571 filas Excel dejó `0` cacheado aunque las
  fechas son válidas y la edad debería ser > 0 → la fórmula `AF` no fue
  recalculada antes del último guardado (`STALE_CACHE_SUSPECTED`). En las 793
  filas donde la cache de `AF` es fiable, Python coincide al 100 %.
- No se puede afirmar `PASS` global de `AF` contra esta cache, pero **no se
  detectó ningún desacuerdo** Python ↔ workbook.
- `sum(AG..AL hits) ≠ records_with_any_legacy_match` en general (una atención
  puede activar varias categorías); en este dataset no hay filas multi-categoría,
  pero no se asume exclusividad.

## Limitaciones

- Depende de los valores **cacheados** por Excel.
- Sólo cubre `AC:AL`, no la agregación ni el mapping oficial.
- Equivalencia con el workbook **≠** validación oficial MINSAL.

## Estado final

**COMPLETE.** La capa `RAW → AC:AL` es reproducible en Python; la cache de `AF`
queda documentada como no verificable.
