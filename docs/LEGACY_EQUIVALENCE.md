# Equivalencia de la lógica derivada legacy (Sprint 2.2)

`scripts/compare_legacy_derived.py` + `src/remasep/services/legacy_transform.py`.

## Qué se compara

Solo la capa **RAW → AC:AL** del workbook `GENERACION DATOS REMASEP.xlsx`:

```
datos crudos (D, G, H, K, M, O, AA, ...)
        │
        ├─► columnas AC:AL calculadas por Excel (valores cacheados en el .xlsx)
        │
        └─► columnas AC:AL calculadas por Python (remasep.services.legacy_transform)
```

Para cada **fila real de atención** (excluyendo las
`structural_empty_rows` de Sprint 2.1) se comparan celda a celda los 10 valores
derivados.

**No** se comparan todavía las hojas de agregación (`REMASEP 01`, `B2 ANEXO`,
`REMASEP_OD`), ni el mapping a la plantilla oficial. No se ejecuta Excel, COM ni
LibreOffice: se leen los valores que Excel dejó cacheados (`data_only=True`).

## Por qué

Antes de traducir la agregación o el mapping oficial hay que demostrar que la
implementación Python reproduce **exactamente** las transformaciones derivadas del
workbook actual. Es una prueba de compatibilidad técnica, no de corrección
funcional.

## Las 10 transformaciones (`legacy_transform.py`)

| Col | Fórmula legacy | Implementación |
| --- | --- | --- |
| `AC` | `=O & K` | `TIPO_DE_CITA + SUCURSAL` (concatenación literal, sin separador) |
| `AD` | `=O & H` | `TIPO_DE_CITA + SEXO` |
| `AE` | `=AA & M & H` | `PRESTACION + ESPECIALIDAD + SEXO` |
| `AF` | `=IF(G>D, 0, DATEDIF(G, D, "Y"))` | `legacy_age_years()` — años completos; `nacimiento > atención → 0` |
| `AG` `AH` `AI` | `=IFS(O = literal, salida, ...) & H` | igualdad de texto **case-insensitive, con tildes** (`normalize_legacy_text`); sin match el `IFS` da `#N/A` y `& H` lo propaga → valor de celda `"#N/A"` |
| `AJ` `AK` `AL` | `=(Σ COUNTIF(AA, "*patrón*")) & H` | nº de patrones presentes (substring **case-insensitive, con tildes**) `+ SEXO` |

`equals` y `contains` usan la **misma** función `remasep.core.text.normalize_legacy_text`:
case-insensitive (`casefold`), preserva diacríticos y **preserva el whitespace
exactamente** (no recorta ni colapsa — Excel tampoco). Es la única semántica
legacy: la comparten `LegacyRuleSet` (bucketing) y `legacy_transform` (valores
AC:AL). `normalize_text` global (que elimina tildes y colapsa espacios) **no** se
usa para reproducir lógica legacy y no se modificó.

Detalles reproducidos exactamente: concatenación de strings vacíos, casing,
tildes, el `#N/A` de `IFS`, el conteo de `COUNTIF` (un registro puede sumar 2+),
la edad y su excepción legacy. **No** se "mejora" la lógica.

Los literales de salida de `AG/AH/AI` (`"consulta médico"`, `"control odon esp"`,
`"evaluacion odon esp"`) y los patrones de `AJ:AL` viven en
`config/legacy_current_logic_2026/rules.yaml`.

## Definición de PASS

Por columna (`column_summary.csv` → `status`):

- **PASS** — todas las filas activas comparadas son **exactas**.
- **FAIL** — hay al menos un desacuerdo real Python ↔ cache (`VALUE_MISMATCH` /
  `SEMANTIC_MISMATCH` / `PY_NONE`).
- **CACHE_UNAVAILABLE** — sin desacuerdos, pero alguna fila no tiene un valor
  cacheado utilizable.

`overall_status` (`summary.json`): `PASS` solo si **todas** las columnas son
`PASS`; `FAIL` si alguna es `FAIL`; `CACHE_UNAVAILABLE` en el resto.

`exact_equivalence` = strings idénticos. `semantic_equivalence` = distintos pero
iguales tras normalizar tildes/casing → se reporta como `SEMANTIC_MISMATCH`, **no**
como equivalencia exacta.

## Privacidad

`mismatches.csv` contiene **solo** `row_number`, `derived_column`,
`comparison_status`, `mismatch_type`. Nunca `expected_value` / `actual_value`
(los valores `AC:AL` incluyen sexo y prestación). El diagnóstico con valores
concretos solo existe en memoria y en los tests sintéticos.

## Múltiples categorías por registro

`sum(AG..AL hits)` **no** tiene por qué ser igual a
`records_with_any_legacy_match`: una atención puede activar más de una categoría
(p.ej. `AJ` y `AK` a la vez). `summary.json → legacy_category_hits` mantiene
ambos números por separado.

## Limitaciones

- Depende de los **valores cacheados** por Excel; si Excel no recalculó una
  celda antes de guardar, ese valor no es fiable (ver `AF` en el resultado de
  referencia).
- Solo cubre `AC:AL` de la hoja de detalle, no la agregación ni el mapping
  oficial.
- El parseo de fechas de `AF` asume el orden `dd/mm/aaaa` (locale es-CL).

## Resultado del workbook de referencia

`GENERACION DATOS REMASEP.xlsx`, hoja `Atenciones - Detalles de citas`
(SHA256 `fc2e1536…d15b79`): **2006 filas físicas = 642 estructuralmente vacías +
1364 activas**. 13 640 comparaciones (1364 × 10).

| columna | exactas / activas | estado |
| --- | --- | --- |
| `AC` `AD` `AE` | 1364 / 1364 (100 %) | **PASS** |
| `AG` `AH` `AI` | 1364 / 1364 (100 %) | **PASS** |
| `AJ` `AK` `AL` | 1364 / 1364 (100 %) | **PASS** |
| `AF` | 793 / 1364 (58 %) | **CACHE_UNAVAILABLE** |

**0 mismatches reales.** Para `AF`, en 571 filas Excel dejó `0` cacheado aunque,
con las fechas presentes y válidas, la edad debería ser > 0: la fórmula `AF` no
fue recalculada por Excel antes del último guardado (`mismatch_type =
STALE_CACHE_SUSPECTED`). En las 793 filas donde la cache de `AF` es fiable
(edades > 0 y ceros genuinos) Python coincide al 100 %.

`overall_status = CACHE_UNAVAILABLE` — no se puede afirmar `PASS` global porque
`AF` no es completamente verificable contra esta cache, pero **no se detectó
ningún desacuerdo** entre Python y el workbook legacy.

## Aviso

Equivalencia con el **workbook actual** ≠ **validación oficial MINSAL**. Las
transformaciones reproducen `GENERACION DATOS REMASEP.xlsx`, que sigue pendiente
de validación funcional.
