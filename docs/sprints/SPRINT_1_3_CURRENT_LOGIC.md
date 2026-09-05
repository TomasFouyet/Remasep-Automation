# Sprint 1.3 — Current Legacy Logic Inventory

## Objetivo

Describir, sólo a partir de hechos observables en las fórmulas, **qué hace** cada
columna derivada `AC:AL` y qué reglas de clasificación aplica, sin traducirlas a
semántica clínica ni afirmar que sean oficiales.

## Problema que resolvía

El grafo de Sprint 1.2 dice qué celda lee cuál; falta caracterizar el *tipo* de
cada transformación y extraer la lista de condiciones de clasificación para poder
representarlas como datos.

## Implementación

`scripts/inventory_current_logic.py` — clasifica cada columna derivada en un tipo
técnico (`CONCAT`, `AGE_DATEDIF`, `EXACT_MAP`, `PATTERN_FLAG`, `UNKNOWN`),
extrae las condiciones de las columnas de clasificación y calcula las columnas
raw que alimentan cada salida de forma transitiva. Genera
[`docs/CURRENT_LOGIC.md`](../CURRENT_LOGIC.md) y CSV en
`artifacts/current_logic/`.

## Archivos principales

- [`scripts/inventory_current_logic.py`](../../scripts/inventory_current_logic.py)
- [`docs/CURRENT_LOGIC.md`](../CURRENT_LOGIC.md)

## Decisiones técnicas

- Los tipos (`CONCAT`, `AGE_DATEDIF`, …) son etiquetas **técnicas**, no
  categorías clínicas.
- Aviso estable en el documento: los conteos son **filas físicas**, no
  atenciones (`structural_empty_rows` pertenece al análisis Medinet).
- `pattern_count > 1` en una columna se marca como advertencia (no ocurre: es 1
  en todas).

## Tests

Workbooks sintéticos: clasificación de tipos de transformación, extracción de
reglas candidatas, cálculo transitivo de columnas raw.

## Resultado sobre workbook de referencia

**Tipos por columna:** `AC AD AE` = `CONCAT`; `AF` = `AGE_DATEDIF`;
`AG AH AI` = `EXACT_MAP` (`IFS` de igualdad sobre `TIPO DE CITA`);
`AJ AK AL` = `PATTERN_FLAG` (`COUNTIF` "contiene" sobre `PRESTACIÓN`). Todas
concatenan `SEXO` al final. `pattern_count = 1` en todas, sin huecos.

**7 columnas raw que alimentan las salidas de forma transitiva:**
`DIA CITA`, `FECHA NACIMIENTO`, `SEXO`, `SUCURSAL`, `ESPECIALIDAD`,
`TIPO DE CITA`, `PRESTACIÓN`.

**27 reglas candidatas** extraídas de las columnas de clasificación:
`AG` 3 · `AH` 4 · `AI` 6 · `AJ` 6 · `AK` 3 · `AL` 5.

## Hallazgos

- Columnas presentes pero **sin uso** en las dependencias REMASEP: `ESTADO` (P),
  `MODALIDAD` (J), `PRESTACIÓN REALIZADA` (AB). `AC` usa `TIPO DE CITA + SUCURSAL`
  (no `MODALIDAD`); las reglas usan `PRESTACIÓN` (no `PRESTACIÓN REALIZADA`).
  Requieren confirmación funcional.

## Limitaciones

- **Compatibilidad legacy ≠ reglas oficiales MINSAL.** Estas 27 reglas
  reproducen lo que hace el workbook actual; están pendientes de validación
  funcional.
- Falta el catálogo completo de prestaciones y su mapeo a categorías REMASEP.

## Estado final

**COMPLETE.** Las 27 reglas se versionan como datos en
`config/legacy_current_logic_2026/rules.yaml` a partir de Sprint 2.1.
