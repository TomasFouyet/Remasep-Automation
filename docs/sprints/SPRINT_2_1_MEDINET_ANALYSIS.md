# Sprint 2.1 — Real Medinet Analysis MVP

## Objetivo

Leer un export **real** de Medinet, validarlo, diagnosticarlo y clasificarlo con
la lógica legacy (las 27 reglas de Sprint 1.3), y mostrarlo en la UI en modo
**REAL** (distinto del modo demo). No genera REMASEP.

## Problema que resolvía

Pasar de la maqueta a datos reales: entender cuántas atenciones hay de verdad en
el archivo, qué campos trae, qué estados, y qué proporción cubren las reglas
legacy.

## Implementación

- `src/remasep/adapters/medinet.py` — `read_medinet(path)`: localiza la hoja de
  datos por **puntuación de encabezados** (no por posición), mapea encabezados a
  nombres semánticos vía alias (sin fuzzy), descarta columnas no reconocidas
  (RUN, nombre…), y **preserva el texto de cada celda tal cual** (sólo normaliza
  encabezados). Sólo `.xlsx`.
- `config/legacy_current_logic_2026/rules.yaml` + `services/legacy_rules.py` —
  las 27 reglas como datos (`status: pending_functional_validation`).
- `src/remasep/services/medinet_analysis.py` — `MedinetAnalysisService.analyze()`
  → `MedinetAnalysisResult`. Detecta `structural_empty_rows`, valida por fila con
  `RecordProblem(row, field, error_code)` (sin valores), valida período,
  reproduce la edad legacy y clasifica en 3 buckets.
- UI: `AppState.analysis_mode` (`DEMO`/`REAL`), badge "ANÁLISIS REAL" vs "DATOS DE
  DEMOSTRACIÓN"; el modo real muestra un resumen diagnóstico y **omite** la
  pantalla de excepciones.

## Archivos principales

- [`src/remasep/adapters/medinet.py`](../../src/remasep/adapters/medinet.py)
- [`src/remasep/services/medinet_analysis.py`](../../src/remasep/services/medinet_analysis.py)
- [`src/remasep/services/legacy_rules.py`](../../src/remasep/services/legacy_rules.py)
- [`config/legacy_current_logic_2026/rules.yaml`](../../config/legacy_current_logic_2026/rules.yaml)
- [`docs/MEDINET_ANALYSIS.md`](../MEDINET_ANALYSIS.md)

## Decisiones técnicas

- **Fila estructuralmente vacía** = todos los campos semánticos reconocidos
  vacíos. No cuenta en `total_records`, no es inválida, no genera problema; sí en
  `structural_empty_rows`. Invariante:
  `physical_rows_examined = structural_empty_rows + total_records`.
- La detección de "vacío" trata el whitespace-only como vacío, pero **nunca muta
  el valor**.
- `legacy_age_years` = `DATEDIF(nac, cita, "Y")` con la excepción legacy
  `nacimiento > atención → 0`.
- `PRESTACION` vacía es **diagnóstico informativo**, no un error.
- Privacidad: ni RUN, ni nombre, ni fecha de nacimiento completa salen del
  servicio.

## Tests

Adaptador (detección de hoja, alias, columnas descartadas), servicio (buckets,
invariantes, período, edad), UI real vs demo. Workbooks sintéticos.

## Resultado sobre workbook de referencia

`GENERACION DATOS REMASEP.xlsx`, hoja `Atenciones - Detalles de citas`, período
**julio 2026**:

| Métrica | Valor |
| --- | ---: |
| `physical_rows_examined` | 2006 |
| `structural_empty_rows` | 642 |
| `total_records` | 1364 |
| `valid_records` | 1364 |
| `invalid_records` | 0 |
| `RecordProblem` | 0 |

**`ESTADO`** (1364 no vacíos): `Atendido` 1336 · `Atención Pausada` 16 ·
`En Sala de Espera` 8 · `En Atención` 4.

**Clasificación legacy:** `AG` 111 · `AH` 45 · `AI` 44 · `AJ` 256 · `AK` 11 ·
`AL` 18 → **485 registros con ≥ 1 match**, 879 válidos no cubiertos por reglas
legacy, ~302 con `PRESTACION` vacía.

## Hallazgos

- **Las 2006 filas del workbook no son 2006 atenciones.** Las atenciones reales
  de julio 2026 son **1364**; las otras 642 son fórmulas `AC:AL` arrastradas.
- **`MODALIDAD`** contiene categorías de **previsión / tramo** (`Fonasa A/B/C/D`,
  `GES …`, `Particular`, …), **no** presencial/telemedicina. El campo no se
  renombra (es el encabezado real de Medinet).
- Las 27 reglas legacy dejan la mayoría de los registros como "válido no
  cubierto"; falta el catálogo completo de prestaciones.

## Limitaciones

- Sólo `.xlsx`; la fila 1 debe ser el encabezado.
- Compatibilidad legacy, **no** equivalencia formal ni mapping oficial (eso
  llega en 2.2–2.5).
- El export Medinet **exacto** de julio 2026 sigue pendiente de obtener del
  cliente (se usó la hoja `Atenciones` del generador como sustituto).

## Estado final

**COMPLETE.** Ingesta real + diagnóstico + clasificación legacy funcionando en la
UI. Preguntas funcionales abiertas en [`docs/CLIENT_PENDING.md`](../CLIENT_PENDING.md).
