# Contrato de input Medinet & alcance de período

Hotfix de producción (cierre Sprint 3.5). Formaliza **qué archivo** es el input
real de la aplicación y corrige un bug de alcance en el análisis mensual.

## 1. Qué es input y qué no

| rol | archivo | uso |
| --- | --- | --- |
| **INPUT DE PRODUCCIÓN** | export directo de Medinet **"Detalle de citas"** — p.ej. `detalle_citas - 2026-09-07T123630.940.xlsx` | **único** input real de la app. El usuario lo descarga de Medinet y lo carga. |
| **REFERENCIA LEGACY** | `GENERACION DATOS REMASEP.xlsx` | **sólo** ingeniería inversa (Sprints 1–3.5). Es el workbook con el que se elaboraba el REMASEP a mano. **La app no depende de él para funcionar.** |

El export "Detalle de citas" trae **~47 columnas** (incluidas RUN, nombre, teléfono,
etc.). El adaptador (`remasep.adapters.medinet`) conserva **sólo** las 10
columnas semánticas y descarta el resto al leer; nada personal entra al core.

## 2. Alcance de período — el bug y el fix

Un export Medinet contiene **varios meses** (el archivo real de referencia:
enero–julio 2026). El análisis:

- detectaba correctamente los registros dentro/fuera del período;
- pero ejecutaba la **clasificación legacy y los diagnósticos sobre TODOS los
  registros válidos**, no sólo los del mes.

Eso es incorrecto para un REMASEP **mensual**. Ahora se separan explícitamente:

| concepto | definición |
| --- | --- |
| `physical_rows_examined` | filas físicas del archivo |
| `structural_empty_rows` | filas de fórmula arrastrada (todas las celdas semánticas vacías) — se ignoran |
| `total_records` | filas con al menos un campo semántico |
| `valid_records` / `invalid_records` | estructuralmente válidos / con DIA·SEXO·TIPO faltante o inválido |
| `records_in_period` / `records_outside_period` | de los **válidos**, dentro / fuera del mes-año |
| **`processing_scope_records`** | **`valid_records` ∩ `records_in_period`** — el ÚNICO subconjunto sobre el que se calcula el REMASEP mensual |

**Todas** las métricas y clasificaciones mensuales (`legacy_counts`,
`legacy_matches_total`, `non_target_records`, desglose de edad; y, en el futuro,
`SemanticMetric` y los target mappings) se calculan **sólo** sobre
`processing_scope_records`.

Los registros de otros períodos **no se eliminan ni se modifican**: simplemente
quedan fuera del cálculo. El mensaje de UI cambió de
*"…fuera del período (no se descartan)"* (con estilo de advertencia) a:

> *"9.387 registros pertenecen a otros períodos y no se incluirán en el REMASEP
> de Julio 2026."* — sin tratarlos como error.

## 3. Reconciliación con la referencia legacy (privacy-safe)

`scripts/validate_medinet_production_input.py` +
`src/remasep/services/medinet_input_reconciliation.py`.

Para el archivo real de julio 2026:

```
export directo "Detalle de citas"
  raw_file_records          11 501
  in-period (Julio 2026)     2 114
  out-of-period              9 387
  processing_scope_records   2 114

referencia legacy (hoja de detalle de GENERACION DATOS REMASEP.xlsx)
  registros activos          1 364   (todos de julio 2026)

diferencia                     750
```

Comparación por **multiset** de huellas semánticas (`DIA_CITA ·
FECHA_NACIMIENTO · SEXO · SUCURSAL · ESPECIALIDAD · TIPO_DE_CITA · PRESTACION ·
ESTADO · MODALIDAD · PRESTACION_REALIZADA`) — **sin** RUN, nombre, teléfono,
id de paciente. Se manejan huellas repetidas como cuentas, no como conjunto.

| comparación | resultado |
| --- | --- |
| huella **sin ESTADO** (9 campos) | los **1 364 legacy son un subconjunto multiset EXACTO** de los 2 114 del export directo. `direct_export_excess = 750`. |
| huella **con ESTADO** (10 campos) | 1 363 casan; **1 registro difiere** de ESTADO (el export directo es del 2026-09-07, posterior al snapshot legacy; un ESTADO progresó `En Sala de Espera → Atendido`). |

## 4. Los 750 "extras" — evidencia por ESTADO

Agregados privacy-safe de los 750 registros del export directo de julio que el
proceso antiguo dejó fuera (`extra_records_by_status.csv` etc.):

| ESTADO | registros |
| --- | ---: |
| Cancelado | 545 |
| No Se Presenta | 182 |
| Agendado | 11 |
| Confirmado | 7 |
| Re-Agendado | 5 |
| **total** | **750** |

## 5. Hipótesis de ESTADO — **NO es una regla implementada**

`config/medinet_period_scope_2026/estado_processing_hypothesis.yaml`
(`status: CURRENT_LEGACY_BEHAVIOR_PENDING_FUNCTIONAL_CONFIRMATION`).

El generador legacy, en julio 2026, **conserva**
`{Atendido, Atención Pausada, En Sala de Espera, En Atención}` y **descarta**
`{Cancelado, No Se Presenta, Agendado, Confirmado, Re-Agendado}`:

```
export directo Julio 2026, ESTADO ∈ kept       = 1 337 + 16 + 7 + 4 = 1 364
                            ESTADO ∈ excluded  = 545 + 182 + 11 + 7 + 5 = 750
legacy activo                                   = 1 364
```

El subconjunto de estados "kept" **reproduce EXACTAMENTE** los 1 364 registros
activos del detalle legacy.

> Esto es **evidencia reproducible del comportamiento observado**, no una regla
> funcional confirmada por Fundación Gantz / MINSAL. El alcance de procesamiento
> del REMASEP mensual sigue siendo, **hoy**, sólo por período. El filtro por
> ESTADO queda documentado para confirmación funcional del cliente antes de
> implementarse.

## Artefactos — `artifacts/medinet_production_input_validation/`

`summary.json` · `period_scope_summary.csv` · `legacy_subset_comparison.csv` ·
`extra_records_by_status.csv` / `_by_appointment_type` / `_by_branch` /
`_by_modality` / `_by_specialty` · `README.md`. Ninguna fila individual, ningún
dato personal.
