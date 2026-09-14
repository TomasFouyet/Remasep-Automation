# Validación histórica del motor MEDINET — Abril–Julio 2026

## Estado

**Cerrado.** Dos validaciones independientes, ambas de sólo lectura, sin
modificar ninguna lógica productiva:

1. **AF/edad vs Microsoft Excel real (COM)** — `legacy_age_years` reproduce
   **exactamente** la fórmula legacy `=IF(FECHA_NACIMIENTO>DIA_CITA, 0,
   DATEDIF(FECHA_NACIMIENTO, DIA_CITA, "Y"))`: **5 799/5 799 pares únicos
   MATCH (100.00 %), 0 mismatch, 0 error**. La implementación Python de edad
   fue contrastada **independientemente contra Microsoft Excel**, no contra
   los valores AF cacheados del workbook legacy.
2. **Regresión histórica del pipeline productivo vs 4 REMASEP reales**
   (Abril–Julio 2026) — **equivalence rate 87,70 %** (3 936/4 488). El 100 %
   de los `REAL_MISMATCH` (552/552) dependen de la columna derivada AF, lo
   que es **completamente consistente** con la obsolescencia de AF ya
   documentada en `docs/LEGACY_AGGREGATION_EQUIVALENCE.md` — no con un defecto
   del motor de clasificación/agregación.

---

## Fase A — AF/edad vs Microsoft Excel real (COM)

### Objetivo

Demostrar que `legacy_age_years` (`src/remasep/services/legacy_transform.py`)
reproduce exactamente la fórmula Excel legacy, usando **Microsoft Excel real**
como referencia — no los valores AF cacheados de
`GENERACION DATOS REMASEP.xlsx` (esos ya se sabían obsoletos).

### Método

- Script: `scripts/validate_af_age_excel_com.py`.
- Fuente: mismo `detalle_citas` actual
  (`data/local/detalle_citas - 2026-09-07T123630.940.xlsx`).
- Alcance: registros **estructuralmente válidos** de Abril, Mayo, Junio y
  Julio 2026 (`processing_scope_frame`, motor productivo, reutilizado sin
  modificar).
- Se extraen los pares únicos `(FECHA_NACIMIENTO, DIA_CITA)` (deduplicados
  para minimizar llamadas COM) de los registros que tienen ambas fechas
  (`FECHA_NACIMIENTO` es opcional en el scope estructural; sin ella no hay
  nada que comparar).
- Puente Excel: subproceso `powershell.exe` (funciona tanto en Windows nativo
  como desde WSL vía interop) que crea una instancia Excel **propia y
  aislada** (`New-Object -ComObject`, nunca `GetActiveObject`), añade un
  workbook **en memoria** (nunca abre ni guarda un archivo), escribe sólo las
  fechas necesarias, calcula `=IF(A>B,0,DATEDIF(A,B,"Y"))` y devuelve por
  stdout únicamente los enteros resultantes. Las fechas viajan sólo por stdin
  del subproceso, nunca se escriben a un archivo intermedio ni a un log. Al
  terminar: `Workbook.Close(False)` (sin guardar) y `Excel.Quit()` de su
  propia instancia, en un bloque `finally` — nunca se mata un proceso Excel.

### Resultado

| métrica | valor |
| --- | ---: |
| pares únicos comparados | **5 799** |
| MATCH | **5 799** |
| MISMATCH | **0** |
| errores / unavailable | **0** |
| equivalencia | **100.0000 %** |

No hubo ningún mismatch: no aplica ningún diagnóstico adicional ni cambio a
`legacy_age_years`.

### Tests

`tests/test_validate_af_age_excel_com.py`:

- extracción de pares (dedup, exclusión de registros sin `FECHA_NACIMIENTO`,
  scope de mes/año, multi-mes) — puro Python, corre siempre en el suite Linux
  habitual, sin Excel;
- `python_expected_ages` delega en `legacy_age_years` (no reimplementa nada);
- `compare_ages`: clasificación MATCH/MISMATCH/error y que el detalle de un
  mismatch **nunca** expone fechas (sólo índice + valores enteros);
- fallo limpio (`RemasepError`) cuando `powershell.exe` no está disponible.
- Integración real contra Excel: marcada `@pytest.mark.excel`, requiere la
  variable de entorno `REMASEP_VALIDATE_AF_EXCEL_COM=1` (opt-in explícito,
  igual que el resto de integración COM del proyecto en
  `test_excel_com_integration.py`) — nunca se dispara sola en un `pytest`
  normal, aunque el entorno de desarrollo actual sí puede automatizar Excel
  vía WSL interop.

---

## Fase B — regresión histórica del pipeline productivo vs REMASEP reales

### Fuente y alcance

- Medinet: único export directo disponible,
  `data/local/detalle_citas - 2026-09-07T123630.940.xlsx` — no existen
  exports Medinet históricos separados por mes.
- Períodos: **Abril, Mayo, Junio y Julio de 2026**, cada uno recortado por la
  lógica productiva (`processing_scope_frame` + filtro ESTADO `CONFIRMED`)
  sobre el mismo `detalle_citas`.
- Referencia: los 4 REMASEP históricos reales correspondientes
  (`REMASEP 2026 V1.4 Abril 2026.xlsm`, `2026-5 REMASEP_V1.4.xlsm`,
  `REMASEP 2026 V1.4 Junio 2026.xlsm`, `REMASEP_V1.4 Julio 2026.xlsm`),
  abiertos siempre en modo sólo lectura (`data_only=True, read_only=True`),
  con verificación de SHA256 antes/después de leer.
- Comparación restringida a las **1 122 celdas del `write_manifest`** por mes
  (automatización basada en `detalle_citas`); egresos hospitalarios, tabla
  quirúrgica y recursos/capacidad quedan fuera de alcance.
- Script: `scripts/compare_historical_remasep.py` (reutiliza
  `build_production_pending_writes` y `compare_reference_values` tal cual
  están; no reimplementa ningún motor de cálculo).

### Resultado agregado (4 488 comparaciones = 4 × 1 122)

| categoría | cantidad | % |
| --- | ---: | ---: |
| EXACT_MATCH | 3 932 | 87,61 % |
| ZERO_VS_BLANK_EQUIVALENT | 4 | 0,09 % |
| **equivalencia (MATCH + ZERO_VS_BLANK)** | **3 936** | **87,70 %** |
| REAL_MISMATCH | 552 | 12,30 % |
| REFERENCE_UNAVAILABLE | 0 | 0,00 % |

### Por mes

| mes | referencia | EXACT_MATCH | ZERO_VS_BLANK | REAL_MISMATCH | UNAVAILABLE | equivalence_rate |
| --- | --- | ---: | ---: | ---: | ---: | ---: |
| Abril 2026 | REMASEP 2026 V1.4 Abril 2026.xlsm | 994 | 1 | 127 | 0 | 88,68 % |
| Mayo 2026 | 2026-5 REMASEP_V1.4.xlsm | 984 | 2 | 136 | 0 | 87,88 % |
| Junio 2026 | REMASEP 2026 V1.4 Junio 2026.xlsm | 974 | 1 | 147 | 0 | 86,90 % |
| Julio 2026 | REMASEP_V1.4 Julio 2026.xlsm | 980 | 0 | 142 | 0 | 87,34 % |

**Invariante de los cuatro meses**: en cada mes, `pending_writes = 1 122`,
`completeness.ok = True` y `estado_filter_status = CONFIRMED` — ninguno quedó
`BLOCKED`. Las 4 comparaciones se hicieron sobre el mismo universo de 1 122
celdas, sin excepciones ni casos especiales por mes.

### Por hoja destino

| hoja | compared | EXACT_MATCH | ZERO_VS_BLANK | REAL_MISMATCH | equivalence_rate |
| --- | ---: | ---: | ---: | ---: | ---: |
| REMASEP_OD | 3 808 | 3 409 | 0 | **399** | 89,52 % |
| REMASEP 01 | 660 | 507 | 0 | **153** | 76,82 % |
| B2 ANEXO | 20 | 16 | 4 | **0** | 100,00 % |

`B2 ANEXO` tiene equivalencia perfecta (0 mismatches) — es, además, la única
hoja cuyas métricas **no dependen de AF** (ver más abajo).

### Los 552 REAL_MISMATCH dependen de AF — sin excepción

| eje | mismatches (de 552) | % |
| --- | ---: | ---: |
| **AF / edad** | **552** | **100,0 %** |

- Población que depende de AF: 4 468 comparaciones → **552 mismatches (12,35
  %)**.
- Población que **no** depende de AF: 20 comparaciones → **0 mismatches
  (0,00 %)**.
- Es decir: **0 de los 20** casos no-AF fallan; **552 de los 4 468** casos
  AF-dependientes fallan. No hay un solo `REAL_MISMATCH` fuera de la
  dependencia de AF en ninguno de los cuatro meses.

Dirección del delta (histórico − calculado) en los 552 mismatches:

| dirección | cantidad |
| --- | ---: |
| referencia > calculado | 128 |
| referencia < calculado | 424 |

### Limitación del método

Se usa un **único export actual de Medinet** (`detalle_citas`), filtrado por
mes, como fuente para los cuatro períodos — no existen exports Medinet
históricos independientes para Abril, Mayo y Junio 2026. Si el detalle
histórico real difería del export actual (altas/bajas de registros
posteriores, correcciones de datos en Medinet), esta comparación no puede
detectarlo: sólo valida la **lógica del motor** sobre el dato disponible hoy,
no la fidelidad histórica del dato fuente en el momento en que se generó cada
REMASEP real.

### Conclusión

Los 552/552 `REAL_MISMATCH` dependiendo de AF, junto con la equivalencia
perfecta (100,00 %) tanto en `B2 ANEXO` (única hoja no-AF) como en el
subconjunto no-AF general (0/20), **son completamente consistentes con la
obsolescencia de AF previamente identificada** (los valores AF cacheados del
workbook legacy están desactualizados frente a la fórmula que el propio
workbook define — ver `docs/LEGACY_AGGREGATION_EQUIVALENCE.md`) y **no hay
evidencia de divergencia fuera de AF** en ninguna de las 4 468 comparaciones
restantes.

Esto **no prueba causalidad absoluta**: el método compara el motor productivo
actual contra un único export Medinet reutilizado para los cuatro meses (ver
limitación arriba), así que no puede excluir por completo otras fuentes de
diferencia en el dato histórico real. Lo que sí establece, con evidencia
directa y sin excepciones: de los 552 mismatches observados, el 100 %
correlaciona con AF, y el 100 % de los casos sin dependencia de AF (20/20,
incluida toda la hoja `B2 ANEXO`) coincide exactamente con el REMASEP
histórico.

Dado que la Fase A confirmó, contra Microsoft Excel real, que
`legacy_age_years` reproduce la fórmula legacy con **100 % de equivalencia**
(5 799/5 799), los 552 mismatches **no pueden atribuirse a un error de la
implementación Python de AF**: la implementación de edad es correcta: la
divergencia está en que los **REMASEP históricos de referencia** fueron
generados con el AF **cacheado** (obsoleto) del workbook legacy, no con la
fórmula que el propio workbook define — exactamente la brecha ya documentada
en `docs/LEGACY_AGGREGATION_EQUIVALENCE.md`.

No se cambió ninguna regla de clasificación, filtro ESTADO ni archivo
histórico a partir de este diagnóstico.

---

## Fase C — cierre técnico

Ejecutado sobre el estado de este diagnóstico (ver también el resumen de
sesión): `pytest`, `ruff check .`, `git diff --check`, `git status`. Resultado
detallado en el mensaje de cierre de la sesión que generó este documento.

## Qué no se versiona

- `historical_regression_detail.csv` / `historical_regression_summary.csv` /
  `summary.json` (contienen conteos por celda; se generan bajo demanda con
  `scripts/compare_historical_remasep.py`, nunca en `artifacts/`, que está en
  `.gitignore`).
- Exports Medinet y workbooks REMASEP históricos (`data/local/`, ya en
  `.gitignore`).
- Cualquier archivo temporal de Excel COM (el puente de la Fase A no escribe
  ninguno: el workbook vive sólo en memoria del proceso Excel y nunca se
  guarda).
- Ningún artifact con PII: ni fechas de nacimiento individuales, ni RUN,
  nombre u otro dato personal aparecen en este documento, en los scripts ni en
  sus salidas — sólo conteos agregados.
