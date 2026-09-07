# Technical Overview

Vista técnica transversal del proyecto REMASEP Automation. Complementa
[`docs/ARCHITECTURE.md`](ARCHITECTURE.md) (componentes) con el stack, las
fronteras, la privacidad, las semánticas de texto, el motor legacy y las fuentes
de datos.

---

## 1. Stack

| Área | Herramienta | Notas |
| --- | --- | --- |
| Lenguaje | Python ≥ 3.11 | `src/` layout, paquete `remasep` |
| UI | PySide6 | escritorio; tests en modo *offscreen* |
| Datos | pandas | frames de Medinet |
| Excel (lectura) | openpyxl | inspección, `data_only` F/T, `keep_vba`, `read_only` |
| Tests | pytest | 379 tests; workbooks sintéticos, sin `data/local/` |
| Lint | ruff | `ruff check .` limpio |
| Excel (escritura) *(futuro)* | pywin32 / Excel COM | sólo Windows, generación oficial |
| Empaquetado *(futuro)* | PyInstaller | `.exe` para el piloto |

## 2. Fronteras

- **openpyxl** se usa sólo para **inspeccionar / leer** workbooks. Nunca escribe
  el archivo de referencia.
- **Excel COM** (Windows) queda para la **generación oficial**: abrir copia,
  validar estructura, escribir, recalcular, revisar `CONTROL`, guardar, cerrar
  sólo la instancia creada. Todavía no implementado.
- **No** hay base de datos, **no** hay nube, **no** hay telemetría. Todo local.
- **No** se usa LibreOffice como sustituto de Microsoft Excel para la validación
  final.

## 3. Privacidad

- Procesamiento **local**.
- Los archivos reales viven en `data/local/` y están **gitignored** (contienen
  filas de pacientes).
- Ninguna herramienta exporta PII: los artefactos y CSV contienen sólo fórmulas,
  coordenadas, nombres de columna, metadatos y **valores agregados**.
- `RecordProblem` referencia sólo `fila`, `campo`, `error_code` — nunca el valor.
- Los logs futuros no deben contener RUN, nombre ni fecha de nacimiento
  individuales.

## 4. Input de producción, filas estructurales y alcance de período

**Input de producción** = el export directo de Medinet **"Detalle de citas"**
(`detalle_citas - …xlsx`). `GENERACION DATOS REMASEP.xlsx` es sólo una
**referencia legacy** de ingeniería inversa; la app **no** depende de él. Ver
[`docs/MEDINET_INPUT_CONTRACT.md`](MEDINET_INPUT_CONTRACT.md).

Un export Medinet trae **varios meses**. El REMASEP mensual se calcula SOLO sobre
`processing_scope_records` = **registros estructuralmente válidos ∩ del mes/año
seleccionados**. Los registros de otros períodos **no se descartan** del archivo;
sólo quedan fuera del cálculo (clasificación legacy, edades y —a futuro—
`SemanticMetric`).

`GENERACION DATOS REMASEP.xlsx`, hoja `Atenciones - Detalles de citas`, período
**julio 2026** (SHA256 `fc2e1536…d15b79`):

```
physical_rows_examined = 2006
structural_empty_rows  =  642   (fórmulas AC:AL arrastradas más allá de las atenciones)
active records         = 1364   (= 1364 válidos, 0 inválidos)
```

Export directo "Detalle de citas" de julio 2026 (`3331d61c…`):

```
raw_file_records         = 11 501
in-period (Julio 2026)    =  2 114   -> processing_scope_records
out-of-period             =  9 387
```

Los 1 364 registros activos del detalle legacy son un **subconjunto multiset
exacto** (huella semántica sin ESTADO) de los 2 114 del export directo. La
diferencia de 750 se explica por ESTADO: el generador legacy conserva
`{Atendido, Atención Pausada, En Sala de Espera, En Atención}` (= 1 364 exactos) —
**hipótesis con evidencia**, marcada `CURRENT_LEGACY_BEHAVIOR_PENDING_
FUNCTIONAL_CONFIRMATION`, **no** implementada como regla.

**Las 2006 filas físicas NO son 2006 atenciones.** Invariante:
`physical_rows_examined = structural_empty_rows + total_records`.

## 5. Semánticas de texto

| Función (`remasep.core.text`) | Comportamiento | Uso |
| --- | --- | --- |
| `normalize_text` | tolerante: MAYÚSCULAS, **sin tildes**, espacios colapsados | normalización semántica general (encabezados, etc.) |
| `normalize_legacy_text` | `str(v).casefold()` (`None → ""`): case-insensitive, **con tildes**, **whitespace exacto** | reproducir la lógica legacy del workbook (`LegacyRuleSet`, `legacy_transform`, criterios `COUNTIF`) |

`normalize_legacy_text` es la **única** semántica de matching legacy; reproduce
lo que hace Excel (`=` y `COUNTIF "*x*"` no quitan tildes ni recortan espacios).

## 6. Motor legacy (`src/remasep/services/`)

| Módulo | Sprint | Reproduce |
| --- | --- | --- |
| `legacy_transform.py` | 2.2 | `AC:AL` — `CONCAT`, `IFS` con `#N/A`, `DATEDIF "Y"` (edad, `nac > cita → 0`), `COUNTIF` "contiene" |
| `legacy_rules.py` + `config/legacy_current_logic_2026/rules.yaml` | 2.1 | las 27 reglas de clasificación como datos (`status: pending_functional_validation`) |
| `legacy_aggregation.py` | 2.3 → 2.5 | `COUNTIF`/`COUNTIFS` + `+`; `± celda ± literal` (`parse_derived_formula`); `SUM` (`parse_sum_formula`); `IF` (`parse_if_formula`) |
| `medinet_analysis.py` | 2.1 | ingesta, validación, buckets, diagnóstico |

Constructs Excel implementados: `AC:AL`, `COUNTIF`, `COUNTIFS`, `+`, `-`, `SUM`,
`IF` (`= <> < <= > >=`, `IF` anidado 1 nivel). **Sin `eval()`**: AST explícito.
Cualquier construct fuera del subset observado → `UNSUPPORTED_*`, nunca inferido.

## 7. DAG de dependencias

Construido a partir del **texto de las fórmulas** (analizador A1 de
`scripts/formula_refs.py`) sobre `REMASEP 01`, `B2 ANEXO`, `REMASEP_OD`.

- **Nodo:** `(hoja, celda)`; `metric_id = LEGACY::<hoja>::<celda>`.
- **Universo Medinet-dependiente:** 2043 celdas (1427 directas + 616 transitivas).
- Clasificación de nodo: `BASE_AGGREGATION` (1325) · `DERIVED_AGGREGATION` (102) ·
  `DOWNSTREAM_TOTAL` (344) · `VALIDATION` (272) · `UNSUPPORTED_OTHER` (0).
- **Profundidad máx.: 4. Ciclos: 0. Dependencias faltantes: 0.**
- `evaluation_status` (`SUPPORTED` / `DEPENDENCY_UNAVAILABLE` / `UNSUPPORTED_SUM`
  / `UNSUPPORTED_IF` / `UNSUPPORTED_FORMULA` / `CYCLE`) se mantiene **separado**
  de `cache_comparison_status`.

## 8. Limitaciones de la cache

La equivalencia se compara contra el **valor cacheado** por Excel
(`data_only=True`), no contra un recálculo. En el workbook de referencia:

- **`AF` (edad) obsoleta**: 571 filas con `0` cacheado pese a fechas válidas →
  todas las `CACHE_DIFFERENCE` (248 en el cierre completo) **dependen de `AF`**.
  Nunca se convierten en `MATCH`; se anotan.
- **136 `CACHE_UNAVAILABLE`**: validaciones `IF` de resultado **texto** para las
  que Excel no guardó valor (`<v>` ausente).

**No se afirma "100 % equivalente a Excel".** Se separan:

- **FORMULA SUPPORT** — 2043 / 2043 evaluables (`formula_support_status: PASS`).
- **CACHE CONSISTENCY** — `DIFFERENCES` (248 diferencias por `AF` + 136 sin cache).

## 9. Estado de la plantilla oficial

`REMASEP 2026_V1.4.xlsm` inventariada (Sprint 1.4): 11 hojas, 11 466 fórmulas,
VBA presente (no interpretado), todo protegido salvo `CONTROL`. Sólo hay
**candidatos** de alineación por etiqueta (230 exact, 1189 ambiguos). **El
mapping semántico oficial NO está implementado.**

## 10. Fuentes de datos (Data Sources)

El REMASEP se arma de **varias** fuentes, no sólo Medinet. Confirmado con el
cliente (ver [`docs/CLIENT_PENDING.md`](CLIENT_PENDING.md)):

| Fuente | Aporta | Destino en el REMASEP |
| --- | --- | --- |
| **MEDINET** | consultas / atención ambulatoria (subgrafo legacy actual, 2043 celdas) | `REMASEP 01`, `B2 ANEXO`, `REMASEP_OD` (parte ambulatoria) |
| **EGRESOS HOSPITALARIOS** | **previsión** → REMASEP 01 (Quirófano); **edad + sexo** → REMASEP B1; **códigos de intervenciones quirúrgicas** → B2 ANEXO | segunda fuente de **producción estructurada**, no sólo complemento |
| **TABLA QUIRÚRGICA / RESOURCE CALCULATION** | capacidad teórica (días hábiles × ~7–8 h/día) y utilización real de pabellón (días/horas reales de ocupación) | secciones de recursos/utilización |
| **OFFICIAL REMASEP TEMPLATE** | — | **output**: `REMASEP 2026_V1.4.xlsm` (formato MINSAL) |

Todavía **no** hay código para egresos ni para el cálculo de recursos: falta
recibir ejemplos reales y confirmar reglas.

## 11. Modelo futuro de métricas — *provenance*

El modelo semántico de Fase 3 deberá **conservar la fuente** de cada métrica.
Conceptualmente:

```
source ∈ { MEDINET, EGRESOS, RESOURCE_CALCULATION }
```

Es un **requisito de diseño**, no un enum ni código a implementar ahora: cuando
una métrica llegue a la plantilla oficial hay que poder decir de qué fuente
proviene (para trazabilidad, revisión de excepciones y `CONTROL`).
