# Análisis de Medinet (Sprint 2.1)

Este documento describe el análisis **real** de un export Medinet:
`remasep.adapters.medinet` + `remasep.services.medinet_analysis`.

Alcance de este sprint: **leer, validar, diagnosticar y clasificar con lógica
legacy**. No genera REMASEP, no escribe Excel, no usa COM ni macros.

---

## 1. Esquema esperado

El archivo debe ser `.xlsx`. La hoja de datos se **detecta por sus encabezados**
(no por posición): se elige la hoja cuya fila 1 reconoce más campos requeridos.
La fila 1 son encabezados; los datos empiezan en la fila 2.

El adaptador trabaja con **nombres semánticos**, nunca con letras de columna
(`D`, `G`, `H`, `K`, `M`, `O`, `AA`). El encabezado se normaliza con
`normalize_field_name` (MAYÚSCULAS, sin tildes, espacios → `_`) y luego se mapea:

| Campo semántico | Encabezados aceptados (tras normalizar) |
| --- | --- |
| `DIA_CITA` | `DIA_CITA`, `FECHA_CITA`, `FECHA_DE_CITA`, `FECHA_ATENCION` |
| `FECHA_NACIMIENTO` | `FECHA_NACIMIENTO`, `FECHA_DE_NACIMIENTO` |
| `SEXO` | `SEXO` |
| `SUCURSAL` | `SUCURSAL` |
| `ESPECIALIDAD` | `ESPECIALIDAD` |
| `TIPO_DE_CITA` | `TIPO_DE_CITA`, `TIPO_CITA` |
| `PRESTACION` | `PRESTACION` |
| `ESTADO` *(opcional)* | `ESTADO`, `ESTADO_CITA` |
| `MODALIDAD` *(opcional)* | `MODALIDAD` |
| `PRESTACION_REALIZADA` *(opcional)* | `PRESTACION_REALIZADA` |

Sin fuzzy matching. Cualquier columna no reconocida (RUN, nombre, teléfono…) se
**descarta al leer**.

**Solo se normaliza el encabezado, nunca el valor de la celda.** `read_medinet`
preserva el texto de cada celda tal cual llega de Excel (sin `strip` ni colapso
de espacios), para que la clasificación legacy use exactamente el mismo texto que
el workbook. `None`/`NaN` se representan como `""`; las fechas se parsean a
`datetime` (`dd/mm/aaaa`). La detección de "vacío" (validación,
`structural_empty_rows`) sí considera vacía una celda de solo whitespace, pero es
**solo detección**: no muta el valor almacenado.

### Campos requeridos

`DIA_CITA`, `FECHA_NACIMIENTO`, `SEXO`, `SUCURSAL`, `ESPECIALIDAD`,
`TIPO_DE_CITA`, `PRESTACION`.

Si falta alguno → `SourceValidationError` con el nombre semántico del campo
ausente y la hoja analizada.

### Campos opcionales / diagnósticos

`ESTADO`, `MODALIDAD`, `PRESTACION_REALIZADA`. Si faltan, el análisis continúa y
se listan en `missing_optional_columns`.

---

## 2. Filas estructuralmente vacías

El workbook legacy arrastra sus fórmulas `AC:AL` más allá del conjunto real de
atenciones, dejando filas físicas sin ningún dato de fuente. Estas se detectan
**antes** de validar:

> Una fila es **estructuralmente vacía** solo cuando **todos** los campos
> semánticos reconocidos están vacíos: los 7 requeridos y también
> `ESTADO` / `MODALIDAD` / `PRESTACION_REALIZADA` si están presentes.

Una fila estructuralmente vacía:

- **no** cuenta en `total_records`;
- **no** es un `invalid_record`;
- **no** genera `RecordProblem`;
- sí se contabiliza en `structural_empty_rows` (diagnóstico).

Si **al menos uno** de esos campos tiene información, la fila es un registro
candidato y se valida normalmente (nunca se elimina en silencio), aunque esté
incompleta.

Invariantes:

```
physical_rows_examined = structural_empty_rows + total_records
total_records          = valid_records + invalid_records
```

Confirmado en Sprint 2.1 contra `GENERACION DATOS REMASEP.xlsx`
(hoja `Atenciones - Detalles de citas`, julio 2026):
`physical_rows_examined = 2006`, `structural_empty_rows = 642`,
`total_records = 1364`. Es decir, **las 2006 filas del workbook no son 2006
atenciones**: las atenciones reales son 1364.

## 3. Validación de registros

Ninguna fila candidata se elimina en silencio. Por fila se generan problemas
estructurados `RecordProblem(row_number, field, error_code, message)`; `message`
**nunca** contiene valores de la fila.

| `error_code` | Significado | ¿Invalida la fila? |
| --- | --- | --- |
| `DIA_CITA_INVALID` | `DIA_CITA` vacía o no parseable como fecha | Sí |
| `FECHA_NACIMIENTO_INVALID` | `FECHA_NACIMIENTO` presente pero no parseable | Sí |
| `SEXO_EMPTY` | `SEXO` vacío | Sí |
| `TIPO_DE_CITA_EMPTY` | `TIPO_DE_CITA` vacío | Sí |
| `BIRTH_AFTER_SERVICE` | `FECHA_NACIMIENTO > DIA_CITA` (edad legacy = 0) | No |

`PRESTACION` puede estar vacía: no hay evidencia funcional de que sea obligatoria
para toda atención. Se cuenta en `prestacion_empty_records` como **diagnóstico
informativo** (no `RecordProblem`, no `ValidationResult` de warning global).

### Período

Se usa el mes/año elegido en la UI. Para los registros **válidos**:
`records_in_period`, `records_outside_period`, `min_service_date`,
`max_service_date`. Los registros fuera de período **no se descartan**; se emite
un `ValidationResult` `warning` (o `error` si no hay ningún registro del período).

### Edad — `legacy_age_years(birth, service)`

Años completos cumplidos a la fecha de atención, equivalente a
`DATEDIF(fecha_nacimiento, dia_cita, "Y")`.

**Compatibilidad legacy:** si `fecha_nacimiento > dia_cita`, el workbook actual
devuelve `0`; esta función reproduce ese comportamiento. Devuelve `None` si falta
alguna de las dos fechas. Los conteos agregados son `age_computed`,
`age_legacy_zero`, `age_missing` (el valor individual nunca sale del servicio).

---

## 4. Reglas legacy

Configuración versionada: [`config/legacy_current_logic_2026/rules.yaml`](../config/legacy_current_logic_2026/rules.yaml).
Las fórmulas de Excel **no** se ejecutan como strings; se representan como datos.

> Estas reglas reproducen la lógica observada en `GENERACION DATOS REMASEP.xlsx`
> (columnas AG:AL) y están **pendientes de validación funcional**.
> **No son reglas oficiales MINSAL validadas.**

| Código | Categoría técnica | Campo | Operador |
| --- | --- | --- | --- |
| `AG` | consultas médicas | `TIPO_DE_CITA` | `equals` |
| `AH` | controles odontológicos de especialidad | `TIPO_DE_CITA` | `equals` |
| `AI` | evaluaciones odontológicas de especialidad | `TIPO_DE_CITA` | `equals` |
| `AJ` | controles ortodoncia | `PRESTACION` | `contains` |
| `AK` | controles ortopedia prequirúrgica | `PRESTACION` | `contains` |
| `AL` | instalaciones ortopedia | `PRESTACION` | `contains` |

El matching usa `remasep.core.text.normalize_legacy_text` (**case-insensitive**,
**con tildes**, whitespace exacto — la misma semántica que reproduce el
workbook Excel, ver [`LEGACY_EQUIVALENCE.md`](LEGACY_EQUIVALENCE.md)).
`equals` = igualdad normalizada exacta; `contains` = el valor normalizado de la
regla aparece dentro del `TIPO_DE_CITA`/`PRESTACION` normalizado. Un mismo módulo
lo comparten `LegacyRuleSet` y `legacy_transform`. 27 valores en total
(3 + 4 + 6 + 6 + 3 + 5).

---

## 5. `matched` vs `non-target` vs `invalid`

Las 27 reglas cubren **solo una parte** del universo Medinet. Que un registro no
active ninguna regla **no** significa "prestación desconocida".

| Estado | Definición |
| --- | --- |
| **matched_legacy_rule** | registro válido que activa ≥ 1 categoría AG:AL (`legacy_matches_total`) |
| **valid_but_not_targeted** | registro válido que no activa ninguna (`non_target_records`) |
| **invalid_record** | registro que falla la validación estructural (`invalid_records`) |

Invariante: `valid_records == legacy_matches_total + non_target_records` y
`total_records == valid_records + invalid_records`.

Los `non_target_records` **no** generan excepciones de UI.

---

## 6. Diagnósticos funcionales

Para `ESTADO`, `MODALIDAD`, `PRESTACION_REALIZADA` (si existen):

- presente / ausente;
- nº de registros no vacíos;
- nº de valores únicos;
- para `ESTADO` y `MODALIDAD`: distribución categórica agregada (`valor → conteo`),
  hasta 40 valores únicos — no identifica pacientes;
- para `PRESTACION_REALIZADA`: **solo** presencia y conteo (texto libre; no se
  listan los textos completos).

---

## 7. Privacidad

El servicio procesa los valores en memoria, pero **no escribe** en logs, CSV,
excepciones, UI ni tests: RUN, nombre, dirección, teléfono, correo, identificador
individual ni fecha de nacimiento completa. Los `RecordProblem` referencian solo
`fila`, `campo` y `error_code` con mensajes estáticos.

---

## 8. Limitaciones

- Solo `.xlsx` por ahora (no `.csv`, `.xls`, `.xlsm`).
- Se asume que la fila 1 es el encabezado (sin filas de título previas).
- Las fechas string se interpretan con `dayfirst=True` (formato chileno DD/MM/AAAA).
- Un registro puede activar más de una categoría legacy; se cuenta en cada una,
  pero `legacy_matches_total` lo cuenta una sola vez.
- La lista de `problems` se trunca a 500 (se anota en `notes`).
- El análisis es de **compatibilidad legacy**: no hay equivalencia formal probada
  contra el workbook, ni mapping hacia la plantilla oficial.

---

## 9. Preguntas funcionales pendientes

Confirmar con el responsable funcional de Fundación Gantz / MINSAL:

1. **`ESTADO`** — existe con categorías propias (p.ej. `Atendido`, `Anulado`,
   `En Sala de Espera`) pero no participa en las dependencias REMASEP observadas.
   ¿Debe filtrarse por estado antes de contar?
2. **`MODALIDAD`** — hecho observado: en el workbook de referencia la columna con
   encabezado `MODALIDAD` contiene categorías de previsión (`Fonasa A/B/C/D`,
   `GES …`, `Particular/Libre Elección`, `FFAA`, …), por lo que **no parece
   representar presencial/telemedicina**. El campo **no se renombra** (es el
   encabezado real de Medinet). Pregunta: confirmar que presencial/telemedicina
   debe derivarse mediante `SUCURSAL`, como hace actualmente
   `AC = TIPO_DE_CITA + SUCURSAL`.
3. **`PRESTACION_REALIZADA`** — existe pero las reglas actuales usan `PRESTACION`.
   ¿Cuál es la columna de verdad para clasificar?
4. **Cobertura legacy** — las 27 reglas dejan la mayoría de los registros como
   `valid_but_not_targeted`. Falta el catálogo completo de prestaciones y su
   mapeo a categorías REMASEP.
5. **Cálculo de edad** — ¿la excepción `nacimiento posterior a la atención → 0`
   es intencional o un bug tolerado del workbook?
