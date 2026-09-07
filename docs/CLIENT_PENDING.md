# Información funcional — resuelta y pendiente

Preguntas al responsable funcional de Fundación Gantz / MINSAL. Se separan las
**resueltas** (respuestas recibidas del cliente) de las **pendientes**.

---

## RESUELTAS

### A. Egresos hospitalarios — para qué se usan

El archivo de **egresos hospitalarios** es una **segunda fuente de producción
estructurada** del REMASEP (no sólo un complemento de cirugías). Se usa para:

1. **Previsión** → completar `REMASEP 01` (sección **Quirófano**).
2. **Edad y sexo** → `REMASEP B1`.
3. **Códigos de intervenciones quirúrgicas** → `B2 ANEXO`.

### B. Recursos / capacidad / utilización de pabellones — origen y proceso

**No** provienen del archivo de egresos. El proceso manual informado es:

- **Capacidad disponible**: días hábiles del mes × **aproximadamente 7 a 8
  horas** por día.
- **Ocupación real**: revisar los días reales de ocupación del pabellón usando la
  **tabla quirúrgica** y calcular las horas reales de ocupación.

No se formaliza todavía una fórmula definitiva (falta la tabla quirúrgica real y
varias reglas — ver *Pendientes*).

### C. `2026-7 REMASEP_V1.4.xlsm` recibido previamente — no es el final

El archivo `2026-7 REMASEP_V1.4.xlsm` recibido antes **NO era el REMASEP final**:
le faltaban las cirugías. **No debe considerarse golden reference.**
(`reference_status = INCOMPLETE_PRE_SURGERY`.)

### D. REMASEP julio 2026 terminado — candidato a golden reference

Existe el archivo `REMASEP_V1.4 Julio 2026.xlsm`, que el cliente identifica como
el **REMASEP de julio 2026 terminado**. Inspección técnica hecha en el Sprint 3.4
(ver [`docs/FINAL_REMASEP_DELTA_PROVENANCE.md`](FINAL_REMASEP_DELTA_PROVENANCE.md)):
0 cambios de expresión de fórmula, 113 celdas de valor cambiadas (52 input
directo + 61 propagación), estructura idéntica. `CONTROL = 0`
(`PASS_INTERNAL_VALIDATION`).

**No** consta que se enviara a MINSAL ni que MINSAL lo aceptara, así que se maneja
como `reference_status = FINAL_PER_CLIENT` + `approval_status = UNCONFIRMED`, y
como `candidate_golden_status = CANDIDATE_PENDING_SOURCE_COMPLETENESS` (falta el
Medinet y los egresos reales de julio, la tabla quirúrgica y un criterio formal
de aceptación).

### F. Egresos hospitalarios — destino confirmado por sección

Confirmado (y contrastado con el delta del REMASEP final):

- **Previsión** de cirugía → `REMASEP 01` (Quirófano).
- **Edad + sexo** de cirugía → `REMASEP B1` Sección A (8 inputs directos en el
  archivo final).
- **Códigos de intervención quirúrgica** → `B2 ANEXO` (6 inputs directos en el
  archivo final).

### G. Recursos / pabellones — no salen de egresos

Confirmado: los recursos **no** provienen del archivo de egresos.

- **Capacidad disponible**: días hábiles del mes × **~7 a 8 horas** por día.
- **Ocupación real**: se revisa desde la **tabla quirúrgica** (días y horas
  reales de ocupación).

En el REMASEP final, la Sección D de `REMASEP 01` recibe inputs que —por el texto
de sus encabezados— se clasifican como `RESOURCE_CALCULATION` (capacidad) o
`SURGICAL_TABLE` (horas programadas/ocupadas); la regla exacta **por columna**
sigue sin confirmar.

### H. `EGRESO HOSPITALARIO 2025.pdf` — es el formulario, no una base

El PDF es el **formulario/instructivo** de egreso hospitalario, no una base
mensual de registros. Sigue **pendiente** que el cliente entregue un
archivo/reporte con los egresos reales de julio.

### E. Medinet julio 2026 — no conservado

El cliente **no conserva** la base Medinet original de julio 2026. **Ximena**
puede solicitarla a **Paulo**. Sigue **pendiente** obtener el export Medinet
exacto usado para julio 2026 (hasta ahora se usó la hoja `Atenciones` del
generador como sustituto).

---

## PENDIENTES

### Fuentes de datos a obtener

- Archivo **Medinet original** de julio 2026 (vía Ximena → Paulo).
- Ejemplo real de **egresos hospitalarios**.
- **Tabla quirúrgica** de julio 2026.
- **REMASEP julio 2026 completamente terminado / aprobado** y validación de su
  estado.
- **Criterio de aceptación** del dataset golden.

### Recursos / pabellones

- Regla exacta para usar **7 vs. 8 horas/día**.
- Tratamiento de **feriados**.
- **Cantidad de pabellones** (número de quirófanos en dotación).
- Si **todos** tienen la misma jornada disponible (jornada por pabellón).
- **Fórmula exacta** de porcentaje / utilización.
- Regla **por columna** de la Sección D de `REMASEP 01`: cuáles vienen del
  cálculo de recursos y cuáles de la tabla quirúrgica.

### Egresos

- Identificar las **columnas exactas** de egresos para: previsión, edad, sexo,
  códigos quirúrgicos.
- Manejo de **múltiples intervenciones** por egreso / paciente.
- **Reconciliar 32 vs 31**: la grilla por edad/sexo de `REMASEP B1` Sección A
  tiene una intervención más que la suma de códigos de `B2 ANEXO`.

### Suspensiones de cirugía

- **¿De qué fuente se obtiene la Sección E de `REMASEP 01` (Causas de suspensión
  de cirugías electivas)?** ¿Tabla quirúrgica u otro reporte? (En el REMASEP
  final aparecen inputs directos nuevos sin fuente confirmada — quedan como
  `UNKNOWN_PENDING`.)

### REMASEP julio 2026 — estado

- Confirmar si `REMASEP_V1.4 Julio 2026.xlsm` fue **enviado a MINSAL** y/o
  **aceptado** (hoy `approval_status = UNCONFIRMED`).
- **Criterio formal** de golden reference.

### Medinet — semántica de campos

- Confirmar `SUCURSAL` → presencial / telemedicina (hoy `AC = TIPO_DE_CITA +
  SUCURSAL`).
- Significado funcional de **`ESTADO`** (`Atendido`, `Atención Pausada`,
  `En Sala de Espera`, `En Atención`): ¿filtrar por estado antes de contar?
- Comportamiento esperado de **`PRESTACION`** (puede venir vacía; hoy es
  diagnóstico informativo, no error).
- **`MODALIDAD`**: contiene previsión / tramo (`Fonasa A/B/C/D`, `GES …`,
  `Particular`, …), **no** presencial/telemedicina. El campo no se renombra.
- **`PRESTACION_REALIZADA`** existe pero las reglas usan `PRESTACION` — ¿cuál es
  la columna de verdad para clasificar?
- **Cobertura legacy**: las 27 reglas `AG:AL` clasifican 485 de 1364; falta el
  catálogo completo de prestaciones y su mapeo a categorías REMASEP.
- Cálculo de edad: ¿la excepción `nacimiento posterior a la atención → 0` es
  intencional o un bug tolerado del workbook?

### Plantilla oficial

- Validar los candidatos de alineación (`exact_label` / `ambiguous`) hoja por
  hoja — ninguno se elige automáticamente.
- Confirmar cuáles de los `input_candidates` estructurales son realmente celdas
  de ingreso del proceso.

### Infraestructura (para el piloto)

- Windows + Excel Desktop, políticas de macros, permisos, antivirus /
  SmartScreen, firma del ejecutable, carpeta institucional.

---

## Golden dataset

Por al menos 2–3 meses, con **provenance** por métrica: Medinet · egresos ·
recursos · REMASEP final aprobado, identificando las celdas modificadas a mano.
