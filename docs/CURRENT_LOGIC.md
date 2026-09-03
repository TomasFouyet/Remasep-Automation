# Lógica actual de la hoja de detalle

> Generado por `scripts/inventory_current_logic.py`. Describe **solo hechos observables** en las fórmulas de la hoja de detalle. No traduce a reglas REMASEP, no juzga corrección clínica y no infiere lógica no escrita.
>
> Archivo analizado: `GENERACION DATOS REMASEP.xlsx`  
> SHA256: `fc2e153659041abc50d82d13a84b4a1209e776681257c75b2383328933d15b79`  
> Generado: 2026-09-03T14:07:15Z  
> Hoja de detalle: `Atenciones - Detalles de citas` (2006 filas físicas)  
> Hojas output: `REMASEP 01`, `B2 ANEXO`, `REMASEP_OD`

> **Sobre los conteos:** las cifras de filas y fórmulas de este documento corresponden a **filas físicas** de la hoja y **no deben interpretarse como la cantidad de atenciones reales**. El workbook de referencia arrastra sus fórmulas `AC:AL` más allá de las atenciones, por lo que puede haber filas físicas sin datos. La detección de filas estructuralmente vacías (`structural_empty_rows`) pertenece al análisis Medinet (`remasep.services.medinet_analysis`); para el recuento real del dataset consultar [`docs/MEDINET_ANALYSIS.md`](MEDINET_ANALYSIS.md).

## Esquema de columnas

| col | header | kind | fórmulas | uso directo | uso transitivo |
| --- | --- | --- | ---: | :---: | :---: |
| A | ID CITA | raw | 0 | · | · |
| B | TIPO DE AGENDA | raw | 0 | · | · |
| C | HORA CITA | raw | 0 | · | · |
| D | DIA CITA | raw | 0 | · | sí |
| E | RUN PACIENTE | raw | 0 | · | · |
| F | NOMBRE PACIENTE | raw | 0 | · | · |
| G | FECHA NACIMIENTO | raw | 0 | · | sí |
| H | SEXO | raw | 0 | · | sí |
| I | ASEGURADORA | raw | 0 | · | · |
| J | MODALIDAD | raw | 0 | · | · |
| K | SUCURSAL | raw | 0 | · | sí |
| L | DERIVADOR | raw | 0 | · | · |
| M | ESPECIALIDAD | raw | 0 | · | sí |
| N | PROFESIONAL | raw | 0 | · | · |
| O | TIPO DE CITA | raw | 0 | sí | sí |
| P | ESTADO | raw | 0 | · | · |
| Q | TIPO ASEG. | raw | 0 | · | · |
| R | COPAGO | raw | 0 | · | · |
| S | BONIFICACIÓN | raw | 0 | · | · |
| T | TOTAL | raw | 0 | · | · |
| U | ORIGEN | raw | 0 | · | · |
| V | USUARIO | raw | 0 | · | · |
| W | PAGADO | raw | 0 | · | · |
| X | PROFESIONAL SOLICITANTE | raw | 0 | · | · |
| Y | NUMERO DE ORDEN | raw | 0 | · | · |
| Z | OBSERVACION ORDEN | raw | 0 | · | · |
| AA | PRESTACIÓN | raw | 0 | · | sí |
| AB | PRESTACIÓN REALIZADA | raw | 0 | · | · |
| AC | TIPO CITA / PRESENCIAL O TELEMEDICINA | derived | 2006 | sí | sí |
| AD | TIPO CITA / SEXO | derived | 2006 | sí | sí |
| AE | PRESTACION / ESPECIALIDAD / SEXO | derived | 2006 | sí | sí |
| AF | EDAD | derived | 2006 | sí | sí |
| AG | Consultas Médicas | derived | 2006 | sí | sí |
| AH | Controles Odontologico de especialidad | derived | 2006 | sí | sí |
| AI | Evaluaciones Odontologicas de especialidad | derived | 2006 | sí | sí |
| AJ | Controles Ortodoncia | derived | 2006 | sí | sí |
| AK | Controles Ortopedia preq | derived | 2006 | sí | sí |
| AL | Instalaciones Ortopedia | derived | 2006 | sí | sí |

## Transformaciones derivadas

| col | header | tipo | fuentes | headers fuente | fórmulas | patrones |
| --- | --- | --- | --- | --- | ---: | ---: |
| AC | TIPO CITA / PRESENCIAL O TELEMEDICINA | CONCAT | K\|O | SUCURSAL\|TIPO DE CITA | 2006 | 1 |
| AD | TIPO CITA / SEXO | CONCAT | H\|O | SEXO\|TIPO DE CITA | 2006 | 1 |
| AE | PRESTACION / ESPECIALIDAD / SEXO | CONCAT | H\|M\|AA | SEXO\|ESPECIALIDAD\|PRESTACIÓN | 2006 | 1 |
| AF | EDAD | AGE_DATEDIF | D\|G | DIA CITA\|FECHA NACIMIENTO | 2006 | 1 |
| AG | Consultas Médicas | EXACT_MAP | H\|O | SEXO\|TIPO DE CITA | 2006 | 1 |
| AH | Controles Odontologico de especialidad | EXACT_MAP | H\|O | SEXO\|TIPO DE CITA | 2006 | 1 |
| AI | Evaluaciones Odontologicas de especialidad | EXACT_MAP | H\|O | SEXO\|TIPO DE CITA | 2006 | 1 |
| AJ | Controles Ortodoncia | PATTERN_FLAG | H\|AA | SEXO\|PRESTACIÓN | 2006 | 1 |
| AK | Controles Ortopedia preq | PATTERN_FLAG | H\|AA | SEXO\|PRESTACIÓN | 2006 | 1 |
| AL | Instalaciones Ortopedia | PATTERN_FLAG | H\|AA | SEXO\|PRESTACIÓN | 2006 | 1 |

Tipos técnicos: `CONCAT`, `AGE_DATEDIF`, `EXACT_MAP`, `PATTERN_FLAG`, `UNKNOWN`. No son categorías clínicas.

## Reglas candidatas

Una fila por condición observada en las columnas de clasificación (`EXACT_MAP`/`PATTERN_FLAG`). Detalle completo en `artifacts/current_logic/derived_rule_candidates.csv`.

| col | header | nº reglas | operadores |
| --- | --- | ---: | --- |
| AG | Consultas Médicas | 3 | equals |
| AH | Controles Odontologico de especialidad | 4 | equals |
| AI | Evaluaciones Odontologicas de especialidad | 6 | equals |
| AJ | Controles Ortodoncia | 6 | contains |
| AK | Controles Ortopedia preq | 3 | contains |
| AL | Instalaciones Ortopedia | 5 | contains |

## Columnas raw que alimentan cada output (transitivo)

| output | raw | header | rutas |
| --- | --- | --- | --- |
| REMASEP 01 | D | DIA CITA | AF->D |
| REMASEP 01 | G | FECHA NACIMIENTO | AF->G |
| REMASEP 01 | H | SEXO | AD->H\|AG->H |
| REMASEP 01 | O | TIPO DE CITA | AD->O\|AG->O\|O |
| B2 ANEXO | K | SUCURSAL | AC->K |
| B2 ANEXO | O | TIPO DE CITA | AC->O\|O |
| REMASEP_OD | D | DIA CITA | AF->D |
| REMASEP_OD | G | FECHA NACIMIENTO | AF->G |
| REMASEP_OD | H | SEXO | AD->H\|AE->H\|AH->H\|AI->H\|AJ->H\|AK->H\|AL->H |
| REMASEP_OD | M | ESPECIALIDAD | AE->M |
| REMASEP_OD | O | TIPO DE CITA | AD->O\|AH->O\|AI->O |
| REMASEP_OD | AA | PRESTACIÓN | AE->AA\|AJ->AA\|AK->AA\|AL->AA |

### Conjunto total de columnas raw utilizadas transitivamente

`D` (DIA CITA), `G` (FECHA NACIMIENTO), `H` (SEXO), `K` (SUCURSAL), `M` (ESPECIALIDAD), `O` (TIPO DE CITA), `AA` (PRESTACIÓN)

## Invariantes / advertencias de análisis

Sin advertencias: cada columna derivada tiene una única fórmula
normalizada y un `formula_count` igual a las 2006 filas físicas de la hoja (no atenciones reales; ver nota inicial), sin huecos.

## Observaciones que requieren validación funcional

Hechos observados en headers y dependencias. **No** se afirma que sean errores; requieren confirmación del responsable funcional.

1. **ESTADO** — `P` («ESTADO») no participa en las dependencias.
2. **MODALIDAD** — `J` («MODALIDAD») no participa en las dependencias. La primera columna `CONCAT` (`AC`, «TIPO CITA / PRESENCIAL O TELEMEDICINA») se construye a partir de «SUCURSAL», «TIPO DE CITA».
3. **PRESTACIÓN REALIZADA** — `AB` («PRESTACIÓN REALIZADA») no participa en las dependencias. Las reglas `PATTERN_FLAG` usan «PRESTACIÓN».
4. **Columnas raw utilizadas transitivamente:** `D` (DIA CITA), `G` (FECHA NACIMIENTO), `H` (SEXO), `K` (SUCURSAL), `M` (ESPECIALIDAD), `O` (TIPO DE CITA), `AA` (PRESTACIÓN).

