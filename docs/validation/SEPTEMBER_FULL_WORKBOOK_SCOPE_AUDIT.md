# Septiembre 2026: auditoría de alcance del libro completo

**Estado: implementación bloqueada por validación funcional de fuentes.** Este documento no cambia reglas ni habilita un nuevo EXE. Comparación de sólo lectura entre `REMASEP_2026_09_DRAFT(2).xlsm` (portada completada por usuario) y `2026-9 REMASEP_V1.4 (1).xlsm`, contrastada con el export directo `detalle_citas - 2026-11-03T162809.914 (1).xlsx`. No se copiaron nombres ni identificadores de pacientes.

SHA256: borrador `d5797b9f6edd83d300b57c18b238c059c51d6a6f56b6ed418ac3856eeec455e4`; histórico `a817ec514e522cb71474cbb1bf2259fff6317bc519b18aa34fd42ade751351b2`; Medinet `18b11ce185f7231c74d0d3345095b88196776ffa03efe8d47096692a696066e9`.

## Comparación por todas las hojas

Se compararon los valores guardados de todas las celdas de las once hojas. `Fórmula` cuenta valores cacheados distintos y puede reflejar diferencias de entradas aguas arriba. `Otros números` incluye valores literales positivos y ceros históricos frente a vacíos del borrador; no se debe interpretar cada cero como un dato faltante.

| Hoja | Destinos del manifiesto distintos | Otros números distintos | Fórmulas con caché distinta | Otros textos distintos |
|---|---:|---:|---:|---:|
| NOMBRE | 0 | 0 | 2 | 3 |
| REMASEP 01 | 38 | 102 | 81 | 0 |
| URGENCIAS | 0 | 0 | 1 | 0 |
| REMASEP B1 | 0 | 0 | 4 | 0 |
| B2 ANEXO | 0 | 19 | 5 | 0 |
| REMASEP_OD | 141 | 70 | 130 | 0 |
| EyP_ET | 0 | 0 | 1 | 0 |
| TV_MI | 0 | 0 | 1 | 0 |
| SERV_SANGRE | 0 | 0 | 1 | 0 |
| CONTROL | 0 | 0 | 3 | 0 |
| MACROS | 0 | 0 | 0 | 0 |
| **Total** | **179** | **191** | **229** | **3** |

Entre los 179 destinos diferentes hay 145 diferencias numéricas reales y 34 ceros frente a blancos equivalentes; los otros 943 destinos del manifiesto coinciden exactamente. El borrador modificado aún verifica 1.122/1.122 valores de septiembre calculados del export, conserva 11.466 fórmulas y su fingerprint estructural. `CONTROL!E16` es 38 en el borrador modificado y 0 en el histórico.

## Entradas manuales positivas fuera del manifiesto

Se encontraron **46 celdas positivas** en el histórico que son literales y no destinos del `write_manifest.csv`: 36 en `REMASEP 01`, 2 en `REMASEP_OD`, 8 en `B2 ANEXO`. Los demás 145 cambios de literales numéricos corresponden a ceros históricos frente a blancos del borrador. Esos ceros incluyen campos complementarios requeridos por CONTROL; exigen revisión específica y no un llenado indiscriminado.

| Hoja / sección | Filas con entradas positivas históricas | Observación |
|---|---|---|
| REMASEP 01 / A | 15, 16 | Control de Salud, Médico y Enfermera/o; 16 entradas positivas de edad/sexo. Sus totales y subtotales posteriores son fórmulas, no escrituras pendientes. |
| REMASEP 01 / B | 55, 75, 89 | Genética clínica, Cirugía Plástica y Reparadora, Otorrinolaringología; 20 entradas positivas de especialidad, edad y sexo. |
| REMASEP_OD / B | 61, 81 | Dos entradas positivas de cirugía ambulatoria, origen aún no confirmado desde el directo. |
| B2 ANEXO / D | 1251, 1255, 1256, 1258, 1267, 1303, 1304, 1307 | Ocho entradas positivas de procedimientos y rehabilitación, con grados distintos de respaldo en el directo. |

Las diferencias de fórmulas en `REMASEP 01` incluyen las filas 18, 37, 107, 122 y 307; las de `REMASEP_OD`, las filas 16, 29, 111 y 137; y las de `B2 ANEXO`, las filas 1247, 1302, 1305 y 2867. Son propagaciones de las entradas y conteos, no prueba de fórmulas alteradas. Las hojas restantes sólo muestran diferencias de cabecera o CONTROL en esta comparación.

## Verificación independiente de la Sección A

En Medinet septiembre, tras el filtro productivo de período/agenda/ESTADO, 1.226 registros elegibles. Las consultas médicas de especialidad (89 plástica, 16 otorrinolaringología y 3 genética) suman **108**. Las 37 atenciones telefónicas y 21 atenciones de enfermería suman **58**. Los sexos corresponden exactamente al histórico: 52/56 y 32/26, respectivamente. Esto respalda *esos totales* y la distribución por sexo, pero todavía requiere aprobar formalmente la definición de Control de Salud.

Las edades calculadas con `legacy_derived_record` sobre el directo **no reproducen** los tramos manuales: menores de 10 años, Médico **55** frente a **81** históricos; Enfermera/o **22** frente a **47**. Para el directo, los ocho tramos de Médico son `55, 17, 10, 8, 12, 3, 3, 0`; los históricos son `81, 7, 5, 5, 8, 1, 1, 0`. Para Enfermería, directo `22, 17, 6, 7, 6, 0, 0, 0`; histórico `47, 4, 1, 5, 1, 0, 0, 0`. En la investigación previa, 501 edades del generador legacy estaban cacheadas como cero pese a ser positivas al recalcular; esto es una hipótesis causal fuerte para los tramos de estas filas, pero la procedencia manual exacta debe verificarse antes de afirmarla como causa exclusiva.

## Otras secciones y límites de evidencia

En `REMASEP 01`, los 165 targets actuales están **exclusivamente** en las filas 113, 115, 116, 118 y 119. Por eso el resultado 1.122/1.122 no acredita las filas 15, 16, 55, 75 y 89. Los totales/sexo de las consultas especializadas también muestran coincidencias con categorías del export, pero no se validó cada tramo histórico como correcto.

En `B2 ANEXO`, varias entradas históricas son reproducibles como conteos de tipo de cita del directo: `C1251=14`, `C1256=6`, `C1258=15`, `C1267=1`, `C1303=4`. Los casos de rehabilitación `C1304=135` y `C1307=55` **no** se justifican por una correspondencia simple con las categorías de cita observadas; hay 101 registros de rehabilitación integral codificada y 35 de telemedicina con gratuidad después del filtro, además de otras categorías de rehabilitación. Se necesita definición de agrupación y fuente antes de escribir estos destinos. `C1255=1` tiene un registro con código compatible en PRESTACION, pendiente de validar su denominación. Las dos entradas quirúrgicas de `REMASEP_OD` requieren confirmar si el export de citas basta o se usa un registro quirúrgico externo.

`CONTROL=38` refleja cuatro validaciones menos al completar portada, pero deja 38 campos complementarios señalados; no detecta que la Sección A conserve ceros por ausencia de entradas. Un `CONTROL=0` tampoco demostraría por sí solo equivalencia del libro completo.

## Decisión técnica

**No implementar la copia de los valores manuales ni recompilar todavía.** La referencia no está validada para edad y algunas prestaciones; forzar esos números introduciría una regla no respaldada. Para ampliar alcance hay que acordar por cada fila la fuente, el estado de cita, el tipo de agenda, el criterio clínico, la edad a fecha de atención y el tratamiento de registros de rehabilitación y cirugía. Una vez confirmado, añadir nuevos destinos al manifiesto, actualizar el contrato de completitud, pruebas de regresión por mes y validaciones del libro completo; comparar el resultado con evidencia independiente, no sólo con el histórico.

Preguntas para Fundación Gantz: (1) ¿Qué export o registro produjo las filas 15/16/55/75/89, las dos celdas quirúrgicas y los ocho procedimientos de B2? (2) ¿La edad se calcula a la fecha de cita, y se debe recalcular desde fecha de nacimiento? (3) ¿Qué estados y agendas se incluyen en cada sección? (4) ¿Cómo se separan rehabilitación presencial, gratuidad y telemedicina en `C1304/C1307`? (5) ¿Quién valida los 38 campos complementarios? No adjuntar filas identificables al responder.
