# Observaciones de Fundación Gantz — octubre 2026

Fuente funcional: observaciones transmitidas por el usuario el 26 de noviembre
de 2026. Revisión de sólo lectura del export disponible, sin datos personales.
El REMASEP ya enviado por la fundación se conserva; no se autoriza un reenvío
automático ni se modifican sus originales.

Base de código: `develop` en `456f9fa`. Input revisado:
`detalle_citas - 2026-11-03T162809.914 (1)(1).xlsx`, SHA256
`38c54f592e51ba11d74885bd1831912f8f967e237d38767dacba793c2a841320`.
Octubre: 1.809 citas con fecha en el mes, 27 fuera por agenda, 566 fuera por
ESTADO y 1.216 incluidas. La copia final enviada a Gantz no se recibió en esta
iteración; las observaciones no deben atribuirse sin más al borrador anterior.

## Comprobación por observación

| Observación | Evidencia del export / plantilla | Decisión |
|---|---|---|
| REMASEP 01 W/X, filas 55, 75 y 89, deben ser cero | Las consultas codificadas de genética (7), plástica (119) y otorrino (8) tienen SUCURSAL=Santiago. No hay telemedicina entre estas consultas. | Cero de octubre respaldado por Gantz. La app sigue sin escribir estos campos; no se impone cero a otros meses ni se inventa el reparto teléfono/otras plataformas. Falta cerrar su regla general. |
| REMASEP 01 C89 incompleto | 8 consultas tipo COD.0101205 y 7 controles postoperatorios de especialidad otorrino. El usuario confirmó incluir ambas categorías el 26 de noviembre. | **Implementado:** fila 89 incluye la consulta codificada o TIPO_DE_CITA=CONTROL POST OPERATORIO con ESPECIALIDAD=Otorrinolaringología. Octubre pasa de 8 a 15 (9 hombres y 6 mujeres). Se actualizan edades y sexo; C89 sigue siendo fórmula. |
| B1 otorrino son exámenes, no consultas | `REMASEP B1!C102 = 'B2 ANEXO'!C1247`; resultado guardado 37. | La fórmula ya usa procedimientos diagnósticos, no el total de consultas médicas. Sin cambios. |
| B1 kinesiología sí corresponde | `REMASEP B1!C112 = 'B2 ANEXO'!C830`; resultado 32. | Sin cambios. |
| URGENCIAS y B2 ANEXO aprobados | Confirmación funcional de Gantz. | Se mantienen reglas y fórmulas. |
| OD 58, instrucción de higiene oral, código 5010012 | La fila oficial se llama Educación en salud bucal y no estaba en el manifiesto. El código aparece en agosto/septiembre, pero no en octubre del export disponible, antes o después del filtro. | **Implementado:** 34 destinos por edad/sexo basados en el código 5010012 de PRESTACION. Una cita cuenta una vez aunque el código se repita. En este input: septiembre=1, octubre=0. Si Gantz esperaba un positivo en octubre hay que conciliar el export efectivamente usado. |
| OD 59, terceros molares | La fila no estaba en el manifiesto. En octubre hay 4 citas incluidas con 5003001 EXODONCIA y 1 con 5003002 EXODONCIA COLGAJO. Ninguna identifica tercer molar. No hay columna estructurada de pieza/diente en el export. | No se convierten todas las exodoncias en terceros molares. Se requiere código o pieza y definición de unidad de conteo. |
| OD 65, aparato removible | Las 3 citas con 5003017 INSTALACIÓN CONTENCION REMOVIBLE ya cuentan en fila 67, Instalación aparato de contención ortodóncico. La fila 65 es Instalación aparato removible ortodóncico. Una cita repite el código, sin evidencia de si representa 2 dispositivos o duplicación. | No se trasladan ni duplican esas 3 citas. Confirmar código de aparato, destino clínico y unidad de conteo. |
| OD 108, rehabilitación de prótesis | El usuario distingue la categoría de consulta maxilofacial prótesis de la categoría rehabilitación oral prótesis. En el export aparecen ESPECIALIDAD=Maxilofacial, Prótesis de Paladar y Rehabilitación oral, pero no un valor literal Maxilofacial prótesis. | Separación conceptual confirmada; pendiente identificar los valores exactos/códigos de cada grupo y su destino 108/109. No se cuentan todas las consultas de Maxilofacial ni todas las etapas de prótesis. |
| Nuevos datos de pueblos originarios, país y nacionalidad | Gantz los solicitó a Medinet; no están disponibles en este export. | Integrar cuando se reciba el esquema y se valide su significado. Ausencia de campo no equivale a cero. |
| CONTROL aprobado por Gantz | Se refiere al archivo que enviaron. | No se modifica CONTROL ni se extrapola ese resultado a un output futuro. |

## Cambio de código verificado

El manifiesto pasa de 1.207 a **1.241 destinos** al incorporar OD58. Esa primera
corrección conservó los 1.207 valores anteriores: para septiembre su digest era
`9dbdae29b64de3c839b97a18c1c1b52eeda1ed1bde3afdf859952b4e800e8b32`.
Para octubre su digest era
`155bf59685524665fb5d68ea0d19aa4481da4c2b84d183d1e8f4664398a93c19`.
Esos digests describen el estado anterior a la confirmación de C89 y ya no
deben exigirse al contrato ampliado completo después de corregir otorrino.
Las 34 celdas nuevas están desbloqueadas y sin fórmulas en la plantilla oficial.
Los totales B58:D58 siguen siendo fórmulas de la plantilla, sin sobrescribirlas.

La regla nueva usa el namespace virtual MEDINET_DIRECT_REMASEP_OD en el
catálogo para no reemplazar coordenadas usadas por las dependencias legacy.
No abre el generador histórico en producción ni cambia el filtro ambulatorio.

La segunda corrección modifica únicamente los 19 criterios de edad/sexo de la
fila 89. Una comparación antes/después en septiembre y octubre confirmó que
todos los otros destinos mantienen su valor. Septiembre pasa de 16 a 19 y
octubre de 8 a 15. No se amplía por esta confirmación la Sección A (Control de
Salud), ni se incluyen exámenes, postoperatorios de otra especialidad o citas
de estados excluidos.

Validación Linux: **987 aprobadas, 12 omitidas, 4 pruebas Excel excluidas**.
Incluye regresión sintética para código dentro de una lista, código repetido,
filtro ESTADO y bandas de edad/sexo. Septiembre y octubre producen 1.241
destinos únicos y completeness.ok=true. No se ejecutó Excel COM ni se compiló
un EXE nuevo en este entorno. El conteo positivo de OD58 también puede exigir
completar sus campos complementarios en CONTROL; no se rellenan por defecto.

## Información necesaria para cerrar los pendientes

Confirmar si el input usado para el archivo enviado es el mismo identificado
arriba y los valores/códigos que identifican las dos categorías de prótesis.
La inclusión de los siete postoperatorios de C89 ya está confirmada e implementada.
Para terceros molares y aparatos se necesita el código/pieza/unidad de
conteo que utiliza Gantz. Estas decisiones no se deducen del match histórico.
