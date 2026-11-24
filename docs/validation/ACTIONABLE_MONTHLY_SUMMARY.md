# Resumen orientado a preparar el borrador

La pantalla anterior mostraba distribuciones de citas, pero no explicaba el
alcance del archivo generado ni los datos que la persona usuaria debía completar.
Además, «estados confirmados» podía confundirse con el estado de cita Confirmado,
que está excluido por la regla vigente.

La nueva pantalla muestra primero:

- Citas incluidas, celdas previstas, citas revisadas y exclusiones por estado.
- Citas con fecha en el mes y exclusiones por agenda, que antes no eran visibles.
- Hojas que se rellenan, descripción de su contenido y destinos del manifiesto.
- Portada, campos complementarios, consultas no presenciales y clasificación de
  Control de Salud pendientes de revisión.
- Estado «BORRADOR PARCIAL», comprobación de CONTROL posterior a la generación
  y botón «Generar borrador REMASEP».

Los cuatro gráficos siguen disponibles mediante «Ver detalle de las citas».
El PDF se identifica como estadísticas y distingue citas de celdas previstas.
Si hay registros inválidos se bloquea la generación y desaparece el número de
celdas previstas; la pantalla pide corregir el input.

Los conteos adicionales se obtienen de las mismas máscaras del selector
productivo y del manifiesto cargado. No se cambiaron reglas, estados incluidos,
plantilla, manifiesto ni fórmulas REMASEP. Los pendientes no se rellenan con cero.
El resumen no inventa un recuento de alertas CONTROL antes de ejecutar Excel.

Verificación Linux: 984 pruebas aprobadas, 12 omitidas, 4 pruebas Excel excluidas.
Pruebas específicas de resumen/UI/PDF/diagnóstico: 59 aprobadas, 2 omitidas.
Se revisó la pantalla a 1040×720 y 880×600; acciones accesibles y desplazamiento
vertical en la ventana mínima. Octubre reconcilia 1.809 citas, 27 fuera por agenda,
1.782 ambulatorias, 566 fuera por estado y 1.216 incluidas; 1.207 celdas previstas
en 3 hojas. Son agregados, no datos personales.

Pendiente para Windows: ejecutar la suite, generar y reabrir un borrador con
Excel COM, verificar 1.207 destinos e integridad y recompilar ONEDIR. Esta rama
no contiene un EXE nuevo. La integración solicitada a develop debe conservar
el estado de borrador y la revisión funcional pendiente de Sección A.
