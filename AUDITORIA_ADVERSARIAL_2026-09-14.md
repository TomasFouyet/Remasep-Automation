
REMASEP Automation — auditoría del estado actual y segunda pasada adversarial

Auditoría de sólo lectura cerrada el 14 de septiembre de 2026. Rama audit/senior-project-review; HEAD d5f416e1ead6fef33a6950d5430eb32f2aedc690. Estado Git limpio al inicio y al cierre. No se modificaron archivos productivos, configuración, reglas ni tests del repositorio. No se hicieron commits. Las reproducciones y este informe están fuera del repositorio, en /tmp.

La revisión combina inspección del código, ejecución de la suite, reproducciones sintéticas y comprobaciones estáticas de los workbooks locales. No se ejecutó Microsoft Excel en esta auditoría ni se reconstruyó el ejecutable Windows. Los resultados con un writer simulado se identifican expresamente: prueban decisiones del orquestador, no comportamiento de Excel.

1. EXECUTIVE SUMMARY

Hay cuatro bugs P1 reproducidos, dos riesgos P1 sustentados y tres bugs P2. No confirmo un P0; tampoco interpreto su ausencia como garantía de seguridad o exactitud del archivo final.

El problema funcional principal es la pérdida de diagnósticos de filas inválidas en el flujo actual de UI: una atención desaparece, el resumen muestra 1 / 1 y el productor aprueba las 1.122 escrituras.

La compatibilidad de plantilla no comprueba las expresiones oficiales. Sustituir una fórmula de total o de CONTROL por =0 mantiene el fingerprint esperado.

Un paquete aislado con los assets declarados por el spec pasa SMOKE OK y falla al calcular por ausencia de las reglas legacy productivas.

La navegación admite dos generaciones simultáneas. Cancelar un análisis no impide que su resultado atrasado reemplace el estado actual. Cerrar durante el trabajo aborta el proceso Qt.

La promoción usa una comprobación de existencia seguida de replace: dos ejecuciones pueden declarar éxito sobre un único archivo, sobrescribiendo una a la otra.

Abrir un .xlsm con DispatchEx no demuestra que sus macros de apertura estén desactivadas. Falta una política explícita en esa instancia; requiere validación Windows.

El resultado visual omite CONTROL y advertencias específicas. Se puede mostrar éxito con CONTROL no disponible. El aviso genérico de fuentes pendientes sí existe.

El exportador PDF puede retornar normalmente sin crear ningún archivo.

La suite terminó con 923 passed, 3 skipped, 1 warning. Ruff, git diff --check, compilación e imports revisados pasaron. Estos resultados conviven con las reproducciones anteriores.

La evidencia histórica no demuestra una exactitud general de 87,70 %. Tampoco demuestra causalidad absoluta de todos los mismatches por caché AF: el subconjunto no-AF tiene sólo 20 comparaciones y se reutiliza un export actual.

Conviene estabilizar este flujo antes de añadir otras fuentes. El diseño permite hacerlo sin reescribir el motor.

2. TOP RISKS

Prioridad

ID

Severidad

Tipo

Evidencia que sobrevivió a la refutación

1

F01

P1

BUG

Dos registros → uno visible; 1.122 escrituras completas; sin diagnóstico de la fila perdida en el resumen productivo.

2

F02

P1

RISK

Fórmulas oficiales alteradas conservan el fingerprint; la comparación posterior usa la plantilla elegida como referencia.

3

F03

P1

BUG

Bundle aislado: smoke retorna 0; cálculo lanza LegacyRulesError.

4

F04

P1

BUG

Qt real: aborto al cerrar, dos workers simultáneos y resultado de julio aplicado tras seleccionar junio.

5

F05

P1

BUG

Dos ejecuciones sincronizadas pasan ambas guardas y reemplazan el mismo archivo; ambas retornan éxito.

6

F06

P1

RISK

Apertura COM sin configuración de seguridad de macros; contrato documentado por Microsoft. NEEDS_WINDOWS_VALIDATION.

No atribuyo frecuencia de producción sin telemetría o muestras. En los hallazgos siguientes, “determinista” significa reproducible al darse el escenario, no frecuente en el uso normal. Los bugs de validación y presentación se mantienen en P1/P2; no se los transforma en P0 afirmando una corrupción clínica final que no se ejecutó en Excel.

3. FULL FINDINGS

F01 — P1 · Corrección / pipeline / UI · BUG CONFIRMADO. El flujo productivo pierde los diagnósticos de registros inválidos y permite preparar un informe parcial sin advertirlo.

Evidencia: _scope_selection() identifica registros inválidos; processing_scope_frame() retorna únicamente los válidos del período y descarta las máscaras y problemas. El resumen actual usa directamente ese frame. run_generation() comprueba errores del productor, pero no recibe el diagnóstico de registros descartados. Referencias: medinet_analysis.py:189, medinet_analysis.py:247, medinet_summary.py:170, workers.py:101.

Reproducción: crear dos atenciones kinesiólogicas sintéticas del mismo mes, ambas Atendido. Con ambas válidas, B2 ANEXO!C833=2 y el resumen dice 2 / 2. Cambiar solamente el nacimiento de una fila a un texto inválido produce: diagnóstico antiguo invalid_records=1, resumen actual 1 / 1, excluidos 0, C833=1, 1.122 PendingWrites, completitud verdadera, cero conflictos de tipo y cero fórmulas no soportadas. Se reprodujo también con sexo vacío. Las ejecuciones repetidas dieron los mismos valores.

Refutación: excluir una fila inválida es una decisión explícita del contrato. El defecto confirmado es que la UI actual pierde el diagnóstico y permite continuar sin resolverlo. No recomiendo inventar edad o sexo ni contar automáticamente datos inválidos. Una fila válida adicional basta para superar la detección de período de la UI. Las 1.122 instrucciones no prueban completitud de los registros de origen.

Escenario e impacto: export parcialmente corrupto o incompleto; subreporte presentado sin el dato que permitiría detenerlo. Probabilidad no medida; comportamiento determinista con mezcla de filas válidas e inválidas. El guardado final por COM no se ejecutó en este repro.

Por qué pasa los tests: los problemas de filas se prueban en MedinetAnalysisService; las pruebas del resumen/productor verifican el subconjunto ya filtrado. Falta una prueba del contrato de validación consumido por la UI actual.

Fix recomendado: devolver un resultado de ingesta/validación compartido con cantidades, filas/campos afectados y bloqueos; conservarlo hasta el resultado. Exigir resolución de inválidos relevantes antes de generar. Definir explícitamente el tratamiento de fechas de cita imposibles de asignar a un período.

Test necesario: dos filas, una inválida; la UI debe informar el problema y el servicio de generación no debe invocar el writer hasta resolverlo. Riesgo de regresión medio: un bloqueo indiscriminado sobre todo el export podría impedir meses sanos por problemas de otros períodos.

F02 — P1 · Integridad de plantilla · RISK con brecha reproducida. Una plantilla con fórmulas incorrectas puede pasar la compatibilidad y convertirse en la referencia de integridad.

Evidencia: structural_template_fingerprint() considera coordenadas de fórmulas, protección, desbloqueos, merges y hojas; no incluye el contenido de las fórmulas. generate() toma el snapshot inicial de la plantilla seleccionada y compara el resultado contra ella. Referencias: writable_target_mapping.py:211, excel_writer.py:804, excel_writer.py:872.

Reproducción: abrir la plantilla oficial local sólo en memoria, sustituir REMASEP 01!C15 por =0, y recalcular el fingerprint. Sigue siendo stf:dc624775927d4d4d. Sucede también modificando CONTROL!E16 a =0. Cambios de defined names y visibilidad tampoco alteran esa huella. Con un writer simulado, una fórmula ya alterada en la referencia inicial llega a GENERATED_DRAFT e integridad aprobada. No se guardó la plantilla local modificada.

Refutación: el chequeo posterior detecta muchas modificaciones introducidas durante la escritura, pero no una expresión errónea que ya estaba en el archivo seleccionado. El SHA del archivo es informativo y no se compara con un patrón oficial aprobado. La huella conserva utilidad estructural; el riesgo está en interpretarla como validación del contenido oficial.

Escenario e impacto: elegir una copia retocada o una versión de plantilla con fórmulas distintas y estructura equivalente. Puede conservar totales o validaciones incorrectas. Probabilidad desconocida, dependiente del control de distribución de plantillas. La aceptación está reproducida; la generación Excel con esa plantilla queda NEEDS_WINDOWS_VALIDATION.

Por qué pasa los tests: comprueban cambios estructurales y diferencias antes/después sobre una misma referencia. No incluyen una fórmula incorrecta anterior a la generación con idénticas coordenadas.

Fix recomendado: validar expresiones y elementos semánticos críticos contra un contrato aprobado por versión; definir qué constantes/metadatos pueden variar. Evitar depender exclusivamente del SHA binario completo, porque un re-guardado legítimo puede cambiarlo. Test: cambiar sólo una expresión de total/CONTROL y exigir rechazo, conservando aceptación de re-guardados legítimos. Regresión media/alta: definir mal el contrato puede rechazar plantillas válidas.

El workbook histórico de julio también conserva ese fingerprint y NOMBRE!B5=Julio. El writer no actualiza esa celda al elegir otro mes. Lo registro como riesgo de usar un archivo previo y necesidad de revisión del borrador, no como un bug independiente que obligue a automatizar metadatos fuera del alcance acordado.

F03 — P1 · Packaging / runtime assets · BUG CONFIRMADO. El paquete declarado omite reglas necesarias para generar y el smoke no detecta esa ausencia.

Evidencia: common_paths() incluye sólo runtime_2026 y excel_writer_2026; producción llama a load_legacy_rules(), cuyo archivo está fuera de esos directorios. Su resolver usa parents[3], sin el resolver común del bundle. Referencias: packaging/_spec_common.py:18, legacy_rules.py:23, production_pipeline.py:122, main.py:23.

Reproducción: copiar el paquete Python y exactamente los assets declarados a un directorio temporal con layout _internal, fijar _MEIPASS, cargar los módulos desde esa copia y cambiar el cwd fuera del repo. _smoke_check() retorna 0 y escribe SMOKE OK; calcular una fila sintética lanza LegacyRulesError buscando config/legacy_current_logic_2026/rules.yaml fuera de _internal.

Refutación: los specs no agregan un hook que incluya este YAML. Cambiar solamente el cwd no expone el defecto porque el código del checkout sigue encontrando sus reglas. Añadir únicamente el archivo a _internal/config tampoco corrige el resolver actual. El análisis/resumen nuevo puede funcionar: el fallo llega al cálculo para generar.

Escenario e impacto: instalación aislada construida con estos specs; bloqueo del objetivo principal de la app después de un smoke exitoso. Determinista para el layout probado. No inspeccioné un dist entregado previamente porque no existe uno en este checkout; NEEDS_WINDOWS_VALIDATION para confirmar el ejecutable final, sin retirar el bug de dependencia demostrado.

Por qué pasa los tests: las pruebas de paths/packaging reutilizan los módulos del checkout o sólo verifican los assets actualmente enumerados. El smoke nunca calcula una atención.

Fix recomendado: incorporar las reglas productivas al contrato de assets, resolverlas con el mecanismo del bundle y verificar su integridad. Test: distribución aislada, sin acceso al checkout, que genere 1.122 PendingWrites de una fila conocida. Regresión baja/media: al mover las reglas, deben conservarse exactamente sus literales y semántica.

F04 — P1 · Concurrencia / UI / confiabilidad · BUG CONFIRMADO. El ciclo de vida de los workers permite resultados obsoletos, trabajos simultáneos y aborto al cerrar.

Evidencia: _start() crea otro hilo sin guarda de operación activa; _teardown() ignora si wait() terminó; las señales no se identifican por corrida y _on_done() modifica el estado compartido. MainWindow no implementa una política de cierre durante trabajo. Referencias: generate.py:103, generate.py:181, analysis.py:94, analysis.py:111, main_window.py:62.

Reproducciones con Qt real y backend lento controlado: (a) iniciar generación y cerrar: salida -6, QThread: Destroyed while thread ... is still running; (b) volver al resumen y pulsar Generar mientras sigue el primero: dos QThreads activos para el mismo output; (c) cancelar análisis de julio, cambiar a junio y liberar el worker antiguo: la app vuelve sola al dashboard con estado junio y resumen julio.

Refutación: quit() no interrumpe la función que está ejecutándose. El tiempo de espera no es una cancelación confirmada. El backend lento simulado controla el momento de la carrera; el manejo de QThread y señales es el de producción. La prueba de cierre se aisló en un subproceso sin core dump. No se invocó Excel.

Impacto: revisión de datos de un período distinta del estado seleccionado, trabajos duplicados y terminación abrupta. Escenario cotidiano: usuario cancela, vuelve atrás o cierra mientras espera. Probabilidad no medida; reproducible bajo esos gestos. Excel/workspaces huérfanos son un riesgo adicional NEEDS_WINDOWS_VALIDATION, no un resultado observado aquí.

Por qué pasa los tests: los helpers de UI drenan el hilo y ejecutan run_now(); no verifican cancelación/cierre con trabajo en curso (test_ui_shell.py:77).

Fix: dueño único de la operación activa, identidad de corrida y parámetros inmutables, descarte de resultados cancelados/obsoletos y cierre diferido mientras se libera el worker. Test: los tres escenarios anteriores usando eventos de sincronización. Regresión media: evitar bloquear indefinidamente el hilo de UI o interrumpir Excel durante Save.

F05 — P1 · Writer / concurrencia · BUG CONFIRMADO. La promoción atómica permite sobrescritura entre generaciones concurrentes.

Evidencia: generate() comprueba output_path.exists() y luego _atomic_promote() ejecuta Path.replace(). No hay exclusión mutua ni publicación que falle atómicamente si existe destino. Referencias: excel_writer.py:690, excel_writer.py:919.

Reproducción: dos llamadas al orquestador con writers simulados y archivos temporales distintos; barrera inmediatamente después de las comprobaciones de existencia, antes de invocar la promoción real. Ambas retornan GENERATED_DRAFT, sin errores; existe un solo output y uno de los contenidos fue reemplazado. El sistema de archivos y replace() son reales.

Refutación: la segunda comprobación de existencia reduce la ventana, pero sigue separada de la mutación. El rename atómico protege la publicación de un archivo completo; no promete rechazo del destino existente. Python documenta que os.replace reemplaza un archivo destino cuando los permisos lo permiten. Documentación de Python.

Escenario e impacto: dos instancias o el solapamiento de F04 generan sobre el mismo nombre. El usuario puede abrir una salida distinta a la que su corrida reportó. Baja probabilidad con un solo proceso secuencial; la carrera es real bajo concurrencia. Semántica de permisos/locks/OneDrive en Windows: NEEDS_WINDOWS_VALIDATION.

Por qué pasa los tests: comprueban salida preexistente y promoción aislada, sin sincronizar dos publicaciones. Fix: exclusión interproceso por destino y/o primitiva de publicación sin reemplazo, con tratamiento explícito de reservas y fallos. Test: dos procesos; exactamente uno publica y el otro recibe OUTPUT_ALREADY_EXISTS, sin alterar el primero. Regresión media/alta por las diferencias entre NTFS, volúmenes y recursos compartidos.

F06 — P1 · Seguridad / Excel COM · RISK. La apertura de la plantilla no establece una política explícita que impida macros automáticas. NEEDS_WINDOWS_VALIDATION.

Evidencia: ExcelSession.__enter__() establece Visible y DisplayAlerts; open_workbook() llama Workbooks.Open(..., ReadOnly=...). No configura AutomationSecurity ni EnableEvents. Referencia: excel_com.py:55.

Microsoft documenta que AutomationSecurity inicia en msoAutomationSecurityLow, y que DisplayAlerts=False no controla advertencias de seguridad. La ausencia de Application.Run no basta para demostrar que un evento de apertura no se ejecute. Microsoft: Application.AutomationSecurity.

Refutación: políticas corporativas, procedencia del archivo y configuración de Office pueden bloquear la ejecución. No afirmo que cualquier plantilla ejecute macros en cualquier equipo, ni que la plantilla oficial observada contenga un evento dañino. No hay explotación o fuga reproducida.

Escenario e impacto: plantilla con un evento de apertura que modifica datos o produce efectos externos antes de verificarse el resultado. Probabilidad dependiente de plantilla y políticas; impacto potencial alto.

Cómo validarlo: en una VM Windows de prueba y sin cambiar Trust Center, abrir mediante el adapter un .xlsm sintético con Workbook_Open que sólo escriba un marcador en una celda. Verificar si aparece; registrar versión Office y política efectiva. No usar datos clínicos.

Tests actuales: simulan COM y comprueban sesión/cierre, sin un workbook con eventos. Fix: desactivar macros para la apertura programática en la instancia propia y definir tratamiento de eventos y links; preservar la configuración previa. Excel 4.0 macros requieren tratamiento adicional según la misma documentación. Regresión media: comprobar que el recálculo autorizado de la plantilla no depende de macros o UDF bloqueadas.

F07 — P2 · Integridad del writer · BUG CONFIRMADO. Las comparaciones aceptan equivalencias que no garantizan preservar fórmulas o escrituras explícitas de cero.

Evidencia: _norm_formula() elimina todos los espacios y aplica casefold() también dentro de literales; verify_targets() sustituye un destino ausente por cero si esperaba cero. Referencias: excel_writer.py:355, excel_writer.py:425.

Reproducción: comparar =IF(A1="A B",1,0) con =IF(A1="AB",1,0) produce integridad aprobada. También =EXACT("a","a") frente a =EXACT("a","A"). Retirar del snapshot post-save una celda cuyo PendingWrite era cero produce OK, written_ok=1 y generación exitosa en el harness.

Refutación y severidad: son falsos negativos confirmados de los verificadores. No encontré un camino normal de COM que provoque esas mutaciones espontáneamente. Por eso bajo a P2 y no afirmo corrupción frecuente. La equivalencia histórica cero/blanco no demuestra que se haya cumplido la política productiva WRITE_ZERO.

Impacto: se debilita la detección de modificaciones o escrituras ausentes. Probabilidad en generación real desconocida. Tests: cubren fórmulas cambiadas y valores, pero faltan cambios dentro de strings y cero ausente. Fix: comparar fórmulas preservando tokens y literales, y exigir evidencia del cero en la celda guardada. Tests: los contraejemplos anteriores y casos de normalización legítima. Riesgo de regresión medio por variantes de representación de fórmulas al guardar con Excel.

F08 — P2 · UI / observabilidad · BUG CONFIRMADO. El resultado de generación oculta CONTROL no disponible y advertencias específicas.

Evidencia: el orquestador permite un borrador con CONTROL fallido o no disponible. GenerationOutcome lleva control_status y warnings, pero _show_result() sólo presenta nombre, período, cantidad e integridad. Referencias: excel_writer.py:942, generate.py:230.

Reproducción: snapshot sin valor de CONTROL retorna GENERATED_DRAFT, CONTROL_UNAVAILABLE, sin warnings. Pasar un resultado así, incluso con una advertencia identificable, a la pantalla Qt muestra el título “REMASEP generado correctamente”; ni el estado de CONTROL ni la advertencia aparecen.

Refutación: el MVP genera un borrador y faltan fuentes externas; CONTROL != 0 no debe convertirse automáticamente en fallo del writer. La UI sí advierte genéricamente que faltan secciones. Por ello retiro la acusación de que no existe ninguna advertencia o que un borrador con CONTROL distinto de cero sea por sí mismo un bug P1.

Impacto: el usuario no distingue validación pendiente por alcance de un CONTROL imposible de comprobar, ni ve otras advertencias como cleanup pendiente. Probabilidad determinista si llega ese estado; frecuencia real desconocida. Tests: comprueban éxito y fallos generales, sin exigir que la pantalla preserve estas diferencias.

Fix: mostrar estado de borrador, CONTROL disponible/fallido/no disponible y advertencias pertinentes con texto comprensible; mantener separado el resultado del writer del estado de revisión. Test: los tres estados de CONTROL y warnings específicos. Regresión baja: cambio de presentación, preservando el alcance sólo-Medinet.

F09 — P2 · PDF / false success · BUG CONFIRMADO. Se anuncia que el PDF se exportó aunque el archivo no exista.

Evidencia: export_summary_pdf() no comprueba el éxito de inicio/fin del QPainter; retorna el path. DashboardScreen._export_pdf() anuncia éxito si no hubo excepción Python. Referencias: pdf_report.py:115, dashboard.py:200.

Reproducción: usar un directorio temporal con permisos 0500; la función retorna summary.pdf, exists=False, y Qt emite QPainter::begin(): Returned false. No se genera excepción. Se observaron 147 mensajes Qt derivados del mismo fallo.

Refutación: el chequeo de archivo preexistente funciona, pero no detecta que el directorio existente carece de permisos de escritura. Los mensajes Qt no llegan al flujo de excepciones del dashboard. Es un fallo del PDF; no lo elevo a P1 ni lo presento como corrupción del .xlsm.

Escenario e impacto: pérdida de permisos o carpeta elegida sin acceso de escritura; el usuario cree que dispone del resumen. Determinista con el permiso probado. ACL Windows: NEEDS_WINDOWS_VALIDATION.

Tests: verifican exportación normal y no-overwrite; falta fallo de dispositivo de pintura/escritura. Fix: comprobar apertura y finalización del dispositivo y existencia de un PDF válido antes de confirmar, preferiblemente publicando desde un archivo temporal. Test: destino no escribible y fallo de finalización. Riesgo de regresión bajo.

Registro de refutaciones y reducciones de severidad

Sospecha inicial

Resultado de la segunda pasada

ESTADO desconocido desaparece completamente

Refutada: se muestra como excluido y figura en distribución. Falta una decisión funcional para estados nuevos; no afirmo que deban contarse. Fuera de P0/P1.

Falta ESTADO y se genera todo en cero

Refutada: la ruta actual falla al acceder a la columna. Es mala validación de contrato, pero no el false success sospechado.

Cualquier CONTROL fallido invalida el MVP

Refutada: el borrador sólo-Medinet admite fuentes pendientes. Se conserva F08 por ocultar el estado concreto.

No existe aviso de incompletitud en UI

Refutada: existe un aviso genérico. El defecto es no mostrar la evidencia específica.

Edad Python incorrecta porque hay 552 mismatches

No sustentada. El informe histórico registra validación AF contra Excel; esta auditoría no encontró una refutación de esa implementación.

VBA debe ser idéntico byte a byte

Refutada como criterio universal: el diseño compara código semántico y admite cambios de contenedor. Preservación no equivale a confianza en el VBA de origen.

Cleanup que no borra un archivo bloqueado es un bug crítico

Refutada: evita forzar handles o tocar instancias ajenas; el pending debe ser visible.

Elegir otro mes debería completar todas las fuentes/metadatos

Fuera del alcance acordado. Registrar y revisar metadatos pendientes no obliga a automatizar egresos ni recursos.

Falta de regla para un tipo desconocido siempre prueba subconteo

No sustentada sin clasificación funcional esperada: existen atenciones fuera de alcance. Requiere catálogo/decisión de negocio.

Falsos negativos de comparación prueban corrupción espontánea de Excel

No: F07 queda P2 porque se reprodujo la debilidad de la barrera, no el desencadenante normal.

Warning de /dev/null prueba fuga de memoria productiva

No: se observó en el test que sustituye streams; no justifica P0/P1.

4. THINGS THAT ARE DONE WELL

El pipeline productivo consume el export directo y assets versionados. No abre el generador histórico para calcular. Las reglas con nombre legacy siguen siendo productivas: no deben eliminarse por su nombre.

La selección de período precede al filtro ESTADO, y éste precede a métricas. El conjunto incluido está declarado CONFIRMED en el asset actual.

Los cinco assets enumerados en bundle.yaml verifican hashes; el bundle contiene 1.122 destinos y 1.224 fórmulas. Se validan duplicados y dependencias declaradas.

Los valores de conteo rechazan booleanos, negativos y fracciones; las fórmulas fuera del subset se reportan como no soportadas.

El writer trabaja sobre una copia, comprueba destinos, recalcula y verifica después del guardado antes de promover. El diseño evita reconstruir el .xlsm con openpyxl.

Excel usa DispatchEx, inicialización COM y liberación de la instancia propia. No hay un mecanismo global de matar Excel.

Los tests cubren muchos errores del motor y del writer, y existen tests locales de regresión real y tests opt-in de Excel. Los tres omitidos son visibles en la ejecución.

El logging tiene rotación y una ruta de datos de aplicación. La UI presenta mensajes humanos para muchas excepciones conocidas.

5. FALSE CONFIDENCE RISKS

“1.122 escrituras completas” significa que están presentes los destinos, no que se hayan contabilizado correctamente todos los registros. “Determinista” tampoco demuestra corrección: el subconteo de F01 se repite exactamente.

“Fingerprint compatible” demuestra una parte de la estructura. No certifica las expresiones oficiales ni que los datos manuales pertenezcan al período seleccionado. “VBA íntegro” comprueba preservación respecto de una referencia, no que esa referencia sea confiable.

“SMOKE OK” no recorre la dependencia de reglas que necesita el cálculo. “Build OK” en el script de construcción comprueba que existe el ejecutable; no equivale a una generación fuera del checkout.

“923 tests verdes” incluye gran cantidad de lógica de ingeniería inversa y verificaciones con fakes. Las pruebas nuevas muestran fallos en las fronteras entre esos componentes. Los repros del writer simulado tampoco deben presentarse como tests Windows completos.

Sobre el histórico: el documento registra 3.932 coincidencias exactas, cuatro equivalencias cero/blanco y 552 diferencias. Sólo 20 de 4.488 comparaciones no dependen de AF. La coincidencia de ese subconjunto y la validación de edad apoyan la hipótesis de caché obsoleta, pero no excluyen cambios del dato fuente histórico. La conclusión categórica del último párrafo del documento excede su propia salvedad metodológica. No se alteró ninguna regla para mejorar porcentajes. Informe histórico del repositorio.

6. MISSING TEST MATRIX

Área / escenario

Resultado exigible

Evidencia pendiente

Fila inválida entre válidas, también fuera de período

Diagnóstico preservado y política explícita antes de generar

Integración UI → validación → productor

ESTADO nuevo/vacío, sexo/tipo no reconocido

Categoría conocida, exclusión justificada o revisión; sin interpretación inventada

Confirmación funcional y tests

Fechas texto mezcladas, seriales Excel, celdas con error/fórmula

Interpretación validada o rechazo visible

Contrato de formatos del export real

Cumpleaños, bisiestos y fecha futura

Reproducir política acordada y señalar anomalías

Ampliar bordes sin sustituir la evidencia AF existente

Plantilla con expresión/etiqueta/name alterados, mismo layout

Rechazo de cambios incompatibles

Contrato semántico de plantilla

Cero desaparecido y literal de fórmula cambiado

Fallo de integridad

Tests puros y luego Excel

Cancelar/cambiar período/cerrar con worker activo

Resultado antiguo descartado; cierre ordenado

Qt con event loop real; Windows para COM

Dos procesos, mismo output

Un solo éxito, sin sobrescritura

Multiproceso y NTFS/OneDrive/recurso compartido

CONTROL PASS/FAIL/UNAVAILABLE y cleanup pendiente

Mensajes distintos y revisión explícita

UI con estados reales del servicio

PDF no escribible o fallo al terminar

Sin mensaje de éxito y sin archivo parcial publicado

Qt y ACL Windows

Paquete fuera del repo, sin Python/Git

Cálculo real con todos los assets

Build Windows y máquina limpia

Excel ocupado, output abierto, diálogo modal, sin Excel

Fallo o espera acotada, sin tocar otras instancias

Windows/Excel

Evento Workbook_Open, links y Protected View

Política explícita, sin efectos no autorizados

VM Windows controlada

Excepciones de parser con contenido sensible sintético

Ningún valor sensible en logs/diagnósticos

Tests adversariales de privacidad

Muchas corridas y fallo/crash durante save

Recursos liberados y residuos atribuibles a una corrida

Prueba prolongada Windows

No propongo property-based testing o mutation testing por su nombre. Aportarían valor sobre reconciliación de filas, normalización de fórmulas y estado de corridas. Primero hacen falta los tests de comportamiento que reproducen los defectos encontrados.

7. ARCHITECTURE REVIEW

Mapa real: entrypoint remasep.main:main → ui.main_window.run_app → pantallas y workers. El análisis visible llama build_monthly_medinet_summary → processing_scope_frame → adapter Medinet y filtro ESTADO. La generación vuelve a leer Medinet mediante build_production_pending_writes, carga el bundle y las reglas YAML, deriva AC, evalúa las fórmulas del catálogo, produce MetricValues y las une al manifiesto para obtener PendingWrites. GenerationService resuelve política/capacidad y llama al writer: preflight → snapshot → copia temporal → COM → recálculo/Save → inspección → promoción → diagnósticos.

scripts/generate_remasep.py ofrece un camino CLI. Los scripts de inventario, mapping, comparación histórica y regeneración de assets son herramientas de desarrollo. MedinetAnalysisService sigue existiendo y se prueba, pero no es el resultado de validación que muestra el dashboard actual. El motor de reglas estricto de core/rules.py tampoco acredita que la ruta legacy productiva bloquee todas las clasificaciones desconocidas.

Estructura inspeccionada: src/remasep/{adapters,core,domain,services,ui,testing,config}, scripts, tests, config, packaging, docs y sus sprints/validaciones; directorios locales data y artifacts excluidos de Git. No se encontraron instrucciones AGENTS.md en las ubicaciones de ascendencia revisadas. README y documentos iniciales afirman que el writer/mapping aún no existen: esas afirmaciones ya no describen el código.

Pregunta

Evaluación

1. ¿Arquitectura razonable?

Sí para una app local que preserva un workbook oficial; separar cálculo, manifiesto y escritura es apropiado.

2. ¿Acoplamientos excesivos?

UI y estado mutable de corridas; ingestión y selección que devuelve sólo el frame; reglas productivas fuera del bundle; configuración específica de 2026 dispersa.

3. ¿Qué no tocaría?

Semántica legacy validada, separación MetricValue/PendingWrite, writer sobre copia y frontera COM inyectable.

4. ¿Qué refactorizar antes de nuevas fuentes?

Un resultado compartido de validación/alcance; dueño de la operación UI; contrato completo de assets y plantilla por versión.

5. ¿Sobreingeniería?

Hay abundante infraestructura histórica de análisis. No demuestra un problema productivo por sí misma; su distinción respecto del runtime necesita claridad.

6. ¿Deuda peligrosa?

Sí: validaciones que existen pero no llegan al usuario, seguridad de apertura implícita, y éxito que mezcla significados.

7. ¿Demasiadas capas?

Las fronteras principales se justifican. No añadiría un framework genérico de plugins/fuentes ni más capas antes de reparar los contratos actuales.

8. ¿Negocio escondido en infraestructura?

Parte vive en YAML y fórmulas/columnas legacy congeladas. La ruta productiva todavía depende de coordenadas del generador como representación del cálculo; la afirmación documental de independencia total es excesiva.

9. ¿Soporta 2027?

Reutilizable, pero no basta cambiar el año: necesita bundle, mapping, CONTROL y contrato de plantilla nuevos, seleccionados explícitamente. No se ha demostrado compatibilidad 2027.

10. ¿Agregar egresos/tabla ahora?

Estabilizar primero. Añadir fuentes multiplicaría los problemas de alcance, validación y estado de ejecución.

Las reglas de legacy_current_logic_2026/rules.yaml siguen declaradas pending_functional_validation. La política del writer conserva marcas preliminares y comentarios de ESTADO pendiente, mientras el asset productivo ESTADO ya es CONFIRMED. Esto no prueba que el filtro actual sea incorrecto: sí exige reconciliar la documentación y distinguir confirmación técnica de confirmación funcional. No se encontraron fundamentos para sustituir esas reglas por heurísticas nuevas.

8. PRIVACY / SECURITY REVIEW

El adapter conserva las columnas semánticas y no transmite datos por red en el flujo inspeccionado. Los cálculos necesitan fechas individuales en memoria; “nada personal entra al core” es una afirmación demasiado amplia de la documentación. Además, pd.read_excel materializa inicialmente todas las columnas y después selecciona las semánticas. Esto no demuestra persistencia o fuga.

Los modelos de resultados del cálculo son agregados y los problemas del analizador antiguo utilizan fila/campo/código. No encontré una fuga real de PII reproducida ni secretos que justifiquen un P0 en esta auditoría. Tampoco certifico ausencia universal de PII: humanize_error() registra el texto crudo de excepciones (errors.py:163), las etiquetas categóricas provienen del input y la privacidad ante parsers corruptos no quedó demostrada. Hace falta la prueba sintética de la matriz, sin usar pacientes reales.

Los logs rotan aproximadamente a 1 MB con dos backups. Los diagnósticos usan archivos de nombre fijo en el directorio de aplicación: no constituyen por sí solos una historia inmutable de cada informe. No se verificaron ACL Windows ni aislamiento efectivo entre usuarios. La retención de workspaces bloqueados es conservadora, pero F04 puede impedir la limpieza normal por terminación abrupta.

YAML usa safe_load, no se halló deserialización ejecutable en el flujo revisado, y los paths críticos se resuelven para varias comprobaciones de contención. No se confirmó explotación de symlinks/junctions. Las brechas de seguridad con soporte concreto son la sobrescritura concurrente F05 y la política de apertura F06. Las dependencias admiten rangos y no hay lockfile de build inspeccionado; eso afecta reproducibilidad, sin demostrar por sí mismo una CVE explotable.

9. WINDOWS / EXCEL REVIEW

Verificado por código y tests de adapter: instancia propia, CoInitialize, escritura Value2, llamada a CalculateFullRebuild, espera por CalculationState, cierre y liberación conservadora. Los tests simulados verifican muchos caminos de error. Esto constituye una base útil, sin equivaler a una certificación de Office.

El timeout de cálculo empieza después de que retorna CalculateFullRebuild; no acota una llamada COM que se quede bloqueada dentro de esa operación. Abrir/guardar/diálogos, archivos abiertos, antivirus, OneDrive, recursos de red, Protected View y distintas instalaciones Office siguen requiriendo pruebas Windows. No se afirmó un bloqueo reproducido donde sólo existe el camino posible.

A la pregunta “¿Existe algún cambio peligroso del template que hoy podría pasar las validaciones?”, la respuesta es sí: cambiar una expresión manteniendo su coordenada. F02 demuestra la aceptación por la huella; F07 muestra otros cambios que la comparación antes/después también puede perder. Debe comprobarse el contrato oficial además de preservar la referencia seleccionada.

Las pruebas Windows prioritarias son F03 fuera del checkout, F04 durante COM, F05 con dos procesos y F06 con un evento benigno. También deben comprobarse el resultado y sus celdas después de cerrar Excel, junto con la preservación semántica del VBA. No se debe matar Excel globalmente ni cambiar Trust Center para conseguir que pasen.

10. PACKAGING / DEPLOYMENT REVIEW

Los specs ONEDIR incluyen los directorios de runtime/writer y los imports COM explícitos; excluyen módulos Qt no usados y evitan distribuir datos locales. El entrypoint empaquetado es pequeño y delega en el mismo main. La resolución de assets principales no depende del cwd. La omisión y resolver de reglas en F03 son una excepción productiva importante.

No hay dist en este checkout. No se reconstruyó un .exe, ni se verificaron firma, reputación del binario o políticas de ejecución del equipo del piloto. La existencia de un build anterior no permite atribuirle automáticamente las propiedades de este HEAD.

La versión 0.1.0 aparece en distintos lugares; los diagnósticos registran run_id e integridad, pero no conforman un registro autocontenido que vincule para siempre output, versión exacta de aplicación, hashes de todos los assets e input. Para release: congelar el entorno de build, incluir todos los assets del cálculo, probar fuera del repo y registrar metadatos de generación sin PII. Mantener una versión anterior validada como rollback, sin sobrescribir archivos del usuario.

11. RECOMMENDED ROADMAP

NOW: reparar F01 y F03; establecer un ciclo de vida seguro de corridas F04 y publicación sin sobrescritura F05. Definir y verificar la compatibilidad semántica de plantilla F02. Resolver la política COM F06 con un ensayo Windows. Hacer visibles CONTROL/advertencias F08 y cerrar los falsos negativos de integridad F07 antes de confiar en el resultado del piloto. Corregir el false success PDF F09 es acotado y comprobable. Cada cambio debe incluir su repro como prueba de comportamiento, sin alterar el motor para mejorar porcentajes.

NEXT: completar la matriz Windows/Excel; confirmar tratamiento de registros inválidos, edad faltante y categorías nuevas; reconciliar docs y flags pendientes; vincular output y corrida con versiones/hashes seguros; asegurar privacidad de error paths y ACL/retención local. Preparar un dataset golden con resultado funcional aprobado para la parte Medinet.

LATER: incorporar fuentes adicionales sobre el contrato estabilizado; introducir soporte por versión anual; automatizar builds y pruebas Windows; ensayos prolongados y mejoras de rendimiento medidas. La detección de período y compatibilidad se hacen sincrónicamente al elegir archivos y el input se lee varias veces; medir antes de cachear o refactorizar. No se demostró un cuello O(n²) que justifique microoptimización ahora.

DO NOT DO: reescribir el .xlsm con openpyxl; matar procesos Excel globalmente; relajar seguridad de Office; introducir fuzzy matching clínico sin confirmación; ajustar edades/conteos para imitar una caché histórica; deduplicar atenciones sin un identificador/criterio funcional; añadir egresos para maquillar CONTROL; construir un framework nuevo mientras faltan estos contratos.

12. FINAL VERDICT

READY FOR DEVELOPMENT TESTING

El motor y el writer tienen fronteras aprovechables y una suite extensa. El estado observado no justifica un nuevo piloto interno operativo: hay pérdida de diagnósticos de input, una dependencia ausente del paquete declarado y fallos reproducidos de concurrencia/cierre. Las garantías de plantilla y apertura COM requieren trabajo y validación adicional. El alcance sólo-Medinet no exige implementar todas las fuentes; exige que sus datos y limitaciones se representen correctamente.

Dimensión

Nota / 10

Justificación

Corrección funcional

6

Cálculo/regresión sustanciales; F01 rompe el contrato de completitud visible y faltan decisiones sobre anomalías.

Confiabilidad

4

Writer sobre copia con verificaciones, pero F04/F05 permiten aborto, resultados obsoletos y reemplazo concurrente.

Privacidad

6

Resultados agregados y procesamiento local; error paths y ACL Windows no demostrados. Sin fuga real confirmada.

Seguridad

5

YAML seguro e instancias propias; política de macros implícita y publicación con sobrescritura bajo carrera.

Testing

7

923 pruebas pasan y hay histórico; huecos relevantes en integración UI, empaquetado aislado y Excel real.

Arquitectura

7

Separación cálculo/escritura adecuada; contratos de validación y corrida necesitan consolidación.

Mantenibilidad

6

Núcleo reutilizable, pero configuración por año y documentación desfasada dificultan conocer el runtime real.

UX

5

Flujo claro en camino normal; cancelación/cierre incorrectos y advertencias omitidas.

Packaging

3

Specs y smoke presentes, pero la distribución declarada no cubre todas las dependencias del cálculo.

Readiness piloto

4

Requiere resolver bloqueadores y repetir una generación real en máquina aislada.

Readiness producción

2

No hay base suficiente para operación rutinaria confiable, aunque el MVP se limite a Medinet.

Las notas son juicios sobre la evidencia de este HEAD, no porcentajes de exactitud ni probabilidades de fallo. No se suman para obtener una aprobación.

Ejecuciones y evidencias reproducibles

Entorno: Linux, Python 3.12.3; pandas 2.3.3, openpyxl 3.1.5, PySide6 6.11.2, PyYAML 6.0.3, olefile 0.47, pytest 8.4.2, ruff 0.16.5.

Validación

Resultado

PYTHONDONTWRITEBYTECODE=1 QT_QPA_PLATFORM=offscreen .venv/bin/python -m pytest -ra -W default -p no:cacheprovider --basetemp=/tmp/remasep-audit-pytest

926 recolectados; 923 passed, 3 skipped, 1 warning; 164,13 s.

Warning visible

ResourceWarning: unclosed file ... /dev/null en test_guard_frozen_streams_replaces_none_streams. No se ocultó ni se convirtió en hallazgo alto.

Tests omitidos

Dos de test_excel_com_integration.py por plataforma; uno AF/Excel opt-in por variable de entorno.

.venv/bin/ruff check --no-cache .

Aprobado.

git diff --check

Aprobado.

Compilación en memoria

130 archivos Python de src/scripts/tests/packaging, sin generar bytecode.

Imports productivos

main, production_pipeline, excel_writer, workers, excel_com importables en Linux.

Bundle

Cinco hashes verificados, 1.122 instrucciones y 1.224 fórmulas.

--smoke-mode desde checkout con appdata en /tmp

SMOKE OK; su limitación se demuestra en F03.

Determinismo

Repetición de los cuatro escenarios sintéticos del pipeline con mismos valores por instrucción. No demuestra exhaustividad.

Estado final Git

Limpio; mismo HEAD.

Repro del pipeline y resultados.

Repros de UI, packaging y PDF y resultados. Usan Qt real; el trabajo de larga duración se controla con eventos. El caso cierre espera terminar con aborto y se ejecuta en subproceso.

Repros del writer y resultados. Para las comprobaciones de huella requieren los workbooks locales indicados; nunca los guardan. Las decisiones del orquestador usan FakeWriterHarness; la promoción concurrente utiliza archivos y replace reales.

Los scripts de reproducción se ejecutan desde la raíz del proyecto con PYTHONDONTWRITEBYTECODE=1 .venv/bin/python . Sus fixtures temporales se limpian. Los archivos de informe y evidencia permanecen en /tmp para revisión; no se han añadido al repositorio. La auditoría termina aquí, sin aplicar correcciones.
