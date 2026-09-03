# Información pendiente

## Medinet
- Export mensual real.
- Columnas/hoja.
- Estados válidos.
- Anuladas/inasistencias.
- Duplicados.
- Varias prestaciones por cita.
- Telemedicina.
- Cálculo de edad.
- Campos faltantes.
- Responsable de nuevas clasificaciones.

### Observaciones de Sprint 2.1 (a confirmar)
- **MODALIDAD**: en el workbook de referencia la columna con encabezado
  `MODALIDAD` contiene categorías de previsión (`Fonasa A/B/C/D`, `GES …`,
  `Particular/Libre Elección`, `FFAA`, …), no `presencial`/`telemedicina`.
  Por lo tanto **no parece representar la modalidad de atención**.
  Confirmar que presencial/telemedicina debe derivarse mediante `SUCURSAL`, como
  hace actualmente `AC = TIPO_DE_CITA + SUCURSAL`. (No se renombra el campo: es
  el encabezado real de Medinet.)
- **ESTADO**: existe con categorías propias (`Atendido`, `Atención Pausada`,
  `En Sala de Espera`, `En Atención`) pero no participa en las dependencias
  REMASEP observadas. ¿Debe filtrarse por estado antes de contar?
- **PRESTACION vacía**: ~302 de 1.364 atenciones reales del archivo de
  referencia tienen `PRESTACION` vacía. No hay evidencia de que sea obligatoria.
- **Cobertura legacy**: las 27 reglas AG:AL clasifican 485 de 1.364; el resto
  queda como "válido no cubierto por reglas legacy". Falta el catálogo completo
  de prestaciones.

## Egresos
- Archivo real.
- Campos.
- Reglas cirugía plástica.
- Reglas otorrino.
- Indicadores/celdas que completa Jacqueline.

## Recursos/pabellones
- Fuente real.
- Archivo/reporte.
- Campos.
- Qué es constante y qué cambia mensualmente.

## Golden dataset
Por al menos 2-3 meses:
- Medinet.
- Egresos.
- Recursos.
- REMASEP final.
- Identificar celdas modificadas manualmente.

## Infraestructura
- Windows.
- Excel Desktop.
- Políticas de macros.
- Permisos.
- Antivirus/SmartScreen.
- Firma del ejecutable.
- Carpeta institucional.
