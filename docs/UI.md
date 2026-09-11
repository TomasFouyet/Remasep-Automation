# UI de escritorio — flujo MEDINET + resumen mensual

Sprint 3.10. Convierte la maqueta PySide6 en una aplicación de escritorio
**profesional y simple** para una persona administrativa, sin terminal.

La UI **consume** el backend validado (`medinet_summary`, `production_pipeline`,
`GenerationService`); **no** recalcula reglas REMASEP ni reimplementa el writer
de Excel. El CLI (`scripts/generate_remasep.py`) sigue funcionando igual.

---

## Flujo

```
Inicio → Nuevo informe → Análisis → Resumen mensual (Medinet) → Generar → Resultado
```

| Pantalla | Archivo | Qué hace |
|---|---|---|
| **Inicio** | `screens/home.py` | Título `REMASEP`, subtítulo, botón **Nuevo informe**. Card discreta "Último informe generado" (sólo si hay `outputs/REMASEP_*.xlsm`, con **Abrir carpeta**). |
| **Nuevo informe** | `screens/new_report.py` | Período (mes/año) con **detección asistida** al elegir el Medinet (ver abajo), selector **Archivo de Medinet** (`.xlsx`/`.xlsm`, muestra nombre + estado), selector **Plantilla REMASEP** (`.xlsm`, valida el fingerprint con `check_template_compatibility`). **Analizar datos** se habilita sólo con ambos archivos válidos **y** un período presente en el archivo. |
| **Análisis** | `screens/analysis.py` | Progreso comprensible (`ANALYSIS_STEPS`, sin logs). Al terminar: `state.summary` → Resumen. Ante error: `ErrorBanner` con mensaje humano + acción "Elegir otro archivo". |
| **Resumen mensual** | `screens/dashboard.py` | Fuente **Medinet** únicamente. 4 KPIs + 4 gráficos agregados (sin PII). Botones **Exportar resumen PDF** y **Generar REMASEP**. |
| **Generar** | `screens/generate.py` | **Diálogo "Guardar como"** (nombre sugerido `REMASEP_YYYY_MM_DRAFT.xlsm`) → progreso (`GENERATION_STEPS`) → `GenerationService` con esa ruta explícita. Resultado: nombre / período / atenciones consideradas / integridad; botones **Abrir Excel** / **Abrir carpeta** / **Volver al inicio**. Aviso `NOT_FOR_SUBMISSION` (no es un error). Ante error del backend: mensaje humano; si el nombre ya existe, **"Elegir otro nombre"** reabre "Guardar como" (no vuelve a Nuevo informe). |

Navegación: `MainWindow` + `QStackedWidget`; cada pantalla expone `.name` y
`.on_enter()`. `AppState` (dataclass) comparte período / rutas (`medinet_path`,
`template_path`, `output_path`) / `summary` / `analysis_error` / `generation`.
`reset_flow()` vuelve al inicio conservando el período y limpiando los selectores.

### Detección asistida del período (`detect_medinet_periods`)

`src/remasep/services/medinet_analysis.py::detect_medinet_periods(path)` devuelve
los meses/años con al menos una cita **estructuralmente válida** (misma noción que
`processing_scope_frame`), en orden cronológico. **No** usa el nombre del archivo;
un export puede traer varios meses. Al elegir el archivo Medinet:

- **Un solo período** → se selecciona automáticamente; estado `✓ Período detectado: Julio 2026`.
- **Varios períodos** → si el período actual está presente se conserva; si no, se
  elige el **más reciente** (determinista) y el estado pide `Confirma el mes que deseas procesar.`
- **Período elegido ausente del archivo** → **"Analizar datos" deshabilitado** +
  `No encontramos citas de [Mes Año] en este archivo.` (no se espera a una pantalla posterior).
- **Sin fechas válidas / archivo ilegible** → mensaje humano, sin traceback, análisis bloqueado.

`processing_scope_frame` sigue siendo la autoridad al procesar; esta detección
sólo asiste la elección en la UI.

### "Guardar como" antes de generar

Al pulsar **Generar REMASEP**, `GenerateScreen.on_enter()` abre
`QFileDialog.getSaveFileName` (`_prompt_output_path`) **antes** de iniciar Excel COM:

- Nombre sugerido `REMASEP_{año}_{mes}_DRAFT.xlsm`; el usuario puede cambiar
  nombre/carpeta o **cancelar** (vuelve al Resumen, sin generar, sin error).
- La ruta elegida se pasa **explícita** a `GenerationWorker` → `run_generation` →
  `GenerationService.generate(output_path=...)`.
- **Nunca sobrescribe**: si el nombre ya existe, el diálogo lo explica y vuelve a
  pedir (`DontConfirmOverwrite`). Si aun así el writer devuelve
  `OUTPUT_ALREADY_EXISTS`, el botón **"Elegir otro nombre"** reabre "Guardar como"
  conservando Medinet / plantilla / período / análisis. No se borran archivos
  previos. La protección de no-overwrite del writer **no se modificó**.
- Una segunda generación del mismo mes vuelve a pedir "Guardar como" (la ruta
  anterior ya existe), así `OUTPUT_ALREADY_EXISTS` no aparece como pantalla
  terminal en un flujo normal.

---

## Modelo de datos del resumen — `MonthlyMedinetSummary`

`src/remasep/services/medinet_summary.py`. **Adaptador de sólo lectura**: reusa
`processing_scope_frame` (válidas ∩ período) + `apply_estado_filter` (regla
ESTADO confirmada, Sprint 3.9) + agregación pandas. Independiente de Qt, sin PII,
testeable. La UI y el PDF consumen **el mismo** summary.

Campos (todos agregados): `period_label/_year/_month`, `period_scope_records`,
`included_records`, `excluded_records`, `estado_distribution`
(`EstadoCategoryCount`: label/count/**included**), `sex_distribution`,
`age_distribution`, `service_distribution` (top 8 especialidades + "Otras"),
`estado_included_states` / `estado_excluded_states`, `estado_filter_status`,
`pending_sources`, `source="Medinet"`. Propiedades: `included_percentage`
(1 decimal), `considered_ratio_label` (`"1364 / 2114"`), `as_dict()` (sólo claves
agregadas).

Invariante: `included_records + excluded_records == period_scope_records`.

Julio 2026 (archivo real): scope **2 114**, consideradas **1 364**, excluidas
**750**, **64,5 %** — idéntico a `build_production_pending_writes`
(`tests/test_medinet_summary.py::test_july_2026_summary_matches_production_pipeline`).

### Métricas y gráficos (y por qué)

- **KPIs**: Citas del período · Consideradas para REMASEP · Excluidas por estado ·
  Porcentaje considerado.
- **Distribución por ESTADO** (obligatorio): color distingue consideradas
  (`chart_included`) de excluidas (`chart_excluded`).
- **Por especialidad** (`ESPECIALIDAD`): 16 valores limpios en el archivo real →
  fiable; top 8 + "Otras especialidades".
- **Por sexo** (`SEXO`): 0 vacíos → fiable.
- **Por tramo de edad**: `legacy_transform.legacy_age_years(nac, cita)`, 7 bandas;
  0 sin dato en el archivo real.
- **`PRESTACIÓN` se descarta**: 302 / 1 364 vacías → no fiable (criterio del spec:
  no inventar categorías, no mostrar lo que no es fiable).

### Privacidad

Sólo datos agregados. **Nunca** nombre / RUT / fecha de nacimiento individual /
teléfono / e-mail / identificador personal; sin tablas de pacientes. Cubierto por
`test_summary_has_no_pii` y `test_report_content_and_pdf_have_no_pii`.

---

## Export PDF — `src/remasep/ui/pdf_report.py`

`QPdfWriter` + `QPainter` (nativos de PySide6). **Sin WebEngine / Chromium.** No
es un screenshot: se compone un reporte limpio de **una página** A4.

- `report_content(summary) -> ReportContent` — arma **qué** texto/series van
  (testeable sin pintar): título `REMASEP`, `Resumen mensual · Fuente: Medinet`,
  período, 4 KPIs, 4 gráficos, pie `Fuente: Medinet` + `Generado por REMASEP
  Automation`, nota "…únicamente a datos provenientes de Medinet…".
- `export_summary_pdf(summary, path) -> Path` — sólo **pinta**. Lanza
  `FileExistsError` si el destino existe (sin sobrescritura silenciosa).
- `suggested_pdf_name(summary)` → `Resumen_Medinet_2026_07.pdf`.

La pantalla usa `QFileDialog.getSaveFileName` (sugerido en `~`), fuerza `.pdf` y
mapea `FileExistsError` a un aviso ("Ya existe un archivo…").

---

## Generación del REMASEP desde la UI

`src/remasep/ui/workers.py::run_generation` (función pura, sin Qt) reusa
`load_runtime_bundle` + `build_production_pending_writes` + `GenerationService.generate`
(`mode=PRODUCTION`). **No** reimplementa Excel COM, **no** cierra procesos de
Excel del usuario. En Linux / sin Excel devuelve `ok=False` con un `HumanError`
legible ("Microsoft Excel Desktop es necesario…") — no lanza.

`GenerationOutcome`: `ok`, `status`, `submission_label`, `output_path`,
`written_cells`, `considered_records`, `integrity_ok`, `control_status`,
`warnings`, `human_error`.

Abrir archivo / carpeta: `src/remasep/ui/open_location.py` con
`QDesktopServices.openUrl(QUrl.fromLocalFile(...))` (API segura).

---

## Manejo de errores — `src/remasep/ui/errors.py`

`HumanError(title, detail, action_label, technical)`. El detalle técnico (código
backend, clase de excepción) va **sólo** al log de dev (`_log.warning`), nunca a
la persona usuaria; sin tracebacks en pantalla.

- `humanize_generation_status(status)` — mapea `STATUS_EXCEL_UNAVAILABLE` /
  `STATUS_TEMPLATE_INCOMPATIBLE` / `STATUS_OUTPUT_EXISTS` / … a mensajes humanos.
- `humanize_error(exc)` — `RuntimeAssetError` (`RUNTIME_ASSET_MISSING` →
  "La instalación está incompleta…"), `ExcelWriterError` (`TEMPLATE_INCOMPATIBLE`),
  errores de lectura de Medinet → "No pudimos leer el archivo… / Elegir otro
  archivo", y un fallback genérico.

---

## Componentes reutilizables — `src/remasep/ui/components/`

`Card` / `SectionHeader` / `StatusBadge` / `MetricCard` / `EmptyState` /
`ErrorBanner` (`widgets.py`), `FileSelector` (`file_selector.py`),
`StepIndicator` (`step_indicator.py`), `BarChartCard` (`ui/charts.py`, barras
horizontales en QPainter puro — sin lib de gráficos). Todo el estilo vive en el
QSS central `ui/styles.py` (`STYLESHEET` + `Theme`/`THEME`); **sin QSS duplicado**.

---

## Tests

Todos en modo *offscreen* (`conftest.py`), sin Excel COM:

| Archivo | Cubre |
|---|---|
| `tests/test_medinet_summary.py` | Totales e invariantes del modelo, ESTADO incluido/excluido, período, sin PII, `pending_sources`, equivalencia con el pipeline (julio real). |
| `tests/test_pdf_report.py` | Nombre sugerido, PDF generado (`%PDF-`), sin sobrescritura silenciosa, `ReportContent` con período/KPIs/pie, sin PII. |
| `tests/test_ui_errors.py` | Mapeo de errores backend → `HumanError`. |
| `tests/test_medinet_analysis.py` | `detect_medinet_periods`: un período → autodetección; varios → orden cronológico; ignora filas estructuralmente inválidas; sin fechas → `[]`; archivo ilegible → `SourceValidationError`; **no usa el nombre del archivo**; coincide con `processing_scope_frame`. |
| `tests/test_ui_shell.py` | Arranque en `home`, navegación, período por defecto (mes anterior), selección/validación de archivos, `StepIndicator`, **detección asistida del período** (un mes / varios / ausente / sin fechas / nombre ignorado), análisis → dashboard con KPIs que coinciden con el summary, error de análisis → `ErrorBanner` humano, **"Guardar como"** (ruta elegida se pasa explícita / cancelar vuelve al Resumen sin error / "Elegir otro nombre" reabre el diálogo y no va a Nuevo informe), generación llama a `run_generation` (mock), error backend → mensaje humano, `run_generation` sin Excel es *graceful*, `reset_flow`. |

Lanzar la app: `python -m remasep.main`.

---

## Nota — warning `QFont::setPointSize: Point size <= 0 (-1)` (Windows)

Causa: reglas del QSS central con `font-weight` / `letter-spacing` **sin** un
`font-size` explícito. En esos casos `QStyleSheetStyle` construye un `QFont`
cuyo `pointSize()` es `-1` y lo reaplica, emitiendo el warning (visible sobre
todo en el motor de fuentes de Windows). Corregido añadiendo `font-size: 14px`
(el tamaño que ya heredaban de `QWidget`) a las tres reglas afectadas:
`QPushButton[variant="primary"]`, `QLabel#stepActive`, `QLabel#stepDone`. Sin
cambio visual; no se silencia ningún warning globalmente.
