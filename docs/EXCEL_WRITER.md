# Excel COM writer (Sprint 3.7B)

`src/remasep/services/excel_writer.py` (núcleo) +
`src/remasep/adapters/excel_com.py` (COM, Windows) +
`src/remasep/services/generation_service.py` (orquestador / API para la UI) +
`src/remasep/services/excel_capability.py` +
`scripts/generate_remasep.py` (CLI) +
`config/excel_writer_2026/`.

Escribe los `PendingWrite` del Sprint 3.7A en una **copia** del REMASEP oficial
usando Microsoft Excel Desktop. El writer **no** recalcula MetricValues: sólo
consume el contrato. **No** conecta la UI. El archivo generado se marca
`NOT_FOR_SUBMISSION` mientras el filtro por ESTADO siga pendiente de confirmación
funcional.

## Requisito de plataforma

`ExcelCapability` (`detect_excel_capability`) diagnostica `platform`,
`pywin32_available`, `excel_com_available`, `excel_version`, `can_generate`,
`reason`. En Linux/WSL: `can_generate = False`,
`reason = WINDOWS_EXCEL_REQUIRED`, mensaje al usuario sin traceback. La CLI
termina con código 3. La generación real sólo ocurre en Windows con
`pywin32` + Excel instalados.

Ningún módulo importa `win32com` a nivel de módulo: `import`ar el writer, el
adaptador COM o la CLI funciona en cualquier sistema operativo.

## Arquitectura por capas

| capa | módulo | plataforma |
| --- | --- | --- |
| contrato + preflight + verificación + orquestación | `services/excel_writer.py` | cualquiera |
| implementación COM del contrato | `adapters/excel_com.py` (`ExcelComWorkbookWriter`) | Windows |
| doble de test | `remasep/testing/fake_workbook_writer.py` (`FakeWorkbookWriter`) | cualquiera |
| capacidad de plataforma | `services/excel_capability.py` | cualquiera |
| orquestador + artefactos + API UI | `services/generation_service.py` | cualquiera |

`WorkbookWriter` es un `Protocol`: `sheet_names`, `has_sheet`,
`cell_has_formula`, `cell_in_incompatible_merge`, `cell_is_writable`,
`read_cell`, `write_value2`, `recalculate`, `save`, `snapshot`. El núcleo abre
Excel **nunca**; sólo lo hace la implementación COM.

## Ciclo de vida de la instancia Excel

`win32com.client.DispatchEx("Excel.Application")` — **instancia aislada**, nunca
`Dispatch()` compartido ni `GetObject`. Sobre esa instancia: `Visible = False`,
`DisplayAlerts = False`. Al terminar se cierra **sólo** el workbook y la
instancia creados por nosotros (`Workbook.Close(SaveChanges=False)` +
`Application.Quit()` + `CoUninitialize`). Nunca `taskkill` / matar `Excel.exe` /
tocar procesos del usuario.

## Macros / seguridad

Se preserva el proyecto VBA (`.xlsm`, `keep_vba`). **No** se ejecuta ninguna
macro, **no** se llama `Application.Run`, **no** se tocan Trusted Locations,
registro, macro security, Trust Center ni Protected View.

## Copy-first + salida atómica

```
template ──copy2──▶ output.__working__.xlsm ──write/recalc/save──▶ (verificación) ──replace──▶ output.xlsm
```

Nunca se abre `template_path` en modo escritura. Se verifica el SHA256 de la
plantilla antes y después: debe quedar idéntico (`template_unchanged`). Si algo
falla se borra la copia de trabajo, la plantilla queda intacta y **no** se
produce salida final parcial. La promoción a `output.xlsm` es un `Path.replace`
(rename atómico) y sólo ocurre si toda la verificación pasó.

## Preflight (antes de tocar nada)

`run_preflight` aborta con `ABORTED_PREFLIGHT` si: la plantilla no existe / no es
`.xlsm`; `output == template`; la salida cae en `data/` o `inputs/`; la salida
ya existe (`OUTPUT_ALREADY_EXISTS`); hay `instruction_id` o celdas destino
duplicadas; algún valor no es `int >= 0` para `INTEGER_COUNT`
(`VALUE_TYPE_MISMATCH` / `NEGATIVE_COUNT`); falta algún `MetricValue`
(`MISSING_WRITE_INSTRUCTIONS`); o la `zero_write_policy` no está entre las
aceptadas (`WRITE_ZERO` / `FORM_SPECIFIC`).

## Compatibilidad de plantilla

`structural_template_fingerprint` (Sprint 3.6). El esperado se **lee del
manifiesto** (columna `template_fingerprint_id`, `stf:dc624775927d4d4d`); el
`policy.yaml` sólo lleva un fallback informativo. Si el fingerprint estructural
no coincide → `TEMPLATE_INCOMPATIBLE`, no se escribe. El SHA256 del archivo
puede cambiar por un re-save/VBA y **no** es la compatibilidad principal.

## Defense-in-depth por celda (§9)

Antes de cada escritura: la hoja existe, la celda no tiene fórmula
(`TARGET_FORMULA_CONFLICT` → abort, **nunca** se sobrescribe una fórmula), no
está en un merge incompatible (≠ 1×1), y es escribible (hoja desprotegida o
celda `Locked = False`). No se cambia ninguna protección.

## Escritura de valores (§10)

`INTEGER_COUNT` → `int >= 0`, escrito con `Range.Value2`. El **0 se escribe
explícito**; nunca blanco por 0 (`zero_write_policy = WRITE_ZERO`, evidencia del
3.7A). No hay coerción silenciosa: `1.8 → 1` o `"3" → 3` abortan en preflight.

## Recálculo

Tras escribir: `Application.CalculateFullRebuild()` y espera activa hasta
`CalculationState == xlDone` antes de guardar. No se confía en caches previas.

## Verificación posterior (§17/§18/§19)

Se reabre la copia de trabajo en modo inspección (openpyxl, estático) y se
compara contra el snapshot previo:

- **fórmulas**: `compare_formula_integrity` — 0 `FORMULA_EXPRESSION_CHANGE`,
  0 añadidas, 0 eliminadas; cualquier cambio → `GENERATION_FAILED_INTEGRITY_CHECK`.
- **VBA**: presente antes y después; si se puede, hash estable del payload
  `xl/vbaProject.bin` comparado.
- **celdas destino**: cada `PendingWrite` está en su celda con su valor; ninguna
  se volvió fórmula; ninguna falta.
- **fingerprint estructural**: sigue siendo compatible.
- **hojas**: no falta ninguna.

`writer_integrity_status` ∈ `WRITER_INTEGRITY_PASS` / `WRITER_INTEGRITY_FAIL`.

## CONTROL (§21/§22)

Mapa versionado en `config/excel_writer_2026/control_map.yaml` (layout observado:
`D6:D14` etiquetas, `E6:E14` Nº errores por hoja, `F7:F14` "OK"/"SIN DATOS",
`G8:G14` causales, `E16` total). Tras recalcular se lee la hoja y se arma
`ControlResult`:

- `PASS_INTERNAL_VALIDATION` (total = 0) · `FAIL_INTERNAL_VALIDATION` (total > 0)
  · `CONTROL_UNAVAILABLE` (hoja/celda ausente o sin recalcular).
- `errors_from_pending_modules`: cuántos errores provienen de hojas que este
  sprint todavía **no** produce (URGENCIAS, EyP_ET, TV_MI, SERV_SANGRE, NOMBRE,
  REMASEP B1).

**`writer_integrity_status` es independiente de `control_status`.** En la
primera generación sólo-MEDINET es esperable `CONTROL_FAIL` porque faltan
EGRESOS / recursos / tabla quirúrgica / metadatos. Eso **no** significa que el
writer falló. `PASS_INTERNAL_VALIDATION` ≠ `MINSAL_APPROVED`.

## Modos de generación

| modo | MetricValues de | scope | estado |
| --- | --- | --- | --- |
| `DIAGNOSTIC_REFERENCE` | `LEGACY_EQUIVALENCE_DIAGNOSTIC` | 1364 | `NOT_FOR_SUBMISSION` — para probar el writer end-to-end |
| `PRODUCTION` | `PRODUCTION_PERIOD_SCOPE` | 2114 (sin filtro ESTADO) | `DRAFT` / `NOT_FOR_SUBMISSION` |

Política conservadora (`policy.yaml`): mientras `estado_filter_pending: true`,
`final_ready_allowed: false` — **nunca** se emite `FINAL_READY`. Se permite
`GENERATED_DRAFT` para pruebas.

## Referencia histórica obsoleta (Sprint 3.7A §13)

El writer escribe el `MetricValue` **producido por el motor actual**, nunca el
valor cacheado histórico para forzar igualdad con `REMASEP_V1.4 Julio 2026.xlsm`.
Se registra `historical_reference_cache_status =
KNOWN_STALE_FOR_142_WRITE_READY_VALUES` (142 celdas con caché legacy obsoleta de
`AF`/edad). No se "corrigen" MetricValues para parecerse al archivo histórico.

## Salida

`outputs/` (gitignored), nunca `data/local/`. Nombre determinista
`REMASEP_2026_07_DRAFT.xlsm` (luego `REMASEP_2026_07.xlsm`). Un archivo existente
**no** se sobrescribe: `OUTPUT_ALREADY_EXISTS`.

## Artefactos (`artifacts/excel_writer/`, gitignored)

`generation_summary.json`, `write_audit.csv` (sin PII: `instruction_id`,
`source_metric_id`, `target_sheet`, `target_cell`, `written_value`,
`write_status`), `target_verification.csv`, `formula_integrity.csv`,
`control_result.csv`, `template_verification.json`, `README.md`.

## API para la UI (§31)

`GenerationService.generate(...)` → `GenerationServiceResult` con `status`,
`output_path`, `written_cells`, `control_status`, `warnings`, `errors`,
`capability`. El botón "Generar REMASEP" del sprint siguiente llamará aquí; hoy
**no** está conectado.

## Estado ESTADO — pendiente

`estado_filter_status = PENDING_FUNCTIONAL_CONFIRMATION`. El modo productivo no
aplica el filtro; el diagnóstico lo usa sólo para el estudio de equivalencia. El
archivo generado nunca es un entregable MINSAL en este sprint.

## Qué falta para el primer Excel real

1. Ejecutar `scripts/generate_remasep.py --mode diagnostic` en **Windows** con
   Python + `pywin32` + Microsoft Excel Desktop, con acceso a
   `data/local/REMASEP_V1.4 Julio 2026.xlsm`, `GENERACION DATOS REMASEP.xlsx` y
   el export `detalle_citas`.
2. Revisar `artifacts/excel_writer/` (integridad de fórmulas/VBA, verificación de
   las 1122 celdas, CONTROL).
3. Confirmar funcionalmente el filtro por ESTADO para habilitar el modo
   productivo real (fuera del alcance de este sprint).
