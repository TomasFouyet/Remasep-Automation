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

Dependencia añadida en 3.7B: **`olefile`** (`>=0.46`, pura Python, ~114 KB) para
la comparación semántica de VBA. Si falta, la verificación es **fail-closed**
(`VBA_SEMANTIC_CHECK_UNAVAILABLE` → no se genera la salida).

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
`DisplayAlerts = False`. Nunca `taskkill` / matar `Excel.exe` / tocar procesos
del usuario.

**Cierre determinista** (patch de cierre 3.7B). El desmontaje sigue este orden
exacto para no dejar proxies COM que el GC de Python liberaría (`Release()`)
*después* de desmontar el apartment — la causa de los `RPC_E_DISCONNECTED`
(`0x80010108`) / `RPC server unavailable` (`0x800706ba`) / `RPC call failed`
(`0x800706be`) que aparecían durante el GC:

1. `CoInitialize()` en `__enter__`, siempre balanceado con `CoUninitialize()` en
   `_teardown()` (flag `_co_initialized`).
2. `Workbook.Close(SaveChanges=False)` + soltar el proxy del workbook y del caché
   de hojas (`= None`) + `gc.collect()`.
3. `Application.Quit()` sobre **esa** instancia, defensivo ante `com_error`
   (Excel puede haber muerto).
4. Soltar el proxy de `Application` (`excel = None`) y `gc.collect()` — fuerza su
   `Release()` **con el apartment todavía inicializado**.
5. Recién entonces `CoUninitialize()`.

`describe_com_error()` (en `excel_writer.py`) resume cualquier excepción COM a
`Clase(0xHRESULT)` sin volcar traceback ni importar `pythoncom` a nivel de
módulo. `_probe_excel_com()` de `ExcelCapability` usa el mismo patrón, de modo
que llamarlo repetidas veces es estable y no deja procesos huérfanos.

## Macros / seguridad

Se preserva el proyecto VBA (`.xlsm`, `keep_vba`). **No** se ejecuta ninguna
macro, **no** se llama `Application.Run`, **no** se tocan Trusted Locations,
registro, macro security, Trust Center ni Protected View.

## Copy-first + workspace por corrida + salida atómica

```
template
  └─copy2─▶ outputs/.remasep-tmp/<run_id>/working.xlsm   ← workspace propio y ÚNICO
              └─ write / recalc / save (Excel COM)
              └─ verificación (openpyxl, estático)
                   └─ os.replace ──▶ outputs/REMASEP_2026_07_DRAFT.xlsm   (atómico)
                        └─ rmtree(workspace)  (best-effort)
```

Nunca se abre `template_path` en modo escritura ni se usa un temporal
compartido: cada generación crea `outputs/.remasep-tmp/<run_id>/`
(`run_id` = timestamp + `uuid4`, `mkdir(exist_ok=False)`) — dos corridas **jamás**
comparten ruta y una corrida nunca pisa ni borra el temporal de otra. La
promoción a la salida final es `Path.replace` (= `os.replace`, rename atómico)
con **reintento acotado** (`_PROMOTE_ATTEMPTS = 5`, sin loop infinito) por si
Windows conserva un handle momentáneo tras cerrar Excel; sólo ocurre si toda la
verificación pasó.

**Cleanup conservador** (`cleanup_run_workspace`): borra **sólo** su propio
`<run_id>/` (salvaguarda `_is_within(root, .remasep-tmp)`; si no, `CLEANUP_REFUSED`).
Si Windows retiene un handle (`WinError 32` / `PermissionError`) **no** se fuerza,
**no** hay `taskkill`, **no** hay loop: se devuelve `cleanup_status =
CLEANUP_PENDING` + `workspace_path` (sólo para diagnóstico local) + un warning;
la generación no se marca fallida por eso. No se muestra traceback al usuario.

El SHA256 de la plantilla se verifica antes y después (`template_unchanged`). Si
algo falla, la plantilla queda intacta y **no** hay salida final parcial.

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

Tras escribir: `Application.Calculation = xlAutomatic` +
`Application.CalculateFullRebuild()`, y luego una **espera real** hasta
`Application.CalculationState == xlDone` antes de guardar:

- bucle con `time.monotonic()` y `time.sleep(0.1 s)` entre consultas;
- timeout de **120 s** (`RECALC_TIMEOUT_SECONDS`); si expira →
  `ExcelWriterError` explícito y **no se guarda** el workbook;
- si `CalculationState` lanza una excepción COM durante el *polling*, se
  propaga como `ExcelWriterError` (no se oculta: hay que saber si el cálculo
  terminó).

No se confía en caches previas. `time.sleep` / `time.monotonic` son puntos de
inyección (`_sleep` / `_monotonic`) para los tests unitarios.

## Verificación posterior (§17/§18/§19)

Se reabre la copia de trabajo en modo inspección (openpyxl, estático) y se
compara contra el snapshot previo:

- **fórmulas**: `compare_formula_integrity` — 0 `FORMULA_EXPRESSION_CHANGE`,
  0 añadidas, 0 eliminadas; cualquier cambio → `GENERATION_FAILED_INTEGRITY_CHECK`.
- **VBA**: comparación **semántica** por módulo (ver abajo).
- **celdas destino**: cada `PendingWrite` está en su celda con su valor; ninguna
  se volvió fórmula; ninguna falta.
- **fingerprint estructural**: sigue siendo compatible.
- **hojas**: no falta ninguna.

`writer_integrity_status` ∈ `WRITER_INTEGRITY_PASS` / `WRITER_INTEGRITY_FAIL`.

### Integridad semántica de VBA (`remasep.services.vba_integrity`)

Comparar el **SHA256 bruto** de `xl/vbaProject.bin` es un criterio erróneo: un
`Save` legítimo de Excel **regenera** los streams de caché compilada
(`__SRP_0..n`, `_VBA_PROJECT`, la P-code de cada módulo) sin que cambie una línea
de macro. En el caso real reportado eso producía un falso
`GENERATION_FAILED_INTEGRITY_CHECK: VBA alterado/ausente` aun con `1122/1122`
celdas verificadas y fórmulas intactas.

En su lugar se extrae, **sólo lectura**, el código fuente descomprimido por
módulo:

- `olefile` (dependencia lean, pura Python, ~114 KB) abre el OLE Compound File;
- se descomprime el *CompressedContainer* de cada stream `VBA/<módulo>` según
  **MS-OVBA §2.4.1** (`decompress_vba_container`) — sin ejecutar nada, sin COM,
  sin VBIDE, sin habilitar "Trust access to the VBA project object model", sin
  tocar el Trust Center ni la protección del proyecto;
- se separa el bloque `Attribute VB_*` del **cuerpo de código** y se hashea cada
  parte (`code_sha256`, `attributes_sha256`).

Campos: `vba_present_before/after`, `vba_binary_payload_sha_before/after`,
`vba_binary_payload_stable`, `semantic_available`, `module_names_before/after`,
`added_modules`, `removed_modules`, `code_changed_modules`,
`attributes_changed_modules`.

`status` ∈ **`VBA_INTEGRITY_PASS`** / **`VBA_INTEGRITY_FAIL`** (sólo dos).

**FALLA** (`VBA_INTEGRITY_FAIL`) si: el proyecto VBA existía y **desaparece** (o
aparece); **cambia el conjunto de módulos**; **cambia el código** de un módulo
(`code_sha256`); **o** el VBA existe pero **no se pudo verificar el código**
(sin `olefile` / parseo fallido) → `reason = VBA_SEMANTIC_CHECK_UNAVAILABLE`
(**fail-closed**).

Quedan como **diagnóstico/warning** (no fallo) **únicamente cuando la
comparación semántica sí se ejecutó y demostró equivalencia** (`semantic_available`,
módulos iguales, `code_changed_modules == ()`): el cambio sólo del binario
`vbaProject.bin` (`VBA_BINARY_PAYLOAD_CHANGED_SEMANTIC_EQUIVALENT`) y el
reordenamiento de atributos/controles (`VBA_MODULE_ATTRIBUTES_CHANGED` — p.ej.
las líneas `Attribute VB_Control` de los botones ActiveX de una hoja).

**Fail-closed en `VBA_SEMANTIC_CHECK_UNAVAILABLE`**: si el proyecto VBA existe
pero `olefile` no está instalado o el `vbaProject.bin` no se puede parsear, el
writer **falla** y **no promueve** el workbook al output final. **Nunca** se
asume equivalencia sólo porque `vbaProject.bin` siga presente. Se conservan
presencia + hashes binarios como evidencia en un mensaje legible
(`GENERATION_FAILED_INTEGRITY_CHECK: VBA_SEMANTIC_CHECK_UNAVAILABLE: …`).

Reproducción (Linux, sin Excel): `read_vba_project()` sobre
`data/local/REMASEP_V1.4 Julio 2026.xlsm` extrae **13 módulos**
(`ThisWorkbook`, `Módulo1` con `Sub PROTEGER()`, `Hoja1..Hoja10`, `Hoja21`).
Comparado contra `data/local/2026-7 REMASEP_V1.4.xlsm` (otro `Save` real):
`vbaProject.bin` difiere en binario (40 960 vs 45 056 bytes; `__SRP_*` pasa de
4 a 10 streams) pero **`code_changed_modules == ()`** — sólo
`attributes_changed_modules == ('Hoja21',)` por el reordenamiento de tres
`Attribute VB_Control`. Estado → `VBA_INTEGRITY_PASS` + warnings.

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

1. **Test de integración COM** en **Windows** con Python + `pywin32` + Microsoft
   Excel Desktop. Se pasa la celda de prueba por variables de entorno (no se
   elige automáticamente):

   ```powershell
   $env:REMASEP_TEST_TEMPLATE = "C:\ruta\a\REMASEP_V1.4 Julio 2026.xlsm"
   $env:REMASEP_TEST_SHEET    = "REMASEP_OD"
   $env:REMASEP_TEST_CELL     = "K17"
   pytest -m excel tests/test_excel_com_integration.py -vv -s
   ```

   Celda propuesta **`REMASEP_OD!K17`**: es `WRITE_READY` en el manifiesto
   (`instruction_id wi:00a4bd2ee2668dd9e28c`, source `LEGACY::REMASEP_OD::L8`),
   hoja protegida con la celda **desbloqueada** (`Locked = False`), sin fórmula,
   fuera de todo merge, valor actual `0` — un sentinel `424242` es inequívoco.
   (SECCION A / A.1 / "PRIMERAS CONSULTAS (AÑO) DE ODONTOLOGIA", fácil de
   localizar a mano.) Alternativas equivalentes: `REMASEP_OD!AB69`
   (`wi:00170a112286f2273a2a`) o `REMASEP 01!AK118` (`wi:00aab54517c24d1e515f`).

2. Ejecutar `scripts/generate_remasep.py --mode production` en Windows (desde
   Sprint 3.8 es el default): sólo necesita el export `detalle_citas` y
   `REMASEP_V1.4 Julio 2026.xlsm`; el conocimiento interno viene de
   `config/runtime_2026/` (ver [`RUNTIME_DECOUPLING.md`](RUNTIME_DECOUPLING.md)).
   `--mode diagnostic` (que sí abre `GENERACION DATOS REMASEP.xlsx`) queda como
   camino de comparación dev-only.
3. Revisar `artifacts/excel_writer/` (integridad de fórmulas/VBA, verificación de
   las 1122 celdas, CONTROL).
4. Confirmar funcionalmente el filtro por ESTADO para habilitar el modo
   productivo real (fuera del alcance de este sprint).
