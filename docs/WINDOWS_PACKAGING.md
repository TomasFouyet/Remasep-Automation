# Empaquetado Windows (Sprint 3.11 — piloto)

Construye el primer ejecutable Windows usable de REMASEP Automation con
**PyInstaller, modo ONEDIR** (carpeta, no un único `.exe`). Sin privilegios de
administrador. Sin instalador MSI/NSIS, sin firma digital, sin auto-update
todavía — eso queda para un sprint posterior.

```
dist/REMASEP/
  REMASEP.exe
  _internal/
    config/runtime_2026/...
    config/excel_writer_2026/...
    PySide6/...    (Qt: sólo QtCore/QtGui/QtWidgets)
    ...            (Python + demás dependencias)
```

Este documento es **técnico** (para quien construye el `.exe`). La guía para
la persona usuaria final es [`WINDOWS_PILOT.md`](WINDOWS_PILOT.md) — no
menciona Python, pip, venv ni terminal.

---

## Fase 0 — Auditoría de paths (antes de tocar código)

Clasificación de todo path relativo a `cwd` / `__file__` / literal encontrado
en `src/` y `scripts/`:

| Categoría | Qué | Dónde | Estado |
|---|---|---|---|
| **A — INTERNAL_READ_ONLY** (viaja con la app) | `config/runtime_2026/` | `services/runtime_assets.py::resolve_runtime_root()` | Ya era MEIPASS-aware (Sprint 3.8). Sin cambios. |
| | `config/excel_writer_2026/` | `services/generation_service.py::_DEFAULT_CONFIG_DIR` | Era `Path("config/excel_writer_2026")`, relativo al *cwd*. **Corregido** con inyección desde `app_paths` (ver Fase 1/2). |
| | `config/legacy_current_logic_2026/`, `writable_target_mapping_2026/`, `semantic_mapping_2026/`, `semantic_validation_2026/`, `official_template_alignment_2026/` | 5 módulos con `Path(__file__).resolve().parents[3]/...` | Confirmado por grafo de imports: **ninguno se ejecuta** desde el flujo de producción/UI (sólo scripts dev / `MedinetAnalysisService`, nunca instanciado en runtime). No viajan con la app. |
| **B — USER_INPUT** | Medinet, plantilla REMASEP | `QFileDialog` (`FileSelector`) | Ya independientes de *cwd*. |
| **C — USER_OUTPUT** | REMASEP `.xlsm`, PDF | `QFileDialog` (`GenerateScreen`/`DashboardScreen`, Sprint 3.10) | Ya independientes de *cwd*. El *workspace* temporal de Excel COM vive junto al output elegido (`output_path.parent/.remasep-tmp/<run_id>/`) — correcto, sin cambios. |
| **D — APP_WRITABLE_DATA** | `artifacts/excel_writer/` (diagnóstico) | `generation_service.py::_DEFAULT_ARTIFACTS_DIR` | Relativo al *cwd*. **Corregido**: la UI inyecta `app_paths.diagnostics_dir()`. El **CLI no cambia** (sigue con su default relativo). |
| | "Último informe generado" | `ui/screens/home.py::_OUTPUTS_DIR` | Relativo al *cwd*; en el `.exe` casi nunca encontrará nada porque el output ahora se elige con "Guardar como" (normalmente Escritorio). Degradación silenciosa (la card sólo se oculta) — **no se amplía el alcance** agregando tracking nuevo de "último informe" en este sprint. |
| | Sugerencia inicial de "Guardar como" | `ui/screens/generate.py::_prompt_output_path` | Ya tenía *fallback* a `Path.home()` si `outputs/` no es un directorio → ya era seguro para el `.exe`. |
| | Logging | No había ningún `Handler` configurado | **Agregado**: `remasep/logging_setup.py` (Fase 8). |

### Bugs reales encontrados *ejecutando* (no sólo leyendo código)

La auditoría estática no basta: tres problemas sólo aparecieron al **ejecutar**
código con un `cwd`/entorno distinto al del repo, que es exactamente la
condición del `.exe` empaquetado:

1. **`control_map` en `policy.yaml` traía un path relativo baked-in**
   (`control_map: config/excel_writer_2026/control_map.yaml`), resuelto
   contra el *cwd* del proceso en vez de contra `config_dir`. Con
   `config_dir` ya resuelto vía `app_paths` (una ruta absoluta fuera del
   *cwd*), `GenerationService._control_map` fallaba con
   `GenerationServiceError: falta el mapa de CONTROL`. **Fix**: el valor en
   `policy.yaml` pasa a ser sólo el nombre de archivo (`control_map.yaml`) y
   `_control_map` lo resuelve contra `self._config_dir` (ruta absoluta se
   respeta tal cual). Ver `test_generation_service_packaging.py`.
2. **`resource_root()` chocaba con un subpaquete Python no relacionado**
   llamado igual: existe `src/remasep/config/` (con un único `loader.py`,
   código huérfano de un sprint muy temprano, `remasep.domain`/`rules` — vivo
   pero sin relación con el config de datos). El primer diseño de
   `resource_root()` ("subir hasta encontrar una carpeta `config/`") se
   detenía ahí en vez de seguir hasta la raíz real del repo. **Fix**: se
   exige que `config/` tenga **ambos** árboles reales
   (`runtime_2026`/`excel_writer_2026`) para contar como raíz. Ver
   `test_resource_root_ignores_unrelated_config_named_directories`.
3. **Un test propio del proyecto** (`test_no_process_killing_in_codebase`,
   invariante de seguridad desde el Sprint 3.7B: nunca `subprocess`/`psutil`
   en `src/`/`scripts/`) frenó un primer intento de que
   `scripts/build_smoke_test.py` lanzara `REMASEP.exe --smoke-mode`
   automáticamente. Se respetó la invariante existente: el smoke test del
   dist es **sólo de archivos** (nunca lanza procesos); arrancar el `.exe` de
   verdad se documenta como paso manual (Fase 10/11).

Estos tres hallazgos son la prueba de por qué la Fase 9 pide tests que
*simulen* `sys._MEIPASS` / *cwd* distinto en vez de sólo leer código.

---

## Arquitectura: resources vs writable data

`src/remasep/app_paths.py` — abstracción pequeña y central, dos categorías
que **nunca se mezclan**:

- **`resource_root()` / `find_config_dir(name)` / `runtime_config_dir()` /
  `excel_writer_config_dir()`** — INTERNAL_READ_ONLY. En desarrollo: sube
  desde `app_paths.py` hasta encontrar `config/` con ambos árboles reales. En
  un build PyInstaller: `sys._MEIPASS` (la carpeta `_internal/` en ONEDIR).
  `runtime_config_dir()` delega en el resolver ya validado de
  `runtime_assets.py` (Sprint 3.8); `excel_writer_config_dir()` es el mismo
  patrón, nuevo, para `config/excel_writer_2026/`.
- **`user_data_dir()` / `logs_dir()` / `diagnostics_dir()`** — APP_WRITABLE_DATA.
  `%LOCALAPPDATA%\REMASEP Automation\` en Windows (`~/.local/share/REMASEP
  Automation` fuera de Windows, para desarrollo). **Nunca** junto al `.exe`,
  **nunca** dentro del bundle. `REMASEP_APP_DATA_DIR` permite redirigirlo
  (tests, o aislar un piloto).

Los **outputs que elige la persona usuaria** (REMASEP `.xlsm`, PDF) son
`USER_OUTPUT`: viven donde ella decida con el diálogo — **nunca** se copian a
`user_data_dir()`, y el archivo Medinet completo **nunca** se cachea ahí
tampoco.

---

## Cambios realizados

| Archivo | Cambio |
|---|---|
| `src/remasep/app_paths.py` (nuevo) | `resource_root`, `find_config_dir`, `runtime_config_dir`, `excel_writer_config_dir`, `user_data_dir`, `logs_dir`, `diagnostics_dir`, `is_frozen`. |
| `src/remasep/logging_setup.py` (nuevo) | `configure_logging()` — `RotatingFileHandler` (1 MB × 2 backups) bajo `logs_dir()`; `_guard_frozen_streams()` evita `AttributeError` si `sys.stdout`/`stderr` son `None` (build `--windowed`). |
| `src/remasep/main.py` | `--smoke-mode`: verifica assets internos + política del writer + `user_data_dir` sin abrir ventana ni tocar Excel COM (Fase 10). |
| `src/remasep/ui/main_window.py::run_app()` | Configura logging al arrancar; si algo inesperado falla, deja evidencia en el log y muestra un aviso humano (nunca consola/traceback) antes de salir. |
| `src/remasep/ui/workers.py::run_generation()` | El `GenerationService` por defecto ahora se construye con `config_dir=app_paths.excel_writer_config_dir()` y `artifacts_dir=app_paths.diagnostics_dir()` — inyectado desde la capa UI, **sin tocar los defaults de `GenerationService`** (el CLI no cambia). |
| `src/remasep/services/generation_service.py::_control_map()` | Resuelve `control_map` (de `policy.yaml`) contra `self._config_dir`, no contra el *cwd*. |
| `config/excel_writer_2026/policy.yaml` | `control_map: config/excel_writer_2026/control_map.yaml` → `control_map: control_map.yaml` (relativo a su propio directorio). |
| `pyproject.toml` | Extra `build = ["pyinstaller>=6,<7"]` — **no** es dependencia de runtime del usuario final. |
| `.gitignore` | `+ .venv-win/`. |

**Sin cambios en**: Excel Writer, no-overwrite, integridad VBA, *workspace*
temporal, `PendingWrite`/mappings, reglas REMASEP, `MetricValueProducer`,
`runtime_2026` (regla ESTADO confirmada intacta), CLI (`scripts/generate_remasep.py`
sigue igual — sigue funcionando desde el repo).

---

## Fase 3 — qué incluye el build

**Incluido** (`packaging/_spec_common.py::common_paths()` → `datas`):

- `config/runtime_2026/` completo (manifiesto de escritura, catálogo de
  fórmulas, contrato de detalle, política del cero, filtro ESTADO
  confirmado, fingerprint de plantilla).
- `config/excel_writer_2026/` completo (política del writer, mapa de
  CONTROL).

**Excluido explícitamente** (nunca en `dist/`):

- `data/local/` (exports reales de Medinet, `GENERACION DATOS REMASEP.xlsx`,
  plantillas reales del cliente).
- `artifacts/` de desarrollo (reverse-engineering de sprints anteriores).
- `outputs/`, `tests/`, `.git/`, `.venv/`, `.venv-win/`.
- El resto de `config/` no usado en runtime (`semantic_*_2026`,
  `writable_target_mapping_2026`, `legacy_current_logic_2026`,
  `official_template_alignment_2026`, `templates/`) — sólo para scripts dev.

La app final **exige** a la persona usuaria elegir su propio archivo Medinet
y su propia plantilla REMASEP; los assets internos vienen dentro del `.exe`.

`scripts/build_smoke_test.py` verifica automáticamente esta lista (Fase 10).

---

## Fase 4/5/6/7 — PyInstaller, Qt, Excel COM, identidad

- **`packaging/remasep.spec`** — variante windowed (sin consola), la que se
  entrega. **`packaging/remasep_debug.spec`** — variante con consola, mismo
  contenido, sólo para diagnosticar arranque (`dist/REMASEP-debug/`). Ambas
  comparten config en `packaging/_spec_common.py` (DRY, un solo lugar para
  `datas`/`hiddenimports`/`excludes`).
- **`packaging/run_remasep.py`** — *bootstrap* de entrada, deliberadamente
  fuera del paquete `remasep` (evita que PyInstaller lo importe dos veces
  bajo dos nombres). Sólo llama a `remasep.main.main()`.
- **Qt (Fase 5)**: `grep` sobre `src/` confirma que la app sólo importa
  `QtCore`, `QtGui`, `QtWidgets` (el PDF usa `QPdfWriter`, nativo de
  `QtGui` — no hace falta el módulo `QtPdf`). El `.spec` excluye
  explícitamente `QtWebEngine*`/`QtQml`/`QtQuick*`/`Qt3D*`/`QtMultimedia*`/
  `QtBluetooth`/`QtSensors`/`QtWebChannel`/`QtPositioning*`/`QtNfc`/
  `QtDesigner`/`QtHelp`/`QtRemoteObjects`/`QtSql`/`QtTest` — todo evidenciado
  como no usado, nada "por si acaso". El plugin de plataforma Windows
  (`qwindows.dll`) lo resuelve el hook oficial de PyInstaller para PySide6
  (no hace falta configurarlo a mano).
- **Excel COM / pywin32 (Fase 6)**: `win32com`/`pythoncom`/`pywintypes` se
  importan de forma **perezosa** dentro de funciones
  (`adapters/excel_com.py`, `services/excel_capability.py`) — el análisis
  estático de PyInstaller no los ve solos, así que van como `hiddenimports`
  explícitos (con la razón documentada en el `.spec`); `win32timezone` es un
  requisito de runtime de pywin32+PyInstaller ampliamente conocido. **El
  ciclo de vida COM no se toca**: sigue creando sólo su instancia aislada de
  Excel (`DispatchEx`), soltando referencias antes de `CoUninitialize`, sin
  `taskkill`/`Stop-Process`, sin afectar el Excel del usuario (Sprint 3.7B,
  intacto).
- **Identidad (Fase 7)**: `packaging/version_info.txt` — `ProductName`/
  `FileDescription` = "REMASEP Automation", `CompanyName` = "Fundación Gantz"
  (ya usado en el título de ventana desde el Sprint 3.10 — no es información
  inventada), versión `0.1.0` (la del proyecto). **Sin `LegalCopyright`**: no
  se inventa información legal. **Sin icono aprobado todavía** —
  `packaging/assets/README.md` documenta dónde poner `remasep.ico` cuando
  exista uno; mientras tanto el build usa el icono por defecto de
  PyInstaller (no bloquea el sprint).

---

## Build reproducible

```powershell
# una sola vez, para preparar el entorno de build
python -m venv .venv-win
.venv-win\Scripts\pip install -e .[build]

# build de piloto (windowed)
.\scripts\build_windows.ps1

# variante debug/consola, para diagnosticar
.\scripts\build_windows.ps1 -Debug
```

`scripts/build_windows.ps1`: valida que corre en Windows, valida
`.venv-win\Scripts\python.exe` + PyInstaller instalado, limpia **sólo**
`dist/`/`build/` de este repo (nunca otra ruta), ejecuta PyInstaller con el
`.spec` versionado, confirma que `dist\REMASEP\REMASEP.exe` (o
`REMASEP-debug\REMASEP-debug.exe`) existe, imprime la ruta final. **Nunca**
`taskkill`, **nunca** toca Excel abierto, **nunca** borra carpetas fuera de
`dist/`/`build/` del propio repo.

Después de construir:

```powershell
python scripts\build_smoke_test.py --dist dist\REMASEP
```

Verifica estructuralmente (sin lanzar ningún proceso): `REMASEP.exe`
presente, `config/runtime_2026` y `config/excel_writer_2026` presentes con
sus archivos clave, y que **no** aparece `data/local`, `GENERACION DATOS
REMASEP`, ningún export de Medinet, ni carpetas `artifacts/`/`tests/`/`.git`/
`.venv*`/`outputs` en ningún punto del árbol. Para confirmar que el `.exe`
arranca de verdad (assets resueltos, sin ventana, sin Excel COM), a mano:

```
dist\REMASEP\REMASEP-debug.exe --smoke-mode
```

---

## Pruebas de paths (Fase 9)

`tests/test_app_paths.py`, `tests/test_runtime_assets.py`,
`tests/test_generation_service_packaging.py`, `tests/test_logging_setup.py`,
`tests/test_build_smoke_test.py` — todas corren en Linux/WSL simulando lo
que en Windows resuelve `sys._MEIPASS` (no hace falta un build real para
probar la lógica de resolución):

- `resource_root()` independiente del *cwd*; usa `sys._MEIPASS` simulado
  cuando está "congelado"; ignora directorios `config/` no relacionados
  (regresión del hallazgo #2 de más arriba).
- `runtime_config_dir()` / `excel_writer_config_dir()` encuentran los assets
  reales, con rutas que incluyen **espacios y caracteres Unicode** (ej.
  `"Carpeta con espacios y ñ"`), simulando una instalación en una ruta
  realista de Windows (Escritorio, OneDrive con tildes).
- `user_data_dir()`/`logs_dir()`/`diagnostics_dir()` son escribibles, nunca
  caen dentro del bundle simulado, soportan espacios/Unicode
  (`"María José Ñúñez"`).
- La inyección UI→`GenerationService` (Fase 2) usa exactamente
  `excel_writer_config_dir()`/`diagnostics_dir()`, nunca el *cwd* (regresión
  del hallazgo #1).
- El logging nunca lanza si el disco no es escribible, rota por tamaño, nunca
  escribe junto al bundle.
- El *smoke check* estructural detecta correctamente un build correcto vs.
  uno con fugas de desarrollo (`data/local`, el workbook legacy,
  `artifacts/`, etc.), en árboles sintéticos.

Ningún test cambia el *cwd* global del proceso: todos usan
`monkeypatch.chdir`/`REMASEP_APP_DATA_DIR` (auto-revertidos al terminar).

---

## Lo que este sprint **no** puede probar desde Linux/WSL

No hay Windows disponible en este entorno: PyInstaller no cross-compila, así
que **no se ejecutó un build real** ni la prueba manual de las Fases 11-13.
Lo que sí se hizo:

- Toda la lógica de resolución de paths, probada simulando `sys._MEIPASS` /
  *cwd* distinto / rutas con espacios y Unicode (arriba).
- El `.spec`, el script de build y el smoke test están completos y listos
  para ejecutarse en Windows sin cambios.
- El flujo UI completo (Medinet → período → plantilla → análisis → dashboard
  → PDF → Generar → "Guardar como") sigue verificado *desde source* (suite
  completa, 884 passed / 2 skipped — sólo COM Windows-only).

**Pendiente, a hacer en Windows** (ver [`WINDOWS_PILOT.md`](WINDOWS_PILOT.md)
para los pasos de la persona usuaria):

1. `.\scripts\build_windows.ps1` sobre una máquina Windows con Excel Desktop.
2. `python scripts\build_smoke_test.py --dist dist\REMASEP`.
3. Copiar `dist\REMASEP\` a una carpeta **fuera** del repo (ej.
   `Escritorio\REMASEP_PILOT\`) y probar el flujo completo desde ahí — sin
   `.venv-win` activado, sin relación con el repo.
4. Repetir una generación con el mismo nombre de salida: debe pedir otro
   nombre, nunca sobrescribir.
5. Revisar `%LOCALAPPDATA%\REMASEP Automation\logs\` y `\diagnostics\` tras
   la prueba: no debe haber nombre/RUT/fecha de nacimiento/teléfono/email de
   paciente (Fase 13).
