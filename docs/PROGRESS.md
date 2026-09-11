# Estado del proyecto

Vista de estado post Sprint 3.5 (capa semántica: inventario + mapeo de
dimensiones + readiness técnica; delta del REMASEP final + provenance de fuente;
alineación semántica con la plantilla oficial). Para el detalle por sprint ver
[`docs/sprints/`](sprints/README.md).

## Fases

| Fase | Contenido | Estado |
| --- | --- | --- |
| **1 — Reverse Engineering** | Inventario del workbook, dependencias, lógica legacy, plantilla oficial, UI shell | **COMPLETE** |
| **2 — Medinet & Legacy Equivalence** | Ingesta Medinet real, `AC:AL`, agregaciones directas, cierre por dependencias, `SUM`/`IF` downstream | **COMPLETE** |
| **3 — Semantic Metrics** | Inventario semántico (3.1 ✅), mapeo de dimensiones sexo/edad/alcance/código (3.2 ✅), readiness técnica + cola de revisión (3.3 ✅), delta del REMASEP final + provenance de fuente (3.4 ✅), alineación semántica con la plantilla oficial (3.5 ✅), writable target mapping (3.6 ✅), metric value producer + equivalencia de valores (3.7A ✅) | **IN PROGRESS** |
| **4 — Additional Sources / Complete Dataset** | Adaptador de egresos, cálculo de recursos, dataset golden | PENDING |
| **5 — Official Template Mapping** | Mapping semántico generador/Medinet → plantilla MINSAL (alineación de candidatos hecha en 3.5; escritura Excel pendiente) | PENDING |
| **6 — Excel Generation / CONTROL** | Escritura vía Excel COM (Windows), recálculo, verificación `CONTROL` | **IN PROGRESS** (3.7B: writer + verificación implementados; falta corrida real en Windows) |
| **7 — Pilot / Packaging** | Piloto manual + automático en paralelo, empaquetado `.exe` | PENDING |

## Completado

- Inventario del workbook generador (Sprint 1.1)
- Análisis de dependencias de fórmulas (Sprint 1.2)
- Inventario de la lógica legacy actual (Sprint 1.3)
- Inventario y alineación de la plantilla oficial (Sprint 1.4)
- UI shell navegable con datos mock (Sprint 1 UI)
- Ingesta real de Medinet + clasificación legacy (Sprint 2.1)
- Detección de filas estructuralmente vacías (Sprint 2.1)
- Equivalencia `AC:AL` (Sprint 2.2)
- Equivalencia de agregaciones directas `COUNTIF`/`COUNTIFS` (Sprint 2.3)
- Cierre por dependencias — agregaciones derivadas (Sprint 2.4)
- Equivalencia downstream `SUM` / `IF` (Sprint 2.5)
- **Cierre completo de la lógica de fórmulas legacy derivada de Medinet:
  2043 / 2043 celdas evaluables**
- **Inventario semántico preliminar de las 1771 métricas (Sprint 3.1)**:
  rótulos de sección / fila / columna (raw + normalizado), `semantic_signature`
  candidata, evidencia de fórmula (cotas de edad, criterios de texto),
  consistencia rótulo ↔ fórmula, `source = "MEDINET"`. Ver
  [`docs/SEMANTIC_METRIC_INVENTORY.md`](SEMANTIC_METRIC_INVENTORY.md).

## En curso — Sprint 3

- **3.1 — Semantic Metric Inventory: ✅ COMPLETE.** 1771 métricas candidatas,
  1741 `COMPLETE` / 30 `PARTIAL` / 0 `AMBIGUOUS` / 0 `NO_CONTEXT`, 0 firmas
  duplicadas. Ver
  [`docs/SEMANTIC_METRIC_INVENTORY.md`](SEMANTIC_METRIC_INVENTORY.md).
- **3.2 — Semantic Metric Mapping: ✅ COMPLETE.** 1771 `SemanticMetric` con
  dimensiones explícitas (sexo / edad / alcance de agregación / código de
  procedimiento), cada una con `value` + `status` + `evidence`. Resultado real:
  1414 `mapping_status=CONFIRMED` · 352 `PARTIAL` · 5 `CONFLICT` (reales, del
  workbook: columna `Mujeres` con fórmula `"*Hombre*"` en `REMASEP 01`).
  597 códigos de procedimiento explícitos (19 distintos). Vocabulario versionado
  en `config/semantic_mapping_2026/`
  (`status: preliminary_pending_functional_validation`). Ver
  [`docs/SEMANTIC_METRIC_MAPPING.md`](SEMANTIC_METRIC_MAPPING.md).
- **3.3 — Semantic Mapping Validation & Readiness: ✅ COMPLETE.** Capa técnica de
  *readiness* sobre `SemanticMetric` (no cambia las inferencias de 3.2). Policy
  versionada en `config/semantic_validation_2026/`
  (`status: preliminary_pending_functional_validation`). Resultado real sobre las
  1771 métricas: **1401 `AUTO_READY` · 50 `REVIEW_REQUIRED` · 5
  `BLOCKED_CONFLICT` · 315 `NOT_APPLICABLE`** (roll-ups del workbook).
  Issues: 5 `BLOCKING` / 69 `REVIEW` / 1360 `INFO`, 13 clusters. Los 5 conflictos
  reales (`REMASEP 01!AB84/AB86/AB87/AB89/AB90`) quedan `BLOCKED_CONFLICT` con
  `recommended_action = HUMAN_REVIEW` — **no se corrigen**, sólo se clasifican.
  `overrides.yaml` vacío (arquitectura, sin resolución). `AUTO_READY` ≠ validado
  MINSAL. Ver
  [`docs/SEMANTIC_MAPPING_VALIDATION.md`](SEMANTIC_MAPPING_VALIDATION.md).
- **3.4 — Final REMASEP Delta & Source Provenance: ✅ COMPLETE.** Comparación
  reproducible y sólo-lectura del `2026-7 REMASEP_V1.4.xlsm` (incompleto,
  `INCOMPLETE_PRE_SURGERY`) contra `REMASEP_V1.4 Julio 2026.xlsm` (terminado
  según el cliente, `FINAL_PER_CLIENT` · `approval_status = UNCONFIRMED`).
  Resultado real: **0 cambios de expresión de fórmula, 113 celdas de valor
  cambiadas** (52 input directo + 61 propagación de fórmula + 0 expresión
  reescrita; por hoja REMASEP 01 = 58 · REMASEP B1 = 41 · B2 ANEXO = 10 ·
  CONTROL = 4). `change_kind` se decide con la fórmula **antes y después**
  (`DIRECT_INPUT_CHANGE` / `FORMULA_RESULT_CHANGE` / `FORMULA_EXPRESSION_CHANGE`),
  no sólo con la de la versión final. Provenance por celda
  (`EGRESOS` 51 · `SURGICAL_TABLE` 30 · `UNKNOWN_PENDING` 17 · `RESOURCE_
  CALCULATION` 10 · `CONTROL_METADATA` 4 · `MIXED_DERIVED` 1). `EGRESOS`
  (edad/sexo → B1, códigos Qx → B2 ANEXO) confirmado por el cliente; Sección D de
  REMASEP 01 inferida por encabezado (recursos vs. tabla quirúrgica, revisión
  humana); Sección E (suspensiones) queda `UNKNOWN_PENDING`. `CONTROL = 0` →
  `PASS_INTERNAL_VALIDATION` (≠ aprobación MINSAL). 32-vs-31 registrado como
  `OPEN_FUNCTIONAL_QUESTION`. `candidate_golden_status =
  CANDIDATE_PENDING_SOURCE_COMPLETENESS`. **No** se automatiza ninguna fuente
  nueva. Ver
  [`docs/FINAL_REMASEP_DELTA_PROVENANCE.md`](FINAL_REMASEP_DELTA_PROVENANCE.md).
- **3.5 — Official Template Semantic Alignment: ✅ COMPLETE.** Empareja cada
  `SemanticMetric` MEDINET **elegible** (`readiness = AUTO_READY`) del generador
  legacy con la(s) celda(s) de la plantilla oficial (`REMASEP_V1.4 Julio
  2026.xlsm`), por **evidencia por dimensión** — nunca por coordenada (los dos
  layouts difieren; 0 de 1122 matches EXACT tienen coordenada igual). Reglas
  versionadas en `config/official_template_alignment_2026/`
  (`status: preliminary_pending_functional_validation`). Resultado real: 1401
  elegibles / 370 excluidos (5 `BLOCKED_CONFLICT` — **no** entran a AUTO mapping —
  · 315 `ROLLUP_NOT_INPUT` · 50 `REVIEW_REQUIRED`); **1122
  `EXACT_SEMANTIC_MATCH` · 245 `STRONG_MATCH` · 34 `AMBIGUOUS` · 0 `NO_MATCH` · 0
  `CONFLICT`** (REMASEP_OD 952/170/34, REMASEP_01 165/75, B2_ANEXO 5 por código
  de prestación). Regiones NO-MEDINET (`EGRESOS` / `RESOURCE_CALCULATION` /
  `SURGICAL_TABLE` / `UNKNOWN`, Sprint 3.4) se excluyen del alignment; `REMASEP
  B1` entero es `EGRESOS` (sin source MEDINET). **Revisión de cierre**: se separan
  `target_kind` (qué es la celda) · `target_alignment_role` (`INPUT_TARGET` /
  `DERIVED_TARGET` / `STRUCTURAL` / …) · `target_source_expectation` (default
  `UNKNOWN`, no `MEDINET`). `DIRECT_INPUT_TARGET` exige **celda desbloqueada en
  hoja protegida**, no sólo "sin fórmula" (rótulos y códigos de prestación
  tampoco tienen fórmula): 2 906 celdas dejaron de contarse como input
  (`INPUT_TARGET` 11285 → **8379**; `STRUCTURAL` → **6985**), **sin cambiar el
  matching**. De los 8379 input targets: **5227 esperados de MEDINET, 1367
  alineados, 3860 genuinamente sin source**; fórmulas (1174) y estructurales
  (6985) **no** son "targets MEDINET faltantes".
  `approval_status = UNCONFIRMED`,
  `candidate_golden_status = CANDIDATE_PENDING_SOURCE_COMPLETENESS`. **No**
  escribe Excel, **no** COM. Ver
  [`docs/OFFICIAL_TEMPLATE_SEMANTIC_ALIGNMENT.md`](OFFICIAL_TEMPLATE_SEMANTIC_ALIGNMENT.md).
- **Hotfix producción — contrato de input Medinet & alcance de período: ✅
  COMPLETE.** Formalizado: el input real de producción es el export directo
  **"Detalle de citas"** de Medinet; `GENERACION DATOS REMASEP.xlsx` es sólo
  referencia legacy (la app no depende de él). Bug corregido: la clasificación
  legacy / edades corrían sobre **todos** los registros válidos; ahora sólo
  sobre `processing_scope_records` = válidos ∩ mes/año. Los registros de otros
  períodos no se descartan del archivo, sólo del cálculo; mensaje de UI
  reformulado. Reconciliación privacy-safe (multiset) del archivo real:
  `raw 11 501 · Julio 2114 · fuera 9 387`; los **1 364** activos del detalle
  legacy son **subconjunto multiset exacto** (huella sin ESTADO) de los 2 114; la
  diferencia de **750** = ESTADO. Hipótesis con evidencia de recuento exacto
  (legacy conserva `{Atendido, Atención Pausada, En Sala de Espera, En
  Atención}` = 1 364), marcada
  `CURRENT_LEGACY_BEHAVIOR_PENDING_FUNCTIONAL_CONFIRMATION`, **no** implementada
  como regla. Ver [`docs/MEDINET_INPUT_CONTRACT.md`](MEDINET_INPUT_CONTRACT.md).
- **3.6 — Writable Target Mapping: ✅ COMPLETE.** Capa por encima del semantic
  alignment (3.5): convierte los alignments seguros en un manifiesto de
  **instrucciones de escritura** versionable. **No escribe Excel.** `write_status`
  ∈ `WRITE_READY` / `WRITE_REVIEW_REQUIRED` / `WRITE_BLOCKED` / `NOT_WRITABLE` —
  capa distinta de `readiness` y `match_status`. Policy versionada en
  `config/writable_target_mapping_2026/`
  (`status: preliminary_pending_functional_validation`). Resultado real (1771
  métricas): **1122 `WRITE_READY`** (todos `EXACT`; manifiesto = 1122
  instrucciones) · **245 `WRITE_REVIEW_REQUIRED`** (todos `STRONG`: la policy no
  baja umbrales — `ROW_PATH PARTIAL` 170 · `SEX MISSING` 65 · `AGE MISSING` 10) ·
  **39 `WRITE_BLOCKED`** (34 `AMBIGUOUS` + 5 `SOURCE_BLOCKED_CONFLICT`) · **365
  `NOT_WRITABLE`** (315 roll-ups calculados por la plantilla + 50 source
  `REVIEW_REQUIRED`). WRITE_READY por forma: REMASEP_OD 952 · REMASEP_01 165 ·
  B2_ANEXO 5. 0 colisiones. `instruction_id` = hash de firma semántica + template
  fingerprint (no la coordenada). `structural_template_fingerprint` distinto del
  `file_sha256` (un re-guardado no invalida el manifiesto). `zero_write_policy =
  UNRESOLVED` y `estado_filter_status = PENDING_FUNCTIONAL_CONFIRMATION`
  (precondiciones del 3.7). Ver
  [`docs/WRITABLE_TARGET_MAPPING.md`](WRITABLE_TARGET_MAPPING.md).
- **3.7A — Medinet Metric Value Producer & Reference Value Equivalence: ✅
  COMPLETE.** Productor real de `MetricValue` para MEDINET, reutilizando el motor
  legacy (`legacy_aggregation` + `legacy_transform` + cierre de dependencias) —
  una sola implementación, sin reimplementar fórmulas. **No escribe Excel.**
  Modos `PRODUCTION_PERIOD_SCOPE` (registros = `processing_scope_records`,
  Julio 2026 = 2114, **sin** filtro por ESTADO, `input_scope = PERIOD_ONLY`) y
  `LEGACY_EQUIVALENCE_DIAGNOSTIC` (opt-in explícito, 1364 = detalle legacy, sólo
  para el estudio de equivalencia). `join_metric_values_with_manifest(...) →
  PendingWrite` + `check_write_completeness(...)` (sin *last-one-wins*).
  Resultado real (Julio 2026): **1122/1122 `MetricValue`**, 0 conflictos de tipo,
  completitud missing/duplicate/orphan = 0. Equivalencia contra la plantilla
  final oficial: **1122/1122 MATCH vs. el valor cacheado por el Excel legacy**;
  980 MATCH / 142 MISMATCH vs. el valor reevaluado, y los 142 son la caché
  obsoleta de `AF`/edad (`docs/TECHNICAL_OVERVIEW.md §8`), no un error del
  productor. Diagnóstico: **1122/1122 `MetricValue` idénticos al valor legacy
  reevaluado** (`EXACT_EQUIVALENCE`). Semántica del cero: 935 casos `source == 0`,
  917 target 0, **0 target blanco** → `zero_write_policy = WRITE_ZERO`
  (evidence-derived, `config/metric_value_producer_2026/zero_write_policy.yaml`).
  `estado_filter_status = PENDING_FUNCTIONAL_CONFIRMATION` (el modo productivo no
  lo aplica; el diagnóstico no lo promueve a regla). Ver
  [`docs/METRIC_VALUE_PRODUCER.md`](METRIC_VALUE_PRODUCER.md).
- **3.7B — Safe Excel COM Writer & Generated Workbook Verification: ✅ COMPLETE
  (código; falta la corrida real en Windows).** Escribe los `PendingWrite` del
  3.7A en una **copia** del REMASEP oficial con Microsoft Excel Desktop (COM,
  sólo Windows). Núcleo independiente de plataforma
  (`services/excel_writer.py`: `WorkbookWriter` Protocol, preflight, snapshot,
  verificación) + `adapters/excel_com.py` (`ExcelComWorkbookWriter`,
  `win32com` perezoso) + `FakeWorkbookWriter` (tests) + `ExcelCapability`
  (en WSL `can_generate = False`, mensaje graceful, sin traceback) +
  `GenerationService` (API para la UI, **no** conectada) + CLI
  `scripts/generate_remasep.py`. Copy-first + **workspace temporal único por
  corrida** (`outputs/.remasep-tmp/<run_id>/working.xlsm`) + promoción atómica
  (`os.replace` con reintento acotado) a `outputs/REMASEP_2026_07_DRAFT.xlsm`;
  cleanup conservador best-effort (`CLEANUP_PENDING` si Windows retiene un handle
  — **sin `taskkill`, sin loop, sin borrado forzado**). SHA256 de la
  plantilla verificado sin cambios. Preflight (duplicados, tipo/negativo,
  faltantes, zero_write_policy, salida fuera de `data/`, no sobrescribe).
  Compatibilidad por `structural_template_fingerprint` leído del manifiesto
  (`stf:dc624775927d4d4d`). Defense-in-depth por celda: nunca sobrescribe una
  fórmula (`TARGET_FORMULA_CONFLICT`). `Value2`, 0 explícito. `CalculateFullRebuild`.
  Verificación posterior: integridad de fórmulas (0 `FORMULA_EXPRESSION_CHANGE`),
  **integridad SEMÁNTICA de VBA** (`services/vba_integrity.py` + `olefile`:
  descompresión MS-OVBA §2.4.1 y comparación de código por módulo — falla si
  desaparece el VBA / cambia el set de módulos / cambia el código; el cambio
  sólo del binario `vbaProject.bin` por un `Save` de Excel es warning, no fallo
  —sólo si la comparación semántica se pudo ejecutar—; **fail-closed** si el VBA
  existe pero no se pudo verificar el código, `reason =
  VBA_SEMANTIC_CHECK_UNAVAILABLE`), 1122 celdas destino,
  fingerprint. `ControlResult`
  (`PASS/FAIL_INTERNAL_VALIDATION` / `CONTROL_UNAVAILABLE`) con mapa versionado
  `config/excel_writer_2026/control_map.yaml` — **`writer_integrity_status` es
  independiente de `control_status`** (faltan EGRESOS / recursos / tabla
  quirúrgica / metadatos). Modos `DIAGNOSTIC_REFERENCE` (scope 1364) /
  `PRODUCTION` (2114); archivo `NOT_FOR_SUBMISSION`, nunca `FINAL_READY` mientras
  `estado_filter_status = PENDING_FUNCTIONAL_CONFIRMATION`.
  `historical_reference_cache_status = KNOWN_STALE_FOR_142_WRITE_READY_VALUES`:
  se escribe el MetricValue del motor actual, nunca el valor cacheado histórico.
  Tests: unitarios con `FakeWorkbookWriter` (writer, workspace, VBA semántico,
  ciclo de vida COM, espera de recálculo) + 2 de integración COM marcados
  `@pytest.mark.excel` (se saltan fuera de Windows). Ver
  [`docs/EXCEL_WRITER.md`](EXCEL_WRITER.md). ✅ **Validado end-to-end en Windows +
  Excel Desktop**: `GENERATED_DIAGNOSTIC_REFERENCE`, `WRITER_INTEGRITY_PASS`,
  `CONTROL: PASS_INTERNAL_VALIDATION`, 1122/1122 celdas verificadas, plantilla y
  fórmulas/VBA preservados.
- **3.8 — Production Runtime Decoupling: ✅ COMPLETE.** El path de producción
  (`services/production_pipeline.py` + `services/runtime_assets.py`) construye
  los 1122 `PendingWrite` desde **sólo** el export Medinet + assets de runtime
  versionados en `config/runtime_2026/` (`bundle.yaml`, `write_manifest.csv`
  1122 filas, `metric_catalog.csv` 1224 fórmulas legacy, `detail_contract.yaml`,
  `zero_policy.yaml`). **No** abre `GENERACION DATOS REMASEP.xlsx`, **no** lee
  `artifacts/`. Assets sin PII (escaneo anti-PII en
  `scripts/build_runtime_assets.py`, que también auto-verifica OLD-vs-NEW).
  Validación en runtime → `RuntimeAssetError`
  (`RUNTIME_ASSET_MISSING / _INCOMPATIBLE / _INVALID`) con mensaje legible, nunca
  `FileNotFoundError` crudo. `resolve_runtime_root()` funciona desde repo y desde
  bundle PyInstaller (`sys._MEIPASS`), sin depender del CWD. El workbook legacy
  queda **dev-only** (`generate_remasep.py --mode diagnostic`, import perezoso).
  Equivalencia exacta: mismo Medinet → mismos 1122 metric_id / valores /
  instruction_id / target_sheet-cell que el path validado (scope 2114, sin
  filtro ESTADO — que sigue `PENDING_FUNCTIONAL_CONFIRMATION`). Ver
  [`docs/RUNTIME_DECOUPLING.md`](RUNTIME_DECOUPLING.md).
- **3.9 — Medinet Functional Validation (ESTADO): ✅ COMPLETE / CLOSED.**
  - *Fase 1 (auditoría, sólo lectura)*: `scripts/audit_medinet_estado.py` —
    distribución de los 2 114 registros de julio por ESTADO (1 337 `Atendido`,
    545 `Cancelado`, 182 `No Se Presenta`, 50 el resto); la hipótesis legacy da
    **exactamente 1 364**; impacto de **122 de 1 122** celdas (Σ = 516).
  - *Fase 2 (regla confirmada e implementada)*: el cliente (Fundación Gantz /
    Jacqueline) confirmó contar sólo `Atendido` · `En Sala de Espera` ·
    `Atención Pausada` · `En Atención` (excluir `Cancelado` · `No Se Presenta` ·
    `Agendado` · `Confirmado` · `Re-Agendado`). Regla **versionada** en
    `config/runtime_2026/estado_filter.yaml` (`status: CONFIRMED`, sha256 en
    `bundle.yaml`), aplicada en `production_pipeline.apply_estado_filter` tras
    "válidos ∩ período" y antes de calcular MetricValues. Normalización
    `str.strip().casefold()` (sin aliases). Julio 2026:
    `processing_scope_records` **2 114 → 1 364**; producción reproduce **exacto**
    el escenario `LEGACY_STATE_HYPOTHESIS` de la fase 1 (1 122
    MetricValues/PendingWrites, **187 no-cero**, **Σ = 928**). `estado_filter_status`
    del runtime bundle = **`CONFIRMED`**. (El del Sprint 3.6 sigue
    `PENDING_FUNCTIONAL_CONFIRMATION` a propósito: otra capa.) Ver
    [`docs/MEDINET_ESTADO_AUDIT.md`](MEDINET_ESTADO_AUDIT.md).
- **3.10 — UI profesional MEDINET + resumen mensual: ✅ COMPLETE.**
  Flujo de escritorio sin terminal: Inicio → Nuevo informe (período + archivo
  Medinet + plantilla REMASEP con validación de fingerprint) → Análisis
  (progreso comprensible, sin logs) → **Resumen mensual Medinet** → Generar →
  Resultado (Abrir Excel / Abrir carpeta). Modelo
  `services/medinet_summary.py::MonthlyMedinetSummary` = **adaptador de sólo
  lectura** sobre `processing_scope_frame` + `apply_estado_filter` (regla ESTADO
  3.9); julio 2026 reproduce **exacto** el pipeline (2 114 / 1 364 / 750 /
  64,5 %). 4 KPIs + 4 gráficos agregados **sin PII** (ESTADO, especialidad, sexo,
  edad; `PRESTACIÓN` descartada por 302/1 364 vacías). Export **PDF** de una
  página con `QPdfWriter`/`QPainter` (sin WebEngine; `report_content` separado del
  pintado; sin sobrescritura silenciosa). Generación vía `GenerationService`
  (`ui/workers.py::run_generation`, función pura; *graceful* sin Excel). Errores
  backend → `HumanError` (`ui/errors.py`; detalle técnico sólo al log). Backend,
  reglas REMASEP y `scripts/generate_remasep.py` **sin cambios**. Ver
  [`docs/UI.md`](UI.md).
  - *Patch de aceptación Windows*: (1) **detección asistida del período** al
    elegir el Medinet — `detect_medinet_periods` (reusa la validez de
    `processing_scope_frame`, nunca el nombre del archivo); autoselección si hay
    un solo mes, aviso si hay varios, "Analizar datos" bloqueado si el mes
    elegido no está en el archivo. (2) **"Guardar como"** antes de iniciar Excel
    COM: la ruta de salida se elige en un `QFileDialog` y se pasa explícita a
    `GenerationService`; sin overwrite silencioso; "Elegir otro nombre" reabre el
    diálogo sin perder el análisis. (3) warning `QFont setPointSize -1`: reglas
    QSS con `font-weight` sin `font-size` → añadido `font-size` explícito en 3
    reglas, sin cambio visual.
- **3.7+ — dimensiones de actividad / especialidad** y traducción a la tabla
  larga semántica: **pendiente**.
- **Modelo de *provenance* / `source`**: `MEDINET` se aplica en 3.1–3.3; el delta
  del REMASEP final (3.4) clasifica `EGRESOS` / `RESOURCE_CALCULATION` /
  `SURGICAL_TABLE` / `CONTROL_METADATA` / `UNKNOWN_PENDING` / `MIXED_DERIVED`.
  Todavía **sin** adaptador de egresos ni lector de tabla quirúrgica.

## Pre-Sprint 3 — inventario de fuentes / golden dataset (pendiente en paralelo)

Antes de avanzar en la generación oficial hay que completar el inventario de las
fuentes reales:

- inspeccionar el **REMASEP julio 2026 terminado** que entregará el cliente
  (candidato a golden reference — ver
  [`docs/CLIENT_PENDING.md`](CLIENT_PENDING.md));
- inspeccionar un ejemplo real de **egresos hospitalarios**, si está disponible;
- inspeccionar la **tabla quirúrgica** de julio 2026, si está disponible;
- obtener el **export Medinet original** de julio 2026;
- definir el **criterio de aceptación** del dataset golden.

## Luego

- **Sprint 3.4+ — dimensiones de actividad / especialidad**: traducir
  `semantic_signature` + `row_path` a `formulario · categoría · especialidad ·
  sexo · edad · modalidad · valor`,
  usando la evidencia de fórmula de 3.1.
- Semantic mapping hacia la plantilla oficial.
- Adaptador de egresos + cálculo de recursos.
- Escritura vía Excel COM + verificación `CONTROL`.
- Piloto en paralelo + empaquetado.

## Pendiente (backlog de fase)

- Mapping semántico de métricas.
- **Modelo de *provenance* / `source`** (ver
  [`docs/TECHNICAL_OVERVIEW.md`](TECHNICAL_OVERVIEW.md) → *Modelo futuro de
  métricas*): cada métrica deberá conservar de qué fuente proviene
  (`MEDINET` / `EGRESOS` / `RESOURCE_CALCULATION`). Requisito de diseño, sin
  código todavía.
- Adaptador de egresos hospitalarios.
- Cálculo de recursos / utilización de pabellones.
- Mapping semántico hacia la plantilla oficial MINSAL.
- Salida vía Excel COM (Windows).
- Validación `CONTROL`.
- Validación contra el dataset golden.
- Piloto manual/automático en paralelo.
- Empaquetado `.exe`.

## Lo que el sistema todavía **no** hace

- No genera todavía un REMASEP oficial entregable: el Excel writer (3.7B) está
  implementado y probado con dobles, pero **falta la corrida real en Windows con
  Excel** y el archivo se marca `NOT_FOR_SUBMISSION` mientras el filtro por
  ESTADO siga pendiente. La UI no está conectada.
- No integra egresos ni recursos/pabellones (CONTROL reportará errores por eso).
- No valida las reglas legacy contra reglas oficiales MINSAL.
- La equivalencia demostrada es **numérica contra el workbook generador**, con la
  cache de `AF` no verificable — no es una validación funcional del REMASEP.
