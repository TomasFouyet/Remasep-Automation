# REMASEP Automation

Aplicación de escritorio **local** para automatizar la preparación mensual del
reporte **REMASEP** de Fundación Gantz, hoy elaborado a mano en Excel a partir de
fuentes operacionales (**Medinet**, **egresos hospitalarios**, **tabla
quirúrgica / recursos de pabellón**).

El objetivo es reemplazar un proceso manual de varias horas por un flujo guiado:
**cargar fuentes → validar → revisar excepciones → generar el workbook oficial**.

> El sistema **todavía no genera el REMASEP final**. Lo que existe es la
> ingeniería inversa completa del workbook actual y la demostración de que su
> lógica de fórmulas derivada de Medinet es reproducible en Python. El mapping
> hacia la plantilla oficial MINSAL y la escritura en Excel **no** están
> implementados.

---

## Objetivo

Automatizar la construcción del REMASEP mensual: leer las fuentes institucionales,
validarlas y clasificarlas, calcular las métricas, mapearlas a la plantilla
oficial `REMASEP 2026_V1.4.xlsm` y generarla en Windows vía Excel COM, con una
verificación final en la hoja `CONTROL` — dejando al usuario sólo la revisión de
excepciones.

---

## Estado actual

| Fase | Contenido | Estado |
| --- | --- | --- |
| **1 — Reverse Engineering** | Inventario del workbook, dependencias, lógica legacy, plantilla oficial, UI shell | **COMPLETE** |
| **2 — Medinet & Legacy Equivalence** | Ingesta Medinet real, `AC:AL`, agregaciones directas/derivadas, `SUM`/`IF` downstream | **COMPLETE** |
| **3 — Semantic Metrics** | Inventario de métricas semánticas, modelo con *provenance* | **NEXT** |
| **4 — Additional Sources / Complete Dataset** | Adaptador de egresos, cálculo de recursos, dataset golden | PENDING |
| **5 — Official Template Mapping** | Mapping semántico → plantilla MINSAL | PENDING |
| **6 — Excel Generation / CONTROL** | Escritura Excel COM (Windows), recálculo, `CONTROL` | PENDING |
| **7 — Pilot / Packaging** | Piloto manual + automático en paralelo, `.exe` | PENDING |

Fase 3 va precedida de un **pre-Sprint 3** de inventario de fuentes / golden
dataset (ver [`docs/PROGRESS.md`](docs/PROGRESS.md)). Sprint 3 **no** está
iniciado.

---

## Capacidades demostradas

- Lectura de un archivo **Medinet** real (detección de hoja por encabezados,
  mapeo semántico, sin fuzzy).
- Identificación de **filas estructuralmente vacías** (fórmulas arrastradas).
- Cálculo de las transformaciones derivadas **`AC:AL`** exacto vs. el workbook.
- Reproducción de **`COUNTIF` / `COUNTIFS`** (agregaciones directas).
- **Agregaciones derivadas** (`COUNTIFS(...) − celda`) vía grafo de dependencias.
- **DAG de dependencias** entre celdas agregadas (orden topológico, ciclos,
  profundidad).
- **`SUM` downstream** y **`IF` validations** evaluados sobre el mismo DAG.
- **UI de escritorio** navegable (PySide6) con flujo completo.
- Separación **DEMO / REAL** en la UI.

Lo que **no** hace todavía: generar el REMASEP oficial, integrar egresos ni
recursos, ni validar contra reglas oficiales MINSAL.

---

## Dataset de referencia

`data/local/GENERACION DATOS REMASEP.xlsx`, hoja `Atenciones - Detalles de
citas`, período **julio 2026** (SHA256 `fc2e1536…d15b79`):

```
physical_rows_examined = 2006
structural_empty_rows  =  642
real records           = 1364   (= 1364 válidos, 0 inválidos)
```

> **Las 2006 filas físicas NO son 2006 atenciones.** Las atenciones reales de
> julio 2026 son **1364**; las otras 642 son filas con las fórmulas `AC:AL`
> arrastradas más allá de los datos.

El export **Medinet original** exacto de julio 2026 sigue **pendiente** de
obtener del cliente (se usó la hoja `Atenciones` del generador como sustituto).

---

## Cobertura legacy

**2043 / 2043** celdas Medinet-dependientes conocidas son **evaluables**
(`formula_support_status: PASS`):

| Clasificación | Celdas |
| --- | ---: |
| `BASE_AGGREGATION` (`COUNTIF`/`COUNTIFS`) | 1325 |
| `DERIVED_AGGREGATION` (`… − celda`) | 102 |
| `DOWNSTREAM_TOTAL` (`SUM`) | 344 |
| `VALIDATION` (`IF`) | 272 |
| **Total** | **2043** |

Universo = **1427 directas** (referencian `Atenciones`) + **616 transitivas**.
**DAG: profundidad máxima 4 · ciclos 0 · dependencias faltantes 0.**

---

## Cache de Excel

La equivalencia se compara contra el **valor cacheado** por Excel, no contra un
recálculo. **No** se afirma "100 % equivalente a Excel". Se separan dos cosas:

| Concepto | Resultado |
| --- | --- |
| **FORMULA SUPPORT** — ¿se pudo evaluar? | 2043 / 2043 → **PASS** |
| **CACHE CONSISTENCY** — ¿coincide con la cache? | **DIFFERENCES** |

| Estado de cache | Celdas |
| --- | ---: |
| `MATCH` | 1659 |
| `CACHE_DIFFERENCE` | 248 |
| `CACHE_UNAVAILABLE` | 136 |

- **Las 248 diferencias dependen todas de `AF` (edad).** La cache de `AF` está
  **obsoleta** en el workbook de referencia (Excel dejó `0` en 571 filas con
  fechas válidas). Las diferencias se anotan; **no** se convierten en `MATCH`.
- **Las 136 `CACHE_UNAVAILABLE`** son validaciones `IF` de resultado **texto**
  para las que Excel no guardó ningún valor cacheado.

Detalle: [`docs/LEGACY_DOWNSTREAM_EQUIVALENCE.md`](docs/LEGACY_DOWNSTREAM_EQUIVALENCE.md).

---

## Arquitectura

```text
EXTRACT    Medinet / Egresos / Recursos
    ↓
TRANSFORM  normalizar / validar / clasificar
    ↓
METRICS    tabla larga semántica (con provenance)
    ↓
MAP        métrica → celda/rango según versión REMASEP
    ↓
LOAD       Excel oficial / recálculo / CONTROL / guardar
```

- La lógica de negocio **no** conoce coordenadas Excel (`G54`).
- Una prestación desconocida o una clasificación ambigua **bloquean** el
  procesamiento (0 → bloquear · 1 → continuar · >1 → bloquear).
- Linux para desarrollo/análisis/UI; Windows + Microsoft Excel sólo para la
  generación oficial (COM).

Detalle: [`docs/ARCHITECTURE.md`](docs/ARCHITECTURE.md) ·
[`docs/TECHNICAL_OVERVIEW.md`](docs/TECHNICAL_OVERVIEW.md).

---

## Desarrollo

```bash
# Ubuntu / Linux
python3 -m venv .venv
source .venv/bin/activate
python -m pip install -e ".[dev]"

# UI (maqueta navegable, datos mock)
python -m remasep.main

# Tests y lint
pytest            # 379 passed
ruff check .      # limpio
```

Los tests de UI configuran Qt en modo **offscreen** automáticamente
(`tests/conftest.py`); no requieren display.

Los scripts de ingeniería inversa (`scripts/*.py`) son de **solo lectura** y
escriben a `artifacts/` (gitignored). Ejemplo:

```bash
python scripts/evaluate_legacy_downstream.py "data/local/GENERACION DATOS REMASEP.xlsx"
```

---

## Estructura del repositorio

```text
src/remasep/
  adapters/     medinet.py · excel_com.py (stub Windows)
  core/         text.py (normalize_text / normalize_legacy_text) · rules.py · aggregation.py
  domain/       models.py
  services/     legacy_transform · legacy_rules · legacy_aggregation · medinet_analysis · mock_remasep
  ui/           screens/ · components/ · main_window.py · styles.py
scripts/        inventory_workbook · formula_refs · analyze_dependencies · inventory_current_logic
                inventory_template · compare_legacy_derived · compare_legacy_aggregations
                close_legacy_aggregations · evaluate_legacy_downstream
config/         legacy_current_logic_2026/rules.yaml · plantillas versionadas
docs/           ver docs/README.md
tests/          workbooks sintéticos; sin data/local/
```

`data/` (archivos reales, con pacientes) y `artifacts/` están **gitignored**.

---

## Seguridad y privacidad

- Procesamiento **local**, sin telemetría.
- Ninguna herramienta exporta PII: sólo fórmulas, coordenadas, metadatos y
  **valores agregados**. `RecordProblem` referencia sólo `fila` / `campo` /
  `error_code`.

---

## Documentación

Índice completo: [`docs/README.md`](docs/README.md).

| Documento | Para qué sirve |
| --- | --- |
| [`docs/PROGRESS.md`](docs/PROGRESS.md) | Vista de estado por fases: completado · siguiente · pendiente |
| [`docs/ROADMAP.md`](docs/ROADMAP.md) | Plan por sprints (sin fechas) |
| [`docs/sprints/`](docs/sprints/README.md) | Un documento por sprint (0 → 2.5) |
| [`docs/TECHNICAL_OVERVIEW.md`](docs/TECHNICAL_OVERVIEW.md) | Stack, fronteras, motor legacy, DAG, cache, fuentes de datos |
| [`docs/CLIENT_PENDING.md`](docs/CLIENT_PENDING.md) | Preguntas funcionales resueltas y pendientes |
