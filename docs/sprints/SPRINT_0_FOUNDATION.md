# Sprint 0 — Base técnica

## Objetivo

Levantar el esqueleto del proyecto: estructura de paquetes, motor de reglas,
configuración versionada, adaptadores stub y la disciplina de tests/lint, todo
bajo una filosofía **local-first** (sin nube, sin telemetría, sin servidor).

## Problema que resolvía

Partir de cero con una arquitectura que separe con claridad la lógica de negocio
de los formatos Excel, para que las decisiones futuras (mapping oficial,
integración COM, nuevas fuentes) no obliguen a reescribir el núcleo.

## Implementación

- Proyecto Python empaquetable (`pyproject.toml`, `src/` layout, entry point
  `remasep`).
- Separación por capas: `core/` (normalización, motor de reglas, agregación en
  tabla larga), `domain/` (modelos), `adapters/` (Medinet, wrapper Excel COM),
  `services/`, `ui/`.
- Motor de reglas con **bloqueo de ambigüedad**: 0 coincidencias → bloquear,
  1 → continuar, >1 → bloquear.
- `audit.py` para trazabilidad; `config/loader.py` + plantillas de configuración
  versionada por versión de REMASEP.
- Adaptador Excel COM (`adapters/excel_com.py`) como stub para la fase Windows.
- Stub de UI (`ui/main_window.py`).
- `pytest` + `ruff` desde el primer commit.

## Archivos principales

- [`src/remasep/core/rules.py`](../../src/remasep/core/rules.py),
  [`src/remasep/core/aggregation.py`](../../src/remasep/core/aggregation.py),
  [`src/remasep/core/text.py`](../../src/remasep/core/text.py)
- [`src/remasep/domain/models.py`](../../src/remasep/domain/models.py)
- [`src/remasep/adapters/excel_com.py`](../../src/remasep/adapters/excel_com.py),
  [`src/remasep/adapters/medinet.py`](../../src/remasep/adapters/medinet.py)
- [`src/remasep/audit.py`](../../src/remasep/audit.py),
  [`src/remasep/config/loader.py`](../../src/remasep/config/loader.py)
- [`docs/ARCHITECTURE.md`](../ARCHITECTURE.md), [`docs/ROADMAP.md`](../ROADMAP.md),
  [`docs/CLIENT_PENDING.md`](../CLIENT_PENDING.md)

## Decisiones técnicas

- **La lógica de negocio nunca conoce coordenadas Excel** (`G54`): el mapping por
  versión traduce cada métrica semántica a una celda/rango.
- Métricas en **tabla larga** (`formulario | categoria | especialidad | sexo |
  edad | modalidad | valor`).
- Configuración **versionada por REMASEP** en `config/`, no hardcodeada.
- Linux/Ubuntu para desarrollo; Windows + Microsoft Excel sólo para la
  generación oficial final (COM). No LibreOffice como sustituto.

## Tests

`tests/test_rule_engine.py`, `tests/test_aggregation.py` — motor de reglas y
agregación en tabla larga.

## Resultado sobre workbook de referencia

N/A — Sprint 0 no procesa el workbook real.

## Hallazgos

Ninguno sobre datos; se fija la arquitectura Extract → Transform → Metrics →
Map → Load.

## Limitaciones

- Adaptadores Medinet/COM son stubs.
- Sin reglas clínicas ni mapping oficial.

## Estado final

**COMPLETE.** Base técnica y disciplina de calidad (`pytest` / `ruff`)
establecidas.
