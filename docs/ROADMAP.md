# Roadmap

Plan por sprints, **sin fechas**. Estado detallado en
[`docs/PROGRESS.md`](PROGRESS.md); un documento por sprint en
[`docs/sprints/`](sprints/README.md).

## Sprint 0 — Base técnica — ✅ COMPLETE
- [x] repositorio
- [x] modelo de reglas
- [x] bloqueo de desconocidos
- [x] bloqueo de ambigüedades
- [x] tabla larga
- [x] adaptador Medinet preliminar
- [x] wrapper Excel COM (stub)
- [x] auditoría
- [x] configuración versionada
- [x] tests básicos

## Sprint 1 — Ingeniería inversa — ✅ COMPLETE
- [x] 1.1 inventario del workbook generador
- [x] 1.2 análisis de dependencias de fórmulas
- [x] 1.3 inventario de la lógica legacy actual (27 reglas candidatas)
- [x] 1.4 inventario y alineación de la plantilla oficial
- [x] UI — application shell navegable (datos mock)

## Sprint 2 — Equivalencia Medinet ↔ legacy — ✅ COMPLETE
- [x] 2.1 análisis real de Medinet (1364 atenciones, julio 2026)
- [x] 2.2 equivalencia derivada `AC:AL` (cache de `AF` no verificable)
- [x] 2.3 equivalencia de agregaciones directas `COUNTIF`/`COUNTIFS`
- [x] 2.4 cierre por dependencias — agregaciones derivadas + DAG
- [x] 2.5 equivalencia downstream `SUM` / `IF` — **2043 / 2043 celdas evaluables**

## Pre-Sprint 3 — Inventario de fuentes / golden dataset — ⏳ NEXT
Antes de empezar Fase 3:
- inspeccionar el **REMASEP julio 2026 terminado** (candidato a golden reference);
- inspeccionar un ejemplo real de **egresos hospitalarios**;
- inspeccionar la **tabla quirúrgica** de julio 2026;
- obtener el **export Medinet original** de julio 2026;
- definir el **criterio de aceptación** del dataset golden.
Ver [`docs/CLIENT_PENDING.md`](CLIENT_PENDING.md).

## Sprint 3 — Métricas semánticas — PENDIENTE
- 3.1 **Semantic Metric Inventory**: catálogo de métricas
  (`formulario · categoría · especialidad · sexo · edad · modalidad · valor`)
  sobre el DAG de métricas, no sobre el texto de fórmula.
- Modelo de *provenance* / `source` (`MEDINET` / `EGRESOS` /
  `RESOURCE_CALCULATION`) — requisito de diseño.

## Sprint 4 — Mapping semántico — PENDIENTE
- Fingerprint de la plantilla `REMASEP 2026_V1.4`.
- Mapping semántico métrica → celda/rango (no posicional).
- Resolución de los candidatos `ambiguous` de Sprint 1.4.

## Sprint 5 — Fuentes adicionales — PENDIENTE
Sólo tras recibir ejemplos reales:
- adaptador de **egresos hospitalarios** (previsión → REMASEP 01; edad/sexo →
  REMASEP B1; códigos Qx → B2 ANEXO);
- **cálculo de recursos / utilización de pabellones** (tabla quirúrgica).

## Sprint 6 — Generación oficial — PENDIENTE
- Escritura vía Excel COM (`DispatchEx`) en Windows.
- Recálculo y guardado de una copia.
- Verificación final en la hoja `CONTROL`.

## Sprint 7 — Piloto y empaquetado — PENDIENTE
- Proceso manual y automático en paralelo por varios meses.
- Wizard para usuario no técnico.
- Empaquetado `.exe` (PyInstaller), firma, despliegue institucional.
