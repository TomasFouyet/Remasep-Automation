# Estado del proyecto

Vista de estado post Sprint 3.1 (inicio de la capa semántica). Para el detalle
por sprint ver [`docs/sprints/`](sprints/README.md).

## Fases

| Fase | Contenido | Estado |
| --- | --- | --- |
| **1 — Reverse Engineering** | Inventario del workbook, dependencias, lógica legacy, plantilla oficial, UI shell | **COMPLETE** |
| **2 — Medinet & Legacy Equivalence** | Ingesta Medinet real, `AC:AL`, agregaciones directas, cierre por dependencias, `SUM`/`IF` downstream | **COMPLETE** |
| **3 — Semantic Metrics** | Inventario semántico (3.1 ✅), dimensiones clínicas (3.2), modelo con *provenance* | **IN PROGRESS** |
| **4 — Additional Sources / Complete Dataset** | Adaptador de egresos, cálculo de recursos, dataset golden | PENDING |
| **5 — Official Template Mapping** | Mapping semántico generador/Medinet → plantilla MINSAL | PENDING |
| **6 — Excel Generation / CONTROL** | Escritura vía Excel COM (Windows), recálculo, verificación `CONTROL` | PENDING |
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
  duplicadas, 0 conflictos rótulo/fórmula. **No** es un mapping validado: no hay
  dimensiones clínicas todavía.
- **3.2 — dimensiones clínicas** (traducir los rótulos a sexo / grupo de edad /
  actividad / especialidad): **pendiente**.
- **Modelo de *provenance* / `source`**: `MEDINET` ya se aplica en el inventario
  de 3.1; `EGRESOS` / `RESOURCE_CALCULATION` quedan como requisito de diseño.

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

- **Sprint 3.2 — dimensiones semánticas**: traducir `semantic_signature` a
  `formulario · categoría · especialidad · sexo · edad · modalidad · valor`,
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

- No genera el REMASEP oficial (no hay mapping ni escritura Excel).
- No integra egresos ni recursos/pabellones.
- No valida las reglas legacy contra reglas oficiales MINSAL.
- La equivalencia demostrada es **numérica contra el workbook generador**, con la
  cache de `AF` no verificable — no es una validación funcional del REMASEP.
