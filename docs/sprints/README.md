# Sprints — índice

Un documento por sprint, con la misma estructura (Objetivo · Problema ·
Implementación · Archivos · Decisiones · Tests · Resultado · Hallazgos ·
Limitaciones · Estado). Los sprints resumen y enlazan a la documentación técnica
especializada de [`docs/`](../).

| Sprint | Objetivo | Estado | Resultado clave | Documento |
| --- | --- | --- | --- | --- |
| **0** | Base técnica y disciplina de calidad | COMPLETE | Arquitectura Extract→Transform→Metrics→Map→Load; `pytest`/`ruff` | [SPRINT_0_FOUNDATION](SPRINT_0_FOUNDATION.md) |
| **1.1** | Inventario del workbook generador | COMPLETE | 4 hojas · 54 682 celdas no vacías · 22 642 fórmulas · 39 patrones · 0 named ranges | [SPRINT_1_1_WORKBOOK_INVENTORY](SPRINT_1_1_WORKBOOK_INVENTORY.md) |
| **1.2** | Grafo de dependencias de fórmulas | COMPLETE | 10 columnas derivadas `AC:AL`; `Atenciones` → `REMASEP_OD`/`REMASEP 01`/`B2 ANEXO` | [SPRINT_1_2_DEPENDENCY_ANALYSIS](SPRINT_1_2_DEPENDENCY_ANALYSIS.md) |
| **1.3** | Inventario de la lógica legacy actual | COMPLETE | 7 campos raw transitivos; 27 reglas candidatas (compatibilidad ≠ MINSAL) | [SPRINT_1_3_CURRENT_LOGIC](SPRINT_1_3_CURRENT_LOGIC.md) |
| **1.4** | Inventario y alineación de la plantilla oficial | COMPLETE | 11 hojas · 11 466 fórmulas · VBA presente · 230 exact_label / 1189 ambiguous · mapping NO implementado | [SPRINT_1_4_TEMPLATE_ALIGNMENT](SPRINT_1_4_TEMPLATE_ALIGNMENT.md) |
| **1 UI** | Maqueta de escritorio navegable | COMPLETE | PySide6, flujo Home→New Report→Analysis→Exceptions→Summary, demo funcional | [SPRINT_1_UI_APPLICATION_SHELL](SPRINT_1_UI_APPLICATION_SHELL.md) |
| **2.1** | Análisis real de Medinet | COMPLETE | 2006 físicas = 642 vacías + **1364 atenciones** (julio 2026); 485 con match legacy | [SPRINT_2_1_MEDINET_ANALYSIS](SPRINT_2_1_MEDINET_ANALYSIS.md) |
| **2.2** | Equivalencia derivada `AC:AL` | COMPLETE | `AC:AL` (salvo `AF`) 1364 / 1364 exactas; `AF` cache obsoleta (571) | [SPRINT_2_2_DERIVED_EQUIVALENCE](SPRINT_2_2_DERIVED_EQUIVALENCE.md) |
| **2.3** | Equivalencia de agregaciones directas | COMPLETE | 1427 celdas directas; 1325 soportadas, 102 derivadas fuera de alcance | [SPRINT_2_3_AGGREGATION_EQUIVALENCE](SPRINT_2_3_AGGREGATION_EQUIVALENCE.md) |
| **2.4** | Cierre por dependencias | COMPLETE | 102 / 102 derivadas; universo **2043** (1427 directas + 616 transitivas); DAG depth 4, 0 ciclos | [SPRINT_2_4_AGGREGATION_CLOSURE](SPRINT_2_4_AGGREGATION_CLOSURE.md) |
| **2.5** | Equivalencia downstream `SUM`/`IF` | COMPLETE | **2043 / 2043** celdas evaluables (344 `SUM` + 272 `IF`); `formula_support_status: PASS` | [SPRINT_2_5_DOWNSTREAM_EQUIVALENCE](SPRINT_2_5_DOWNSTREAM_EQUIVALENCE.md) |

## Fase actual

- **Fase 1 — Reverse Engineering** (Sprint 1.1–1.4): **COMPLETE**
- **Fase 2 — Medinet & Legacy Equivalence** (Sprint 2.1–2.5): **COMPLETE**
- **Fase 3 — Semantic Metrics**: *NEXT* (precedida por el inventario completo de
  fuentes / golden dataset — ver [`docs/PROGRESS.md`](../PROGRESS.md))

Sprint 3 **no** está iniciado.
