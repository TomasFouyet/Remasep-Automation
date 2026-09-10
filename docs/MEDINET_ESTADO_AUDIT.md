# Campo ESTADO — Medinet (Sprint 3.9) — **CONFIRMED / CLOSED**

## Estado

- **Fase 1** (auditoría, sólo lectura): evidencia recogida — ver abajo.
- **Fase 2** (regla confirmada e implementada):
  - **Confirmación funcional del cliente** (Fundación Gantz / Jacqueline): para el
    REMASEP se cuentan **sólo** las citas cuyo ESTADO sea
    **`Atendido` · `En Sala de Espera` · `Atención Pausada` · `En Atención`**; se
    excluyen `Cancelado` · `No Se Presenta` · `Agendado` · `Confirmado` ·
    `Re-Agendado`.
  - Regla **versionada** en `config/runtime_2026/estado_filter.yaml`
    (`status: CONFIRMED`), aplicada en `production_pipeline.apply_estado_filter`
    después de "válidos ∩ período" y antes de calcular MetricValues.
  - `estado_filter_status` del runtime bundle = **`CONFIRMED`**.
  - Julio 2026: `processing_scope_records` **2 114 → 1 364**; el resultado es
    idéntico al escenario `LEGACY_STATE_HYPOTHESIS` auditado en fase 1
    (1 122 MetricValues, 1 122 PendingWrites, 187 no-cero, Σ = 928).

> El `estado_filter_status` del **Sprint 3.6** (`writable_target_mapping`) queda
> en `PENDING_FUNCTIONAL_CONFIRMATION` a propósito: describe otra capa (qué celdas
> son *destino* de escritura, que no depende del ESTADO), no la selección de
> población.

---

## Fase 1 — auditoría (sólo lectura)

`scripts/audit_medinet_estado.py`. Evidencia que se llevó al cliente.

- Medinet: `data/local/detalle_citas - 2026-09-07T123630.940.xlsx`
- Período: Julio 2026, recortado **exactamente** como
  `production_pipeline` (`processing_scope_frame` = registros estructuralmente
  válidos ∩ mes/año).

## 1. Distribución por ESTADO

| ESTADO | cantidad | % del total |
| --- | ---: | ---: |
| Atendido | 1 337 | 63.25 % |
| Cancelado | 545 | 25.78 % |
| No Se Presenta | 182 | 8.61 % |
| Atención Pausada | 16 | 0.76 % |
| Agendado | 11 | 0.52 % |
| Confirmado | 7 | 0.33 % |
| En Sala de Espera | 7 | 0.33 % |
| Re-Agendado | 5 | 0.24 % |
| En Atención | 4 | 0.19 % |
| **TOTAL** | **2 114** | **100.00 %** |

**Confirmado: la suma es 2 114** (= `processing_scope_records` actual de
producción).

## 2. Hipótesis legacy observada

| | estados | registros |
| --- | --- | ---: |
| **INCLUIDOS** | Atendido · Atención Pausada · En Sala de Espera · En Atención | **1 364** |
| **EXCLUIDOS** | Cancelado · No Se Presenta · Agendado · Confirmado · Re-Agendado | 750 |
| no cubiertos por la hipótesis | — | 0 |

**Confirmado: la hipótesis produce exactamente 1 364 registros** (todos los
ESTADO del archivo quedan clasificados; no hay estados sueltos).

Matiz: de esos 1 364, **1 337 son `Atendido`** y sólo **27** son los otros tres
estados (`Atención Pausada` 16 + `En Sala de Espera` 7 + `En Atención` 4).

## 3. Impacto sobre las métricas REMASEP (pipeline actual, sin modificar producción)

| alcance | scope | MetricValues | PendingWrites | MetricValues ≠ 0 | Σ de todos los conteos |
| --- | ---: | ---: | ---: | ---: | ---: |
| **A. PERIOD_ONLY** (producción hoy) | 2 114 | 1 122 | 1 122 | 210 | 1 444 |
| **B. LEGACY_STATE_HYPOTHESIS** | 1 364 | 1 122 | 1 122 | 187 | 928 |

- **122 de las 1 122 métricas cambian** entre A y B (10.9 %).
- **Todas bajan** al aplicar el filtro (A ≥ B): B es un subconjunto de A, nunca
  aparecen conteos nuevos.
- **23 métricas pasan de un valor > 0 a 0**; las otras **99 sólo se reducen**.
- Magnitud de la diferencia: **Σ |Δ| = 516 conteos** (= 1 444 − 928).
  - 57 métricas con Δ = 1 · 29 con Δ 2–3 · 27 con Δ 4–10 · **9 con Δ > 10**
    (hasta Δ = 55).

### Por hoja destino

| hoja | métricas que cambian | Σ \|Δ\| |
| --- | ---: | ---: |
| REMASEP_OD | 82 | 199 |
| REMASEP 01 | 37 | 230 |
| B2 ANEXO | 3 | 87 |

### Métricas más afectadas

| destino | descripción | A | B | Δ |
| --- | --- | ---: | ---: | ---: |
| B2 ANEXO!C953 | Consulta psicólogo clínico (sesiones 45') — 0902001 | 98 | 43 | 55 |
| REMASEP 01!H118 | Consultas fonoaudiólogo/a — por edad | 80 | 52 | 28 |
| REMASEP 01!F118 | Consultas fonoaudiólogo/a — por edad | 60 | 35 | 25 |
| REMASEP 01!G118 | Consultas fonoaudiólogo/a — por edad | 66 | 43 | 23 |
| REMASEP_OD!U71 | Controles ortodoncia OPI — 10-14 años | 88 | 66 | 22 |
| B2 ANEXO!C833 | Atención kinesiológica integral ambulatoria — 0601105 | 51 | 29 | 22 |
| REMASEP 01!I118 | Consultas fonoaudiólogo/a — por edad | 54 | 32 | 22 |
| REMASEP_OD!W71 | Controles ortodoncia OPI — 15-19 años | 52 | 39 | 13 |
| REMASEP 01!J118 | Consultas fonoaudiólogo/a — por edad | 34 | 21 | 13 |
| B2 ANEXO!C957 | Telerehabilitación psicólogo clínico — 0908101 | 16 | 6 | 10 |

Las áreas más sensibles son **consultas de profesionales no médicos**
(fonoaudiología, kinesiología, psicología — `REMASEP 01` Sección C y `B2 ANEXO`)
y **controles de ortodoncia** (`REMASEP_OD`).

## 4. Variante: sólo `Atendido`

Para dimensionar cuánto pesa la parte "no-`Atendido`" de la hipótesis:

| comparación | métricas que cambian | Σ \|Δ\| | registros |
| --- | ---: | ---: | ---: |
| A (2 114) vs B (4 estados, 1 364) | 122 | 516 | −750 |
| A (2 114) vs C (sólo `Atendido`, 1 337) | 123 | 526 | −777 |
| **B vs C** | **5** | **10** | **−27** |

Es decir: la decisión **numéricamente relevante** es *"contar todas las citas del
período"* vs *"contar sólo atenciones realizadas / con presencia del paciente"*
(≈ 750 registros, 516 conteos). Que el criterio sea exactamente los 4 estados de
la hipótesis o sólo `Atendido` casi no mueve la aguja (5 métricas, 10 conteos).

## 5. Lectura (fase 1)

- El export directo cuenta **toda cita agendada del período**, con cualquier
  ESTADO. Si el REMASEP debe reflejar **actividad efectivamente realizada**, las
  citas `Cancelado` (545) y `No Se Presenta` (182) inflarían los conteos.
- La hipótesis legacy incluía estados que no son "atención terminada" y excluía
  `Agendado` / `Confirmado` / `Re-Agendado`.

## 6. Respuesta del cliente (cierre — fase 2)

> **Confirmado (Fundación Gantz / Jacqueline)**: para el REMASEP se consideran
> únicamente los registros cuyo ESTADO sea **`Atendido`**, **`En Sala de
> Espera`**, **`Atención Pausada`** o **`En Atención`**. Se excluyen `Cancelado`,
> `No Se Presenta`, `Agendado`, `Confirmado` y `Re-Agendado`.

Coincide exactamente con la hipótesis legacy → Julio 2026 pasa de 2 114 a
**1 364** registros. Regla implementada y versionada (fase 2, ver *Estado*
arriba).

Pregunta original que se llevó al cliente (para registro):

> En el export de Medinet ("Detalle de citas"), cada cita del mes trae un campo
> **ESTADO**. Para julio 2026 hay **2 114** citas en total:
> **1 337 `Atendido`**, **545 `Cancelado`**, **182 `No Se Presenta`** y **50**
> repartidas entre `Atención Pausada`, `En Sala de Espera`, `En Atención`,
> `Agendado`, `Confirmado` y `Re-Agendado`.
>
> **¿Qué estados deben contarse como "atención realizada" para el REMASEP?**
> ¿Sólo `Atendido`? ¿También `Atención Pausada` / `En Sala de Espera` /
> `En Atención`? ¿Y las citas `Cancelado` / `No Se Presenta` se descartan
> siempre?
>
> (Impacto: cambia el valor de **122 de 1 122 celdas** del REMASEP —
> principalmente consultas de fonoaudiología, kinesiología, psicología y
> controles de ortodoncia—, con una diferencia total de **516 atenciones**
> contadas.)

## 7. Reproducir

Auditoría (fase 1):

    python scripts/audit_medinet_estado.py \
        --medinet "data/local/detalle_citas - 2026-09-07T123630.940.xlsx" \
        --period 2026-07

Regenerar los assets de runtime con la regla ESTADO confirmada (fase 2):

    python scripts/build_runtime_assets.py

Comprobación en producción:

    python -c "from remasep.services.common import Period; \
from remasep.services.production_pipeline import build_production_pending_writes as b; \
r=b('data/local/detalle_citas - 2026-09-07T123630.940.xlsx', Period(7,2026)); \
print(r.scope.period_scope_records, '->', r.scope.processing_scope_records, \
len(r.run.metric_values), sum(1 for x in r.run.metric_values if x.value), \
sum(x.value for x in r.run.metric_values))"
    # 2114 -> 1364 1122 187 928
