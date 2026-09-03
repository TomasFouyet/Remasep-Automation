# Equivalencia de agregaciones legacy (Sprint 2.3)

Este documento describe cómo el proyecto demuestra que la lógica de **conteo
directo** del workbook `GENERACION DATOS REMASEP.xlsx` puede reproducirse en
Python a partir de la fuente Medinet, sin macros, sin COM y sin recalcular Excel.

> **Aviso.** Equivalencia con `GENERACION DATOS REMASEP.xlsx` **no** es validación
> oficial MINSAL. Sólo se demuestra que, dado el mismo conjunto de atenciones,
> las fórmulas COUNTIF/COUNTIFS del workbook y su reimplementación en Python
> producen el mismo número. La corrección *funcional* de esas fórmulas (que
> midan lo que MINSAL espera) sigue pendiente de validación.

## Alcance

Se cubren **únicamente** las celdas de las hojas `REMASEP 01`, `B2 ANEXO` y
`REMASEP_OD` cuya fórmula referencia **directamente** la hoja de detalle
`Atenciones - Detalles de citas`.

Quedan **fuera de alcance** de forma deliberada:

- fórmulas intermedias / totales internos de esas hojas (sumatorias de
  subtotales, restas entre celdas de la misma hoja, etc.);
- la plantilla oficial MINSAL (ver `docs/TEMPLATE_ALIGNMENT.md`);
- cualquier construcción que no sea COUNTIF/COUNTIFS o suma de ellas.

## Qué significa "direct aggregation"

Una celda es *direct source aggregation* cuando su fórmula:

1. es `COUNTIF(rango, criterio)`, `COUNTIFS(r1,c1,r2,c2,…)` o una **suma con `+`**
   de esas llamadas;
2. cada `rango` es una **columna entera** de la hoja de detalle
   (`'Atenciones - Detalles de citas'!$AG:$AG`, `!O:O`, …);
3. cada `criterio` es una concatenación con `&` de literales entre comillas y/o
   referencias a celdas de rótulo de la propia hoja (`$A26`, `$A$3`, …), que se
   resuelven a texto.

Todo lo demás en esas tres hojas se clasifica `UNSUPPORTED` (si usa COUNTIF/S
pero con algún construct fuera del subset) u `OTHER` (si ni siquiera referencia
la hoja de detalle con COUNTIF/S). No se evalúa a medias ni se infiere.

## Subset de Excel implementado

`src/remasep/services/legacy_aggregation.py` implementa sólo lo observado en el
workbook real:

| Elemento | Soporte |
|---|---|
| `COUNTIF` / `COUNTIFS` | sí |
| suma de varias llamadas con `+` | sí |
| rango = columna entera de la hoja de detalle (`$X:$X`, `X:X`) | sí |
| rango que no es columna entera (`$A2:$A100`) | `UNSUPPORTED` |
| rango que apunta a otra hoja | `UNSUPPORTED` |
| criterio de texto exacto | sí |
| comodines Excel `*` y `?` | sí |
| escape `~*`, `~?`, `~~` | sí |
| criterios numéricos `=`, `<>`, `<`, `<=`, `>`, `>=` | sí |
| criterio `<>texto` | sí |
| criterio con `&` (literal y/o referencia de celda) | sí |
| referencia de celda de criterio que no resuelve a valor | `UNSUPPORTED` |
| cualquier contenido tras el `COUNTIF(S)` (`…)-F30`) | `UNSUPPORTED` |
| cualquier otra función | `UNSUPPORTED` |

### Semántica de los criterios de texto

Consistente con la equivalencia legacy del Sprint 2.2
(`normalize_legacy_text` = `str(value).casefold()`):

- **case-insensitive** (`casefold`);
- **accent-sensitive**: `"máscara"` ≠ `"mascara"`;
- **whitespace-exacto**: `"control aparato"` ≠ `"control  aparato"` ni
  `" control aparato"`;
- comodines Excel (`*` = cualquier secuencia, `?` = un carácter), anclados al
  texto completo; `~` escapa el siguiente comodín;
- **sin** `normalize_text` global (que quita tildes y colapsa espacios) y **sin**
  fuzzy matching.

## Dataset de evaluación

- Se abre el workbook dos veces con openpyxl: `data_only=False` (fórmulas) y
  `data_only=True` (valores cacheados). Sólo lectura.
- La hoja de detalle se identifica por **encabezados** (no por letra de
  columna), igual que `MedinetAdapter`.
- Los campos crudos (`SEXO`, `SUCURSAL`, `ESPECIALIDAD`, `TIPO_DE_CITA`,
  `PRESTACION`, …) se preservan **sin recortar** whitespace.
- Las columnas derivadas `AC:AL` se calculan con
  `src/remasep/services/legacy_transform.py` (concatenaciones, `IFS`, edad
  `DATEDIF`, conteo de patrones), no se leen del workbook.
- Las **filas estructuralmente vacías** (todas las columnas semánticas en
  blanco: fórmulas `AC:AL` "arrastradas" más allá de las atenciones reales) se
  excluyen del conteo. No se codifican números fijos; se miden por archivo.

En el workbook de referencia: **2006 filas físicas = 642 estructuralmente
vacías + 1364 activas**.

### Sensibilidad al padding

Para cada fórmula soportada se evalúa si alguna fila estructuralmente vacía
*podría* contar, aplicando sus criterios a los valores que Excel deja en una
fila vacía (`AC:AE = ""`, `AF = 0`, `AG:AI = #N/A`, `AJ:AL = "0"`).

En el workbook de referencia: **0 fórmulas sensibles al padding**. Toda fórmula
directa filtra por una columna de categoría (`AC…AL`) que en una fila vacía vale
`""`, `#N/A` o `"0"` y por tanto nunca cuenta. El artefacto
`padding_sensitivity.csv` lo registra fórmula a fórmula; los tests incluyen
además un caso sintético (`COUNTIF($AF:$AF,"<10")`) donde el padding **sí**
afectaría, para verificar que la detección no es trivialmente `False`.

## Evaluator vs. cache

Se mantienen **dos estados separados** por celda:

- `formula_evaluation_status` — ¿pudo Python evaluar la fórmula? (`supported` /
  `UNSUPPORTED`).
- `cache_comparison_status` — comparación del valor Python con el valor cacheado
  por Excel: `MATCH`, `CACHE_DIFFERENCE` o `CACHE_UNAVAILABLE` (sin cache).

Una `CACHE_DIFFERENCE` **no** se llama "stale" salvo que haya evidencia. El
único caso con evidencia conocida es la columna de edad.

### Impacto conocido de la cache obsoleta de `AF`

El Sprint 2.2 demostró que el workbook tiene la **cache de `AF` (edad) obsoleta**
para una parte de los registros (Excel dejó `0` en filas con fecha válida). El
evaluador Python usa la edad **recalculada** por `legacy_transform`.

Por eso, cuando una fórmula `depends_on_age` (referencia `!$AF:$AF`) produce una
`CACHE_DIFFERENCE`, se registra con la nota

> `Formula depends on AF; workbook reference contains known stale AF cache.`

y **no** se convierte automáticamente a `MATCH`. La diferencia se cuenta con
normalidad; la nota explica su causa probable.

## Resultado del workbook de referencia

`GENERACION DATOS REMASEP.xlsx` (sha256 `fc2e1536…d15b79`):

| Hoja | Directas | Soportadas | No soportadas | Cache MATCH | Cache DIFF | Cache UNAVAIL |
|---|---:|---:|---:|---:|---:|---:|
| REMASEP 01 | 247 | 247 | 0 | 186 | 61 | 0 |
| B2 ANEXO | 24 | 24 | 0 | 24 | 0 | 0 |
| REMASEP_OD | 1156 | 1054 | 102 | 973 | 81 | 0 |
| **Total** | **1427** | **1325** | **102** | **1183** | **142** | **0** |

- **Constructs observados**: `COUNTIFS` (1393) y `COUNTIF` (34); sumas con `+`;
  rangos de columna entera; criterios de texto con comodines `*`;
  concatenación con `&`; criterios numéricos `< <= > >= =`; criterios por
  referencia de celda de rótulo (`$A26` → etiqueta de texto). 32 patrones de
  fórmula distintos.
- **No soportadas**: 102 celdas en `REMASEP_OD`, todas de la forma
  `=COUNTIFS(…)-F30` (restan una celda de la misma hoja → agregación
  *derivada*, fuera de alcance por diseño). 0 en `REMASEP 01` y `B2 ANEXO`.
- **Cache**: 1183 `MATCH`, 142 `CACHE_DIFFERENCE`, 0 `CACHE_UNAVAILABLE`.
- **Dependencia de `AF`**: 1291 de las 1325 celdas soportadas referencian
  `$AF:$AF`. Las **34 celdas que NO dependen de `AF` coinciden 100 % con la
  cache** (34/34 `MATCH`) — es la evidencia más limpia de equivalencia.
- **Las 142 `CACHE_DIFFERENCE` dependen todas de `AF`** y llevan la nota de
  cache obsoleta. Ejemplo: `REMASEP_OD!E4` Python = 67 vs. cache = 85 (18
  registros con edad cacheada en `0`).
- **Padding-sensitive**: 0.
- `formula_support_status`: **PARTIAL** (por las 102 `…-F30`).
- `cache_consistency_status`: **DIFFERENCES** (por las 142 diferencias de `AF`).

## Artefactos

`scripts/compare_legacy_aggregations.py "<workbook>"` →
`artifacts/legacy_aggregation_equivalence/`:

| Archivo | Contenido |
|---|---|
| `direct_aggregation_inventory.csv` | toda celda de las 3 hojas que referencia la hoja de detalle, con `classification` |
| `unsupported_formulas.csv` | fórmulas fuera del subset + motivo |
| `padding_sensitivity.csv` | por fórmula soportada: `padding_can_affect_result` + razón |
| `aggregation_equivalence.csv` | valor Python vs. cache, `cache_comparison_status`, `supported`, notas |
| `sheet_summary.csv` | conteos por hoja |
| `legacy_metrics_long.csv` | `metric_id` técnico (`LEGACY::<hoja>::<celda>`) → valor agregado |
| `summary.json` | totales, `formula_support_status`, `cache_consistency_status` |
| `README.md` | resumen de la corrida |

### Privacidad

La unidad de comparación es una **celda agregada**. Los artefactos contienen
sólo fórmulas, coordenadas, nombres de columna y conteos/valores agregados.
**Ningún** dato individual de paciente (RUN, nombre, fecha de nacimiento, sexo,
prestación ni tipo de cita por fila). No hay "mismatch en la fila N".

## Limitaciones

- No se recalcula Excel: la comparación es contra el **valor cacheado** del
  workbook, que puede estar obsoleto (confirmado para `AF`).
- Sólo se reproduce el subset de constructs **observado**; un workbook con otras
  fórmulas de conteo marcaría `UNSUPPORTED` en vez de adivinar.
- Las agregaciones derivadas (`…-F30`, subtotales) quedan para un sprint
  posterior sobre el modelo de métricas semántico.
- Equivalencia numérica con este workbook **≠** validación funcional MINSAL.
