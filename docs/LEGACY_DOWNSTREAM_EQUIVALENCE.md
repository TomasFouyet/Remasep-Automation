# Legacy downstream equivalence (Sprint 2.5)

`scripts/evaluate_legacy_downstream.py` + `src/remasep/services/legacy_aggregation.py`
(`parse_sum_formula`, `parse_if_formula`).

Sprint 2.4 cerró las 1427 celdas `BASE_AGGREGATION`/`DERIVED_AGGREGATION` y
descubrió **616 celdas transitivas más** (344 `SUM`, 272 `IF`) sin evaluarlas.
Sprint 2.5 cierra la evaluación de las **2043 celdas Medinet-dependientes
conocidas**, reutilizando el mismo grafo de dependencias (orden topológico,
`depends_on_age`, ciclos, dependencias faltantes) construido en Sprint 2.4 —
sin volver a leer el workbook ni recalcular el cierre transitivo.

> **Aviso.** Reproducir `GENERACION DATOS REMASEP.xlsx` **no** es validación
> oficial MINSAL. Se demuestra que, dado el mismo dataset, la aritmética del
> workbook (`SUM`, `IF` de comprobación) puede recalcularse en Python a partir
> de las mismas celdas agregadas — no que esas comprobaciones midan lo correcto.

## Cierre completo del DAG

```
BASE_AGGREGATION  ->  DERIVED_AGGREGATION  ->  DOWNSTREAM_TOTAL  ->  VALIDATION
      (Sprint 2.3)         (Sprint 2.4)              (Sprint 2.5, SUM/IF)
```

El diagrama es conceptual: la evaluación real sigue el **orden topológico del
grafo** (`ClosureResult.topo_order`, Kahn sobre `graph_deps`), no una secuencia
por fases — un `IF` puede depender de un `SUM` que depende de una `DERIVED`, y
un `SUM` puede sumar una fila que mezcla `BASE` y `DERIVED` en la misma
expresión. El workbook de referencia lo confirma: la fila 46 de `REMASEP_OD`
tiene una celda `DERIVED_AGGREGATION` (`F46`) que un `SUM` de la misma fila sí
suma junto a las demás columnas.

## SUM downstream

Antes de implementar se inventariaron las 344 celdas `SUM` reales
(`downstream_formula_inventory.csv`). Dos formas dominan:

- **rango rectangular de la misma hoja**: `SUM(Q86:Q92)` (225 celdas), a veces
  con una celda suelta a continuación: `SUM(Q86:Q92)+Q84` (19+11+1 celdas);
- **cadena de celdas individuales unidas con `+`**: `SUM(F40+H40+J40+…+AL40)`
  (88 celdas) — la forma dominante en `REMASEP_OD`/`REMASEP 01` para sumar
  columnas alternas (Hombre/Mujer intercaladas).

`parse_sum_formula` soporta ambas, más `+`/`-` de celdas sueltas y literales
numéricos — **nada más** (`*`, `/`, comas dentro de `SUM`, referencias a otra
hoja, resta dentro del `SUM` → `UnsupportedFormulaError`, sin inferir nada).
Los valores se resuelven **sumando los nodos ya evaluados del DAG**: nunca se
relee la cache de Excel como atajo de cálculo. Una celda del rango sin fórmula
(hueco en blanco) cuenta como `0`, igual que en Excel.

## IF validation

Las 272 celdas `IF` reales tienen **4 patrones** (`formula_pattern`, 68 cada
uno): dos de nivel superior (`CELDA < CELDA` con resultado numérico o con
mensaje de texto) y dos con un `IF` anidado (`CELDA <> 0` seguido de
`CELDA_MANUAL = ""`, verificando si el operador llenó un campo del formulario —
ej. *"No olvide digitar el campo Migrantes"*). `parse_if_formula` soporta
exactamente eso: condición `izquierda OP derecha` con `= <> < <= > >=`, ramas
`then`/`else` que son literal numérico, literal de texto, referencia de celda o
un `IF` anidado — cualquier otra forma (más de 3 argumentos, operando con `&`,
resultado con una operación aritmética) es `UNSUPPORTED_IF`.

Algunas condiciones comparan una celda contra un **campo manual del
formulario** (no una fórmula, p.ej. `AO9`, donde el operador escribiría a mano
la cantidad de "Migrantes"). Esa celda no es una métrica Medinet — no hay nada
que "calcular" para ella — así que se lee su valor **crudo** (igual que una
columna de Atenciones), nunca su cache como sustituto de un cálculo pendiente.
La comparación `celda=""` usa la misma semántica de texto legacy
(`normalize_legacy_text`) que el resto del proyecto: una celda vacía (`None`)
cuenta como `""`; un `0` numérico no.

**Las 272 `VALIDATION` siguen clasificadas `VALIDATION`** — se evalúan y se
reportan (`validation_summary.csv`), pero no se interpretan como reglas
clínicas ni se tratan como métricas del negocio.

## Tipo de nodo y `evaluation_status`

`Node` (antes "closure Sprint 2.4") gana un campo `evaluation_status`, común a
las cuatro clasificaciones:

| Valor | Significado |
| --- | --- |
| `SUPPORTED` | se calculó un valor |
| `DEPENDENCY_UNAVAILABLE` | fórmula soportada, pero una dependencia no se pudo resolver (falló, es un ciclo, o es una comparación no numérica/textual reconocible) |
| `UNSUPPORTED_SUM` / `UNSUPPORTED_IF` | la fórmula usa un construct fuera del subset `SUM`/`IF` implementado |
| `UNSUPPORTED_FORMULA` | fuera del subset general (`DOWNSTREAM_TOTAL`/`VALIDATION` con otra función) |
| `CYCLE` | parte de un ciclo de dependencias — nunca se evalúa |

Se mantiene **separado** de `cache_comparison_status` (`MATCH` /
`CACHE_DIFFERENCE` / `CACHE_UNAVAILABLE`): "¿pude calcularlo?" y "¿coincide con
Excel?" son preguntas distintas (Sprint 2.3).

## Cache y `depends_on_AF`

La señal `depends_on_age` se propaga transitivamente por **todo** el grafo
(ya lo hacía Sprint 2.4, incluyendo `DOWNSTREAM_TOTAL`/`VALIDATION`): si una
celda `SUM`/`IF` depende — directa o indirectamente — de una celda que depende
de `AF`, hereda la marca. Ninguna `CACHE_DIFFERENCE` se convierte en `MATCH`
por depender de `AF`; solo se anota.

## Resultado del workbook de referencia

`GENERACION DATOS REMASEP.xlsx` (sha256 `fc2e1536…d15b79`), 1364 atenciones
activas (2006 filas físicas − 642 estructuralmente vacías).

| kind | total | evaluadas | cache MATCH | cache DIFF | cache UNAVAIL |
| --- | ---: | ---: | ---: | ---: | ---: |
| `BASE_AGGREGATION` | 1325 | 1325 | 1183 | 142 | 0 |
| `DERIVED_AGGREGATION` | 102 | 102 | 80 | 22 | 0 |
| `DOWNSTREAM_TOTAL` | 344 | 344 | 260 | 84 | 0 |
| `VALIDATION` | 272 | 272 | 136 | 0 | 136 |
| **Total** | **2043** | **2043** | **1659** | **248** | **136** |

- **`formula_support_status: PASS`** — las 2043 celdas Medinet-dependientes
  conocidas se evaluaron. 0 `UNSUPPORTED`, 0 `DEPENDENCY_UNAVAILABLE`.
- **DAG:** profundidad máxima **4** (cadenas `SUM` de `SUM` de `SUM`), 0
  ciclos, 0 dependencias faltantes.
- **Las 248 `CACHE_DIFFERENCE` dependen todas de `AF`** (142 base + 22
  derivadas + 84 totales que suman celdas con edad) — consistente con Sprint
  2.2/2.3/2.4: es la cache de edad obsoleta, no un error de cálculo.
- **136 `CACHE_UNAVAILABLE`**, todas `VALIDATION`: exactamente la mitad de las
  272 (las de resultado **texto** — mensaje o `""`). Se verificó celda a celda
  que Excel no dejó ningún valor cacheado (`<v>` ausente) para esas 136,
  mientras que las 136 de resultado **numérico** (`0`/`1`) sí lo tienen y
  coinciden al 100 %. No es un fallo del evaluador: es lo que el workbook
  guardó.
- **0 validaciones con valor Python distinto de `0`/`""`**: para este dataset,
  ninguna de las 272 comprobaciones legacy resultaría "activa" (ningún
  registro dispararía el aviso de Migrantes/Embarazadas).
- **Constructs secundarios en las 616 celdas transitivas:** únicamente
  `function:SUM` y `function:IF` (ningún `*`, `/`, `&`, paréntesis de
  agrupación ni otra función) — el subset implementado cubre el 100 % de lo
  observado.

## Qué sigue sin estar soportado

Nada, en el workbook de referencia: las 2043 celdas conocidas son evaluables.
El subset implementado es **el observado**, no "Excel completo": un `SUM` con
comas, resta interna o referencia cruzada de hoja, o un `IF` con más de 3
argumentos, condiciones compuestas (`AND`/`OR`) o resultados calculados, se
reportarían `UNSUPPORTED_SUM`/`UNSUPPORTED_IF` sin inferir nada (cubierto por
tests sintéticos, ver `tests/test_legacy_downstream_parsers.py`).

## Limitaciones

- Sigue dependiendo de la **cache de Excel** para la comparación (no la
  recalcula); confirmada obsoleta para `AF` y ausente para las 136 celdas de
  mensaje.
- El universo evaluado es el **conocido** por el grafo de referencias A1
  (`scripts/formula_refs.py`): sin `INDIRECT`/`OFFSET`/named ranges (no
  presentes en el workbook de referencia).
- Las 272 `VALIDATION` se evalúan como lo que son — comprobaciones técnicas del
  formulario — sin juicio sobre su corrección clínica o funcional.
- Equivalencia numérica con este workbook **≠** validación funcional MINSAL.

## Artefactos

`scripts/evaluate_legacy_downstream.py "<workbook>"` →
`artifacts/legacy_downstream_equivalence/`:

| Archivo | Contenido |
| --- | --- |
| `downstream_formula_inventory.csv` | las 616 celdas transitivas: fórmula, dependencias de la misma hoja, profundidad, patrón, `depends_on_AF` |
| `full_equivalence.csv` | las 2043 celdas: `evaluation_status` y `cache_comparison_status` por separado |
| `sheet_summary.csv` | conteos por hoja y tipo (`base_total/supported`, …) |
| `kind_summary.csv` | conteos por `BASE_AGGREGATION`/`DERIVED_AGGREGATION`/`DOWNSTREAM_TOTAL`/`VALIDATION` |
| `validation_summary.csv` | las celdas `VALIDATION`: valor, cache, `depends_on_AF` — sin interpretación funcional |
| `summary.json` | totales; `formula_support_status` y `cache_consistency_status` separados |
| `README.md` | resumen de la corrida |

### Privacidad

Unidad de comparación = celda agregada. Solo fórmulas, coordenadas, nombres de
columna y valores agregados. Ningún dato individual de paciente (RUN, nombre,
fecha de nacimiento, sexo, prestación ni tipo de cita por fila).
